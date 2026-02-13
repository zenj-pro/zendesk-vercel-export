import fs from "fs";
import path from "path";
import XLSX from "xlsx";
import nodemailer from "nodemailer";

export default async function handler(req, res) {
  try {
    const subdomain = process.env.ZENDESK_SUBDOMAIN;
    const email = process.env.ZENDESK_EMAIL;
    const apiToken = process.env.ZENDESK_API_TOKEN;
    const gmailUser = process.env.GMAIL_SENDER;
    const gmailPass = process.env.GMAIL_APP_PASSWORD;
    const recipients = process.env.EMAIL_RECIPIENTS;

    if (!subdomain || !email || !apiToken) {
      return res.status(500).json({ error: "Missing Zendesk credentials" });
    }

    const auth = Buffer.from(`${email}:${apiToken}`).toString("base64");

    const checkpointPath = "/tmp/checkpoint.json";
    const ticketsPath = "/tmp/tickets.json";
    const excelPath = "/tmp/december_export.xlsx";

    let startTime;

    if (fs.existsSync(checkpointPath)) {
      const saved = JSON.parse(fs.readFileSync(checkpointPath, "utf8"));
      startTime = saved.end_time;
    } else {
      // December 1 2025 UTC
      startTime = Math.floor(new Date("2025-12-01T00:00:00Z").getTime() / 1000);
    }

    const url = `https://${subdomain}.zendesk.com/api/v2/incremental/tickets.json?start_time=${startTime}`;

    const response = await fetch(url, {
      headers: {
        Authorization: `Basic ${auth}`
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      return res.status(response.status).json({ error: errorText });
    }

    const data = await response.json();

    const tickets = data.tickets || [];
    const endTime = data.end_time;
    const endOfStream = data.end_of_stream;

    // Load existing tickets
    let existing = [];
    if (fs.existsSync(ticketsPath)) {
      existing = JSON.parse(fs.readFileSync(ticketsPath, "utf8"));
    }

    const combined = existing.concat(tickets);
    fs.writeFileSync(ticketsPath, JSON.stringify(combined));
    fs.writeFileSync(checkpointPath, JSON.stringify({ end_time: endTime }));

    console.log(`Processed ${tickets.length} tickets this call`);
    console.log(`Total accumulated: ${combined.length}`);

    if (!endOfStream) {
      return res.status(200).json({
        processed_this_call: tickets.length,
        total_collected: combined.length,
        completed: false
      });
    }

    // --------------------------
    // EXPORT TO EXCEL
    // --------------------------

    const formatted = combined.map(ticket => ({
      "Ticket ID": ticket.id,
      "Created At": ticket.created_at,
      "Requester Email": ticket.requester_id,
      "Channel": ticket.via?.channel || "",
      "Subject": ticket.subject || "",
      "Status": ticket.status
    }));

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(formatted);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Tickets");
    XLSX.writeFile(workbook, excelPath);

    console.log("Excel generated");

    // --------------------------
    // SEND EMAIL
    // --------------------------

    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: gmailUser,
        pass: gmailPass
      }
    });

    await transporter.sendMail({
      from: gmailUser,
      to: recipients,
      subject: "Zendesk December 2025 Export",
      text: "Attached is the December Zendesk export.",
      attachments: [
        {
          filename: "december_export.xlsx",
          path: excelPath
        }
      ]
    });

    console.log("Email sent successfully");

    // Cleanup after completion
    fs.unlinkSync(checkpointPath);
    fs.unlinkSync(ticketsPath);
    fs.unlinkSync(excelPath);

    return res.status(200).json({
      message: "Export complete and emailed",
      total_tickets: combined.length,
      completed: true
    });

  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}
