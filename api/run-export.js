import { google } from "googleapis";
import nodemailer from "nodemailer";

/* =========================
   GOOGLE AUTH
========================= */

function getGoogleAuth() {
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  return new google.auth.GoogleAuth({
    credentials: creds,
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive"
    ]
  });
}

/* =========================
   HELPERS
========================= */

function getPreviousMonth() {
  const now = new Date();
  const firstDayThisMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastMonth = new Date(firstDayThisMonth - 1);
  const year = lastMonth.getFullYear();
  const month = String(lastMonth.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

function getMonthRange(monthStr) {
  const [year, month] = monthStr.split("-");
  const start = new Date(`${year}-${month}-01T00:00:00Z`);
  const end = new Date(start);
  end.setMonth(start.getMonth() + 1);
  return {
    startTimestamp: Math.floor(start.getTime() / 1000),
    endTimestamp: Math.floor(end.getTime() / 1000)
  };
}

/* =========================
   MAIN HANDLER
========================= */

export default async function handler(req, res) {
  try {
    const auth = await getGoogleAuth();
    const sheets = google.sheets({ version: "v4", auth });
    const drive = google.drive({ version: "v3", auth });

    const SYSTEM_SHEET_ID = process.env.GOOGLE_SHEET_ID;
    const RECIPIENTS = process.env.EMAIL_RECIPIENTS.split(",");

    const monthParam = req.query.month || null;
    const monthStr = monthParam || getPreviousMonth();
    const { startTimestamp, endTimestamp } = getMonthRange(monthStr);

    /* =========================
       READ STATE
    ========================= */

    const stateRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SYSTEM_SHEET_ID,
      range: "State!A:B"
    });

    let checkpoint = startTimestamp;
    if (stateRes.data.values) {
      const rows = stateRes.data.values;
      rows.forEach(row => {
        if (row[0] === "checkpoint") checkpoint = parseInt(row[1]);
      });
    }

    /* =========================
       FETCH INCREMENTAL
    ========================= */

    const subdomain = process.env.ZENDESK_SUBDOMAIN;
    const email = process.env.ZENDESK_EMAIL;
    const token = process.env.ZENDESK_API_TOKEN;

    const authHeader = Buffer.from(`${email}:${token}`).toString("base64");

    const response = await fetch(
      `https://${subdomain}.zendesk.com/api/v2/incremental/tickets.json?start_time=${checkpoint}`,
      { headers: { Authorization: `Basic ${authHeader}` } }
    );

    const data = await response.json();
    const tickets = data.tickets || [];
    const newCheckpoint = data.end_time;

    let rowsToAppend = [];

    for (const ticket of tickets) {
      const createdTimestamp = Math.floor(
        new Date(ticket.created_at).getTime() / 1000
      );

      if (createdTimestamp >= startTimestamp && createdTimestamp < endTimestamp) {

        // Fetch requester email
        let requesterEmail = "N/A";
        if (ticket.requester_id) {
          const userRes = await fetch(
            `https://${subdomain}.zendesk.com/api/v2/users/${ticket.requester_id}.json`,
            { headers: { Authorization: `Basic ${authHeader}` } }
          );
          const userData = await userRes.json();
          requesterEmail = userData.user?.email || "N/A";
        }

        // Fetch comments
        const commentsRes = await fetch(
          `https://${subdomain}.zendesk.com/api/v2/tickets/${ticket.id}/comments.json`,
          { headers: { Authorization: `Basic ${authHeader}` } }
        );
        const commentsData = await commentsRes.json();
        const publicComments = (commentsData.comments || []).filter(c => c.public);

        const formattedComments = publicComments
          .map(c => {
            const role = c.author_id === ticket.requester_id
              ? "**Requester:**"
              : "**Agent:**";
            return `${role} ${c.body}`;
          })
          .join("\n\n---\n\n");

        rowsToAppend.push([
          ticket.id,
          ticket.created_at,
          requesterEmail,
          ticket.via?.channel || "",
          ticket.subject || "",
          formattedComments
        ]);
      }
    }

    /* =========================
       APPEND RAW DATA
    ========================= */

    if (rowsToAppend.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SYSTEM_SHEET_ID,
        range: "Tickets_Raw!A:F",
        valueInputOption: "USER_ENTERED",
        requestBody: { values: rowsToAppend }
      });
    }

    /* =========================
       UPDATE STATE
    ========================= */

    await sheets.spreadsheets.values.update({
      spreadsheetId: SYSTEM_SHEET_ID,
      range: "State!A2:B2",
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [["checkpoint", newCheckpoint]] }
    });

    /* =========================
       CHECK COMPLETION
    ========================= */

    if (newCheckpoint >= endTimestamp) {

      // Create new workbook
      const file = await drive.files.create({
        requestBody: {
          name: `Zendesk Export - ${monthStr}`,
          mimeType: "application/vnd.google-apps.spreadsheet"
        }
      });

      const exportSheetId = file.data.id;

      // Read raw data
      const rawData = await sheets.spreadsheets.values.get({
        spreadsheetId: SYSTEM_SHEET_ID,
        range: "Tickets_Raw!A:F"
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId: exportSheetId,
        range: "Export!A1",
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [
            ["Ticket ID","Created At","Requester Email","Channel","Subject","All Public Comments"],
            ...(rawData.data.values || [])
          ]
        }
      });

      // Share
      for (const emailAddr of RECIPIENTS) {
        await drive.permissions.create({
          fileId: exportSheetId,
          requestBody: {
            role: "writer",
            type: "user",
            emailAddress: emailAddr
          }
        });
      }

      // Email notification
      const transporter = nodemailer.createTransport({
        service: "gmail",
        auth: {
          user: process.env.GMAIL_SENDER,
          pass: process.env.GMAIL_APP_PASSWORD
        }
      });

      await transporter.sendMail({
        from: process.env.GMAIL_SENDER,
        to: RECIPIENTS,
        subject: `Zendesk Monthly Report - ${monthStr}`,
        text: `Export completed.\n\nGoogle Sheet:\nhttps://docs.google.com/spreadsheets/d/${exportSheetId}`
      });

      // Clear state
      await sheets.spreadsheets.values.clear({
        spreadsheetId: SYSTEM_SHEET_ID,
        range: "State!A2:B2"
      });

      await sheets.spreadsheets.values.clear({
        spreadsheetId: SYSTEM_SHEET_ID,
        range: "Tickets_Raw!A:F"
      });

      return res.status(200).json({ message: "Export complete" });
    }

    // Auto-chain
    await fetch(`${process.env.VERCEL_URL ? "https://" + process.env.VERCEL_URL : ""}/api/run-export?month=${monthStr}`);

    return res.status(200).json({ message: "Processing continued..." });

  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
