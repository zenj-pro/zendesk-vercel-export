import fetch from "node-fetch";
import { google } from "googleapis";
import nodemailer from "nodemailer";

const monthStr = process.env.EXPORT_MONTH;

if (!monthStr) {
  console.log("No EXPORT_MONTH provided. Exiting.");
  process.exit(0);
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

const { startTimestamp, endTimestamp } = getMonthRange(monthStr);

const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);

const auth = new google.auth.GoogleAuth({
  credentials: creds,
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
  ]
});

const sheets = google.sheets({ version: "v4", auth });
const drive = google.drive({ version: "v3", auth });

const SYSTEM_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const RECIPIENTS = process.env.EMAIL_RECIPIENTS
  ? process.env.EMAIL_RECIPIENTS.split(",")
  : [];

async function log(status, checkpoint, fetched, saved, lastTicket) {
  const timestamp = new Date().toISOString();

  await sheets.spreadsheets.values.append({
    spreadsheetId: SYSTEM_SHEET_ID,
    range: "Logs!A:G",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[
        timestamp,
        monthStr,
        checkpoint,
        fetched,
        saved,
        lastTicket,
        status
      ]]
    }
  });
}

async function run() {
  console.log("Starting batch for month:", monthStr);

  const stateRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SYSTEM_SHEET_ID,
    range: "State!A:B"
  });

  let checkpoint = startTimestamp;
  let isFreshRun = true;

  if (stateRes.data.values) {
    stateRes.data.values.forEach(row => {
      if (row[0] === "checkpoint" && row[1]) {
        checkpoint = parseInt(row[1]);
        isFreshRun = false;
      }
    });
  }

  if (isFreshRun) {
    console.log("Fresh run detected. Clearing Logs and Tickets_Raw.");

    await sheets.spreadsheets.values.clear({
      spreadsheetId: SYSTEM_SHEET_ID,
      range: "Logs!A:G"
    });

    await sheets.spreadsheets.values.clear({
      spreadsheetId: SYSTEM_SHEET_ID,
      range: "Tickets_Raw!A:F"
    });
  }

  console.log("Current checkpoint:", checkpoint);

  const authHeader = Buffer.from(
    `${process.env.ZENDESK_EMAIL}:${process.env.ZENDESK_API_TOKEN}`
  ).toString("base64");

  let nextUrl = `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/incremental/tickets.json?start_time=${checkpoint}`;

  let totalFetched = 0;
  let totalSaved = 0;
  let lastTicketID = null;

  while (nextUrl) {
    const response = await fetch(nextUrl, {
      headers: { Authorization: `Basic ${authHeader}` }
    });

    const data = await response.json();
    const tickets = data.tickets || [];
    const newCheckpoint = data.end_time;

    totalFetched += tickets.length;

    let rows = [];

    for (const ticket of tickets) {
      const createdTs = Math.floor(new Date(ticket.created_at).getTime() / 1000);

      if (createdTs >= startTimestamp && createdTs < endTimestamp) {

        lastTicketID = ticket.id;

        let requesterEmail = "N/A";
        if (ticket.requester_id) {
          const userRes = await fetch(
            `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/users/${ticket.requester_id}.json`,
            { headers: { Authorization: `Basic ${authHeader}` } }
          );
          const userData = await userRes.json();
          requesterEmail = userData.user?.email || "N/A";
        }

        const commentsRes = await fetch(
          `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/tickets/${ticket.id}/comments.json`,
          { headers: { Authorization: `Basic ${authHeader}` } }
        );
        const commentsData = await commentsRes.json();

        const publicComments = (commentsData.comments || []).filter(c => c.public);

        const formattedComments = publicComments.map(c => {
          const role = c.author_id === ticket.requester_id
            ? "**Requester:**"
            : "**Agent:**";
          return `${role} ${c.body}`;
        }).join("\n\n---\n\n");

        rows.push([
          ticket.id,
          ticket.created_at,
          requesterEmail,
          ticket.via?.channel || "",
          ticket.subject || "",
          formattedComments
        ]);
      }
    }

    if (rows.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SYSTEM_SHEET_ID,
        range: "Tickets_Raw!A:F",
        valueInputOption: "USER_ENTERED",
        requestBody: { values: rows }
      });

      totalSaved += rows.length;
    }

    checkpoint = newCheckpoint;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SYSTEM_SHEET_ID,
      range: "State!A2:B2",
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [["checkpoint", checkpoint]] }
    });

    await log("Running", checkpoint, totalFetched, totalSaved, lastTicketID);

    if (data.end_of_stream || checkpoint >= endTimestamp) {
      break;
    }

    nextUrl = data.next_page;
  }

  if (checkpoint >= endTimestamp) {
    console.log("Month export complete. Creating workbook.");

    const file = await drive.files.create({
      requestBody: {
        name: `Zendesk Export - ${monthStr}`,
        mimeType: "application/vnd.google-apps.spreadsheet"
      }
    });

    const exportId = file.data.id;

    const rawData = await sheets.spreadsheets.values.get({
      spreadsheetId: SYSTEM_SHEET_ID,
      range: "Tickets_Raw!A:F"
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: exportId,
      range: "A1",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [
          ["Ticket ID","Created At","Requester Email","Channel","Subject","All Public Comments"],
          ...(rawData.data.values || [])
        ]
      }
    });

    for (const email of RECIPIENTS) {
      await drive.permissions.create({
        fileId: exportId,
        requestBody: {
          role: "writer",
          type: "user",
          emailAddress: email
        }
      });
    }

    if (RECIPIENTS.length > 0) {
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
        text: `Export complete:\nhttps://docs.google.com/spreadsheets/d/${exportId}`
      });
    }

    await sheets.spreadsheets.values.clear({
      spreadsheetId: SYSTEM_SHEET_ID,
      range: "State!A2:B2"
    });

    await sheets.spreadsheets.values.clear({
      spreadsheetId: SYSTEM_SHEET_ID,
      range: "Tickets_Raw!A:F"
    });

    await log("Export Complete", "FINAL", totalFetched, totalSaved, "DONE");
  }
}

run().catch(async err => {
  console.error("ERROR:", err.message);
  await log(`ERROR: ${err.message}`, "UNKNOWN", 0, 0, "FAIL");
  process.exit(1);
});
