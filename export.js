import fetch from "node-fetch";
import { google } from "googleapis";
import nodemailer from "nodemailer";

/* --------------------------
   MONTH HANDLING
--------------------------- */

function getPreviousMonth() {
  const now = new Date();
  now.setMonth(now.getMonth() - 1);
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

const monthStr = process.env.EXPORT_MONTH || getPreviousMonth();
console.log("Export month:", monthStr);

function getMonthRange(monthStr) {
  const [year, month] = monthStr.split("-");
  const start = new Date(`${year}-${month}-01T00:00:00Z`);
  const end = new Date(start);
  end.setMonth(start.getMonth() + 1);
  end.setSeconds(end.getSeconds() - 1);

  return {
    startDate: start.toISOString().split("T")[0],
    endDate: end.toISOString().split("T")[0]
  };
}

const { startDate, endDate } = getMonthRange(monthStr);

/* --------------------------
   GOOGLE AUTH
--------------------------- */

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

/* --------------------------
   LOGGER
--------------------------- */

async function log(message) {
  const timestamp = new Date().toISOString();

  await sheets.spreadsheets.values.append({
    spreadsheetId: SYSTEM_SHEET_ID,
    range: "Logs!A:B",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[timestamp, message]]
    }
  });

  console.log(message);
}

/* --------------------------
   MAIN
--------------------------- */

async function run() {

  await log(`Starting export for ${monthStr}`);
  await log(`Date range: ${startDate} to ${endDate}`);

  const authHeader = Buffer.from(
    `${process.env.ZENDESK_EMAIL}:${process.env.ZENDESK_API_TOKEN}`
  ).toString("base64");

  let url =
    `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/search.json` +
    `?query=type:ticket created>=${startDate} created<=${endDate}` +
    `&sort_by=created_at&sort_order=asc`;

  let allRows = [];
  let totalFetched = 0;

  while (url) {
    const response = await fetch(url, {
      headers: { Authorization: `Basic ${authHeader}` }
    });

    if (response.status === 429) {
      await log("Rate limited. Waiting 60s...");
      await new Promise(r => setTimeout(r, 60000));
      continue;
    }

    if (!response.ok) {
      const text = await response.text();
      throw new Error(text);
    }

    const data = await response.json();
    const tickets = data.results || [];

    totalFetched += tickets.length;
    await log(`Fetched page. Total so far: ${totalFetched}`);

    for (const ticket of tickets) {

      const requesterRes = await fetch(
        `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/users/${ticket.requester_id}.json`,
        { headers: { Authorization: `Basic ${authHeader}` } }
      );

      const requesterData = await requesterRes.json();
      const requesterEmail = requesterData.user?.email || "N/A";

      const commentsRes = await fetch(
        `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/tickets/${ticket.id}/comments.json`,
        { headers: { Authorization: `Basic ${authHeader}` } }
      );

      const commentsData = await commentsRes.json();

      const publicComments = (commentsData.comments || [])
        .filter(c => c.public)
        .map(c => {
          const role =
            c.author_id === ticket.requester_id
              ? "**Requester:**"
              : "**Agent:**";
          return `${role} ${c.body}`;
        })
        .join("\n\n---\n\n");

      allRows.push([
        ticket.id,
        ticket.created_at,
        requesterEmail,
        ticket.via?.channel || "",
        ticket.subject || "",
        publicComments
      ]);
    }

    url = data.next_page;
  }

  await log(`Finished fetching. Total tickets: ${allRows.length}`);

  /* --------------------------
     CREATE EXPORT WORKBOOK
  --------------------------- */

  const file = await drive.files.create({
    requestBody: {
      name: `Zendesk Export - ${monthStr}`,
      mimeType: "application/vnd.google-apps.spreadsheet"
    }
  });

  const exportId = file.data.id;

  await sheets.spreadsheets.values.update({
    spreadsheetId: exportId,
    range: "A1",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [
        [
          "Ticket ID",
          "Created At",
          "Requester Email",
          "Channel",
          "Subject",
          "All Public Comments"
        ],
        ...allRows
      ]
    }
  });

  await log("Export workbook created.");

  /* --------------------------
     EMAIL
  --------------------------- */

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

    await log("Email sent.");
  }

  await log("Export complete.");
}

run().catch(async err => {
  console.error("ERROR:", err.message);
  await log(`ERROR: ${err.message}`);
  process.exit(1);
});
