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
  return { start, end };
}

const { start, end } = getMonthRange(monthStr);

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
const SYSTEM_SHEET_ID = process.env.GOOGLE_SHEET_ID;

const RECIPIENTS = process.env.EMAIL_RECIPIENTS
  ? process.env.EMAIL_RECIPIENTS.split(",")
  : [];

/* --------------------------
   SAFE FETCH
--------------------------- */

async function safeFetchJson(url, headers) {
  const response = await fetch(url, { headers });

  if (response.status === 429) {
    console.log("Rate limited. Waiting 60 seconds...");
    await new Promise(r => setTimeout(r, 60000));
    return safeFetchJson(url, headers);
  }

  const text = await response.text();

  if (!response.ok) {
    throw new Error(`Zendesk API error: ${text}`);
  }

  return JSON.parse(text);
}

/* --------------------------
   LOGGER
--------------------------- */

async function log(message) {
  const timestamp = new Date().toISOString();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SYSTEM_SHEET_ID,
    range: "Logs!A:B",
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[timestamp, message]] }
  });
  console.log(message);
}

/* --------------------------
   MAIN
--------------------------- */

async function run() {

  await log(`Starting export for ${monthStr}`);

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SYSTEM_SHEET_ID,
    range: "Tickets_Raw!A:F"
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SYSTEM_SHEET_ID,
    range: "Tickets_Raw!A1",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[
        "Ticket ID",
        "Created At",
        "Requester Email",
        "Channel",
        "Subject",
        "All Public Comments"
      ]]
    }
  });

  const authHeader = Buffer.from(
    `${process.env.ZENDESK_EMAIL}:${process.env.ZENDESK_API_TOKEN}`
  ).toString("base64");

  let totalSaved = 0;

  for (let d = new Date(start); d < end; d.setDate(d.getDate() + 1)) {

    const dateStr = d.toISOString().split("T")[0];

    const nextDay = new Date(d);
    nextDay.setDate(nextDay.getDate() + 1);
    const nextDateStr = nextDay.toISOString().split("T")[0];

    await log(`Processing day: ${dateStr}`);

    let url =
      `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/search.json` +
      `?query=type:ticket via.channel:messaging created>=${dateStr} created<${nextDateStr}` +
      `&sort_by=created_at&sort_order=asc`;

    while (url) {

      const data = await safeFetchJson(url, {
        Authorization: `Basic ${authHeader}`
      });

      const tickets = data.results || [];
      if (tickets.length === 0) break;

      let rows = [];

      for (const ticket of tickets) {

        let requesterEmail = "N/A";

        if (ticket.requester_id) {
          try {
            const userData = await safeFetchJson(
              `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/users/${ticket.requester_id}.json`,
              { Authorization: `Basic ${authHeader}` }
            );
            requesterEmail = userData.user?.email || "N/A";
          } catch {
            await log(`User lookup failed for ${ticket.id}`);
          }
        }

        let publicComments = "";

        try {
          const commentsData = await safeFetchJson(
            `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/tickets/${ticket.id}/comments.json`,
            { Authorization: `Basic ${authHeader}` }
          );

          publicComments = (commentsData.comments || [])
            .filter(c => c.public)
            .map(c => {
              const role =
                c.author_id === ticket.requester_id
                  ? "Requester:"
                  : "Agent:";
              return `${role} ${c.body}`;
            })
            .join("\n\n---\n\n");

        } catch {
          await log(`Comment fetch failed for ticket ${ticket.id}`);
          publicComments = "Comment retrieval failed.";
        }

        const MAX_CELL_LENGTH = 49000;

        if (publicComments.length > MAX_CELL_LENGTH) {

          const parts = Math.ceil(publicComments.length / MAX_CELL_LENGTH);

          for (let i = 0; i < parts; i++) {
            rows.push([
              ticket.id,
              ticket.created_at,
              requesterEmail,
              ticket.via?.channel || "",
              ticket.subject || "",
              `Part ${i + 1}/${parts}\n\n${publicComments.substring(
                i * MAX_CELL_LENGTH,
                (i + 1) * MAX_CELL_LENGTH
              )}`
            ]);
          }

        } else {

          rows.push([
            ticket.id,
            ticket.created_at,
            requesterEmail,
            ticket.via?.channel || "",
            ticket.subject || "",
            publicComments
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
        await log(`Saved ${rows.length}. Total so far: ${totalSaved}`);
      }

      url = data.next_page;
    }
  }

  await log(`Month complete. Total rows: ${totalSaved}`);

  /* --------------------------
     EXPORT AS EXCEL + EMAIL
  --------------------------- */

  const accessToken = await auth.getAccessToken();

  const exportUrl =
    `https://www.googleapis.com/drive/v3/files/${SYSTEM_SHEET_ID}/export` +
    `?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;

  const exportResponse = await fetch(exportUrl, {
    headers: {
      Authorization: `Bearer ${accessToken.token}`
    }
  });

  const fileBuffer = await exportResponse.buffer();

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
    text: `Attached is the Zendesk monthly export.`,
    attachments: [
      {
        filename: `zendesk_herohq_${monthStr.replace("-", "_")}.xlsx`,
        content: fileBuffer
      }
    ]
  });

  await log("Export emailed successfully.");
}

run().catch(async err => {
  console.error("ERROR:", err.message);
  process.exit(1);
});
