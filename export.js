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
    console.log("Rate limited. Waiting 10 seconds...");
    await new Promise(r => setTimeout(r, 10000));
    return safeFetchJson(url, headers);
  }

  const text = await response.text();

  if (!response.ok) {
    throw new Error(`Zendesk API error: ${text}`);
  }

  return JSON.parse(text);
}

/* --------------------------
   MAIN
--------------------------- */

async function run() {

  console.log(`Starting export for ${monthStr}`);

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SYSTEM_SHEET_ID,
    range: "Tickets_Raw!A:G"
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SYSTEM_SHEET_ID,
    range: "Tickets_Raw!A1",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [[
        "Ticket ID",
        "Created At",
        "Requester ID",
        "Channel",
        "Subject",
        "Comment Author Type",
        "Comment Body"
      ]]
    }
  });

  const authHeader = Buffer.from(
    `${process.env.ZENDESK_EMAIL}:${process.env.ZENDESK_API_TOKEN}`
  ).toString("base64");

  let totalSaved = 0;
  const MAX_CELL = 48000;

  for (let d = new Date(start); d < end; d.setDate(d.getDate() + 1)) {

    const dateStr = d.toISOString().split("T")[0];
    const nextDay = new Date(d);
    nextDay.setDate(nextDay.getDate() + 1);
    const nextDateStr = nextDay.toISOString().split("T")[0];

    console.log("Processing day:", dateStr);

    let url =
      `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/search.json` +
      `?query=type:ticket created>=${dateStr} created<${nextDateStr}` +
      `&sort_by=created_at&sort_order=asc`;

    while (url) {

      const data = await safeFetchJson(url, {
        Authorization: `Basic ${authHeader}`
      });

      const tickets = data.results || [];
      if (tickets.length === 0) break;

      let rows = [];
      const BATCH_SIZE = 5;

      for (let i = 0; i < tickets.length; i += BATCH_SIZE) {

        const batch = tickets.slice(i, i + BATCH_SIZE);

        const results = await Promise.all(
          batch.map(async (ticket) => {

            try {

              const commentsData = await safeFetchJson(
                `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/tickets/${ticket.id}/comments.json`,
                { Authorization: `Basic ${authHeader}` }
              );

              const publicComments = (commentsData.comments || [])
                .filter(c => c.public);

              let expandedRows = [];

              for (const c of publicComments) {

                const authorType =
                  c.author_id === ticket.requester_id
                    ? "Requester"
                    : "Agent";

                const body = c.body || "";

                if (body.length <= MAX_CELL) {

                  expandedRows.push([
                    ticket.id,
                    ticket.created_at,
                    ticket.requester_id || "",
                    ticket.via?.channel || "",
                    ticket.subject || "",
                    authorType,
                    body
                  ]);

                } else {

                  const parts = Math.ceil(body.length / MAX_CELL);

                  for (let p = 0; p < parts; p++) {

                    const chunk = body.substring(
                      p * MAX_CELL,
                      (p + 1) * MAX_CELL
                    );

                    expandedRows.push([
                      ticket.id,
                      ticket.created_at,
                      ticket.requester_id || "",
                      ticket.via?.channel || "",
                      ticket.subject || "",
                      `${authorType} (Part ${p + 1}/${parts})`,
                      chunk
                    ]);
                  }
                }
              }

              return expandedRows;

            } catch {

              return [[
                ticket.id,
                ticket.created_at,
                ticket.requester_id || "",
                ticket.via?.channel || "",
                ticket.subject || "",
                "System",
                "Comment retrieval failed."
              ]];
            }
          })
        );

        rows.push(...results.flat());

        await new Promise(r => setTimeout(r, 1000));
      }

      if (rows.length > 0) {

        await sheets.spreadsheets.values.append({
          spreadsheetId: SYSTEM_SHEET_ID,
          range: "Tickets_Raw!A:G",
          valueInputOption: "USER_ENTERED",
          requestBody: { values: rows }
        });

        totalSaved += rows.length;
        console.log(`Saved ${rows.length}. Total so far: ${totalSaved}`);
      }

      url = data.next_page;
    }
  }

  console.log(`Month complete. Total rows: ${totalSaved}`);

  /* --------------------------
     EXPORT EXCEL
  --------------------------- */

  const client = await auth.getClient();
  const accessTokenObj = await client.getAccessToken();
  const accessToken = accessTokenObj.token;

  const exportUrl =
    `https://www.googleapis.com/drive/v3/files/${SYSTEM_SHEET_ID}/export` +
    `?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;

  const exportResponse = await fetch(exportUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  if (!exportResponse.ok) {
    const errorText = await exportResponse.text();
    console.error("Drive export error:", errorText);
    throw new Error("Drive export failed");
  }

  const arrayBuffer = await exportResponse.arrayBuffer();
  const fileBuffer = Buffer.from(arrayBuffer);

  console.log("Export file size:", fileBuffer.length);

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

  console.log("Export emailed successfully.");
}

run().catch(err => {
  console.error("FATAL ERROR:", err.message);
  process.exit(1);
});
