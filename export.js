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
    const retry = parseInt(response.headers.get("retry-after") || "60");
    await new Promise(r => setTimeout(r, retry * 1000));
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
        "Created at",
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

  const EXCLUDED_CHANNELS = ["messaging", "native_messaging"];

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

      console.log(`Processing ${tickets.length} tickets`);

      let rows = [];

      await Promise.all(
        tickets.map(async (ticket) => {

          const channel = ticket.via?.channel || "";
          if (EXCLUDED_CHANNELS.includes(channel)) return;

          const ticket_id = ticket.id;
          const created = ticket.created_at;
          const subject = ticket.subject || "";

          const requester_email =
            ticket.via?.source?.from?.address ||
            ticket.requester_id ||
            "N/A";

          let publicComments = [];
          try {
            const commentsData = await safeFetchJson(
              `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/tickets/${ticket_id}/comments.json`,
              { Authorization: `Basic ${authHeader}` }
            );
            publicComments = (commentsData.comments || []).filter(c => c.public);
          } catch {}

          const formatted = publicComments.map(c => {
            const role =
              c.author_id === ticket.requester_id
                ? "**Requester:**"
                : "**Agent:**";
            return `${role} ${c.body}`;
          });

          const combined = formatted.join("\n\n---\n\n");

          const safeCombined =
            combined && combined.length > 49000
              ? combined.slice(0, 49000) + "\n\n[Truncated]"
              : combined;

          rows.push([
            ticket_id,
            created,
            requester_email,
            channel,
            subject,
            safeCombined
          ]);
        })
      );

      if (rows.length > 0) {
        await sheets.spreadsheets.values.append({
          spreadsheetId: SYSTEM_SHEET_ID,
          range: "Tickets_Raw!A:F",
          valueInputOption: "USER_ENTERED",
          requestBody: { values: rows }
        });

        totalSaved += rows.length;
        console.log(`Saved ${rows.length}, total: ${totalSaved}`);
      }

      url = data.next_page;
    }
  }

  console.log(`Month complete. Total rows: ${totalSaved}`);

  /* --------------------------
     CLEAN EXPORT
  --------------------------- */

  const temp = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: `zendesk_herohq_${monthStr.replace("-", "_")}` }
    }
  });

  const tempId = temp.data.spreadsheetId;

  const source = await sheets.spreadsheets.values.get({
    spreadsheetId: SYSTEM_SHEET_ID,
    range: "Tickets_Raw!A:F"
  });

  const values = source.data.values || [];

  const cleanedValues = values.filter((row, index) => {
    if (index === 0) return true;

    const [id, created, email, channel, subject, comments] = row;

    const isEmpty =
      !id && !created && !email && !channel && !subject && !comments;

    return !isEmpty;
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: tempId,
    range: "Sheet1!A1",
    valueInputOption: "USER_ENTERED",
    requestBody: { values: cleanedValues }
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: tempId,
    requestBody: {
      requests: [{
        updateSheetProperties: {
          properties: { sheetId: 0, title: monthStr.replace("-", "_") },
          fields: "title"
        }
      }]
    }
  });

  /* EXPORT */

  const client = await auth.getClient();

  const res = await client.request({
    url: `https://docs.google.com/spreadsheets/d/${tempId}/export`,
    params: { format: "xlsx" },
    responseType: "arraybuffer"
  });

  const fileBuffer = Buffer.from(res.data);

  /* EMAIL */

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
    text: "Attached is this month’s Zendesk export.",
    attachments: [{
      filename: `zendesk_herohq_${monthStr.replace("-", "_")}.xlsx`,
      content: fileBuffer
    }]
  });

  console.log("Export emailed successfully.");
}

run().catch(err => {
  console.error("FATAL ERROR:", err.message);
  process.exit(1);
});
