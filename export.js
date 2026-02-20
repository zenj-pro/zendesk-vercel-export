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
    requestBody: { values: [[timestamp, message]] }
  });

  console.log(message);
}

/* --------------------------
   MAIN
--------------------------- */

async function run() {

  await log(`Starting export for ${monthStr}`);

  /* --------------------------
     CLEAR SHEET + HEADER
  --------------------------- */

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SYSTEM_SHEET_ID,
    range: "Tickets_Raw!A:F"
  });

  await log("Cleared Tickets_Raw sheet.");

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

  await log("Inserted header row.");

  const authHeader = Buffer.from(
    `${process.env.ZENDESK_EMAIL}:${process.env.ZENDESK_API_TOKEN}`
  ).toString("base64");

  let totalSaved = 0;

  for (let d = new Date(start); d < end; d.setDate(d.getDate() + 1)) {

    const day = new Date(d);
    const dateStr = day.toISOString().split("T")[0];

    await log(`Processing day: ${dateStr}`);

    let url =
      `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/search.json` +
      `?query=type:ticket created>=${dateStr} created<=${dateStr}` +
      `&sort_by=created_at&sort_order=asc`;

    while (url) {

      const response = await fetch(url, {
        headers: { Authorization: `Basic ${authHeader}` }
      });

      if (response.status === 429) {
        await log("Rate limited. Waiting 60 seconds...");
        await new Promise(r => setTimeout(r, 60000));
        continue;
      }

      if (!response.ok) {
        const text = await response.text();
        throw new Error(text);
      }

      const data = await response.json();
      const tickets = data.results || [];

      if (tickets.length === 0) break;

      let rows = [];

      for (const ticket of tickets) {

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
          { headers: { Authorization: `Basic ${authHeader}` }
        });

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

        rows.push([
          ticket.id,
          ticket.created_at,
          requesterEmail,
          ticket.via?.channel || "",
          ticket.subject || "",
          publicComments
        ]);
      }

      if (rows.length > 0) {
        await sheets.spreadsheets.values.append({
          spreadsheetId: SYSTEM_SHEET_ID,
          range: "Tickets_Raw!A:F",
          valueInputOption: "USER_ENTERED",
          requestBody: { values: rows }
        });

        totalSaved += rows.length;
        await log(`Saved ${rows.length} tickets for ${dateStr}. Total so far: ${totalSaved}`);
      }

      url = data.next_page;
    }
  }

  await log(`Finished month. Total saved: ${totalSaved}`);

  await log("Export complete.");
}

run().catch(async err => {
  console.error("ERROR:", err.message);
  await log(`ERROR: ${err.message}`);
  process.exit(1);
});
