import fetch from "node-fetch";
import nodemailer from "nodemailer";
import XLSX from "xlsx";

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
   SAFE FETCH
--------------------------- */

async function safeFetchJson(url, headers) {
  const res = await fetch(url, { headers });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(text);
  }

  return res.json();
}

/* --------------------------
   MAIN
--------------------------- */

async function run() {

  const authHeader = Buffer.from(
    `${process.env.ZENDESK_EMAIL}:${process.env.ZENDESK_API_TOKEN}`
  ).toString("base64");

  const headers = {
    Authorization: `Basic ${authHeader}`
  };

  const EXCLUDED_CHANNELS = ["messaging", "native_messaging"];

  let ticketMap = {}; // ticket_id → data

  /* --------------------------
     FETCH EVENTS IN BULK
  --------------------------- */

  let startTime = Math.floor(start.getTime() / 1000);

  let url = `https://${process.env.ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/incremental/ticket_events.json?start_time=${startTime}`;

  while (url) {

    const data = await safeFetchJson(url, headers);

    const events = data.ticket_events || [];

    console.log(`Fetched ${events.length} events`);

    for (const event of events) {

      const ticket_id = event.ticket_id;
      if (!ticket_id) continue;

      // Initialize ticket
      if (!ticketMap[ticket_id]) {
        ticketMap[ticket_id] = {
          ticket_id,
          created_at: null,
          subject: "",
          channel: "",
          requester_email: "",
          comments: []
        };
      }

      const ticket = ticketMap[ticket_id];

      // Ticket creation event
      if (event.event_type === "Create") {
        ticket.created_at = event.created_at;
        ticket.subject = event.ticket?.subject || "";
        ticket.channel = event.ticket?.via?.channel || "";
        ticket.requester_email =
          event.ticket?.requester?.email ||
          event.ticket?.via?.source?.from?.address ||
          "N/A";
      }

      // Comment event
      if (event.event_type === "Comment" && event.public) {

        const role =
          event.author_id === event.ticket?.requester_id
            ? "**Requester:**"
            : "**Agent:**";

        const body = event.body || "";

        ticket.comments.push(`${role} ${body}`);
      }
    }

    // Stop when beyond end date
    const lastEvent = events[events.length - 1];
    if (!lastEvent || new Date(lastEvent.created_at) > end) break;

    url = data.next_page;
  }

  /* --------------------------
     BUILD ROWS
  --------------------------- */

  let allRows = [];

  for (const ticket_id in ticketMap) {

    const t = ticketMap[ticket_id];

    if (!t.created_at) continue;

    const createdDate = new Date(t.created_at);
    if (createdDate < start || createdDate >= end) continue;

    if (EXCLUDED_CHANNELS.includes(t.channel)) continue;

    const combined = t.comments.join("\n\n---\n\n");

    const safeCombined =
      combined && combined.length > 30000
        ? combined.slice(0, 30000) + "\n\n[Truncated]"
        : combined;

    allRows.push([
      t.ticket_id,
      t.created_at,
      t.requester_email,
      t.channel,
      t.subject,
      safeCombined
    ]);
  }

  console.log(`Final tickets: ${allRows.length}`);

  /* --------------------------
     BUILD EXCEL
  --------------------------- */

  const worksheet = XLSX.utils.aoa_to_sheet([
    ["Ticket ID", "Created at", "Requester Email", "Channel", "Subject", "All Public Comments"],
    ...allRows
  ]);

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    monthStr.replace("-", "_")
  );

  const fileBuffer = XLSX.write(workbook, {
    type: "buffer",
    bookType: "xlsx"
  });

  /* --------------------------
     EMAIL
  --------------------------- */

  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: process.env.GMAIL_SENDER,
      pass: process.env.GMAIL_APP_PASSWORD
    }
  });

  await transporter.sendMail({
    from: process.env.GMAIL_SENDER,
    to: process.env.EMAIL_RECIPIENTS,
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
