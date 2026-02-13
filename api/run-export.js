import fs from "fs";
import path from "path";

export default async function handler(req, res) {
  try {
    const subdomain = process.env.ZENDESK_SUBDOMAIN;
    const email = process.env.ZENDESK_EMAIL;
    const apiToken = process.env.ZENDESK_API_TOKEN;

    if (!subdomain || !email || !apiToken) {
      return res.status(500).json({ error: "Missing Zendesk credentials" });
    }

    const auth = Buffer.from(`${email}:${apiToken}`).toString("base64");

    // ---- CHECKPOINT FILE ----
    const checkpointPath = path.join("/tmp", "checkpoint.json");

    let startTime;

    if (fs.existsSync(checkpointPath)) {
      const saved = JSON.parse(fs.readFileSync(checkpointPath, "utf8"));
      startTime = saved.end_time;
    } else {
      // December 1, 2025 UTC
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

    // Save checkpoint
    fs.writeFileSync(checkpointPath, JSON.stringify({ end_time: endTime }));

    return res.status(200).json({
      processed: tickets.length,
      next_start_time: endTime,
      completed: endOfStream
    });

  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
