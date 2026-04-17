const express = require("express");
const path = require("path");
const fs = require("fs");
const { runFullScan } = require("./scanner");

const app = express();
const PORT = 3891;
const DATA_FILE = path.join(__dirname, "scan-data.json");

app.use(express.static(path.join(__dirname, "public")));
app.use(express.json());

// Run scan
app.post("/api/scan", async (req, res) => {
  try {
    console.log("Starting scan...");
    const data = await runFullScan();
    fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
    res.json({ success: true, data });
  } catch (e) {
    console.error("Scan error:", e);
    res.status(500).json({ error: e.message });
  }
});

// Get cached scan data
app.get("/api/data", (req, res) => {
  if (fs.existsSync(DATA_FILE)) {
    const data = JSON.parse(fs.readFileSync(DATA_FILE, "utf-8"));
    res.json(data);
  } else {
    res.json(null);
  }
});

app.listen(PORT, () => {
  console.log(`\n  App Scanner running at http://localhost:${PORT}\n`);
});
