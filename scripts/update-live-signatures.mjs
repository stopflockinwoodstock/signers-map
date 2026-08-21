import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const PETITION_URL = "https://www.change.org/p/we-want-flock-safety-out-of-woodstock-il";
const OUTPUT_PATH = path.join(process.cwd(), "docs", "live-signatures.json");
const PAPER_PATH = path.join(process.cwd(), "docs", "current-paper-signatures.json");
const HISTORY_PATH = path.join(process.cwd(), "docs", "live-signatures-history.json");
const HISTORY_DAYS = 15;

const response = await fetch(PETITION_URL, {
  headers: {
    "user-agent": "Woodstock-IL-Cameras signer map signature updater"
  }
});

if (!response.ok) {
  throw new Error(`Change.org returned HTTP ${response.status}`);
}

const html = await response.text();
const match = html.match(/aria-label=["']([\d,]+)\s+of\s+verified\s+signatures/i);
if (!match) {
  throw new Error("Could not find the verified signature count in the Change.org page");
}

const onlineSignatures = Number.parseInt(match[1].replaceAll(",", ""), 10);
if (!Number.isSafeInteger(onlineSignatures) || onlineSignatures < 0) {
  throw new Error(`Invalid online signature count: ${match[1]}`);
}

const paperPayload = JSON.parse(await fs.readFile(PAPER_PATH, "utf8"));
const paperSignatures = paperPayload.paperSignatures;
if (!Number.isSafeInteger(paperSignatures) || paperSignatures < 0) {
  throw new Error("Current paper signature data contains an invalid count");
}

const fetchedAt = new Date().toISOString();
const totalSignatures = onlineSignatures + paperSignatures;
const today = new Intl.DateTimeFormat("en-CA", {
  timeZone: "America/Chicago",
  year: "numeric",
  month: "2-digit",
  day: "2-digit"
}).format(new Date());

let history = [];
try {
  history = JSON.parse(await fs.readFile(HISTORY_PATH, "utf8"));
  if (!Array.isArray(history)) throw new Error("Signature history must be an array");
} catch (error) {
  if (error.code !== "ENOENT") throw error;
}

const payload = {
  onlineSignatures,
  fetchedAt,
  source: PETITION_URL
};

let liveChanged = true;
try {
  const existing = JSON.parse(await fs.readFile(OUTPUT_PATH, "utf8"));
  liveChanged = existing.onlineSignatures !== onlineSignatures;
} catch (error) {
  if (error.code !== "ENOENT") throw error;
}

if (liveChanged) {
  await fs.writeFile(OUTPUT_PATH, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
  console.log(`Wrote ${onlineSignatures.toLocaleString()} online signatures to ${path.relative(process.cwd(), OUTPUT_PATH)}.`);
} else {
  console.log(`Online signature count is unchanged at ${onlineSignatures.toLocaleString()}.`);
}

const currentEntry = {
  date: today,
  onlineSignatures,
  paperSignatures,
  totalSignatures,
  fetchedAt
};
const existingToday = history.find((entry) => entry.date === today);
const nextHistory = history
  .filter((entry) => entry.date !== today)
  .concat(currentEntry)
  .sort((left, right) => left.date.localeCompare(right.date))
  .slice(-HISTORY_DAYS);
const historyChanged = !existingToday || existingToday.onlineSignatures !== onlineSignatures ||
  existingToday.paperSignatures !== paperSignatures || existingToday.totalSignatures !== totalSignatures;
if (historyChanged) {
  await fs.writeFile(HISTORY_PATH, `${JSON.stringify(nextHistory, null, 2)}\n`, "utf8");
  console.log(`Recorded ${totalSignatures.toLocaleString()} total signatures for ${today}.`);
} else {
  console.log(`Today's history already records ${totalSignatures.toLocaleString()} total signatures.`);
}
