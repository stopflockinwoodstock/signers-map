import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const PETITION_URL = "https://www.change.org/p/we-want-flock-safety-out-of-woodstock-il";
const OUTPUT_PATH = path.join(process.cwd(), "docs", "live-signatures.json");

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

try {
  const existing = JSON.parse(await fs.readFile(OUTPUT_PATH, "utf8"));
  if (existing.onlineSignatures === onlineSignatures) {
    console.log(`Online signature count is unchanged at ${onlineSignatures.toLocaleString()}.`);
    process.exit(0);
  }
} catch (error) {
  if (error.code !== "ENOENT") throw error;
}

const payload = {
  onlineSignatures,
  fetchedAt: new Date().toISOString(),
  source: PETITION_URL
};

await fs.writeFile(OUTPUT_PATH, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
console.log(`Wrote ${onlineSignatures.toLocaleString()} online signatures to ${path.relative(process.cwd(), OUTPUT_PATH)}.`);
