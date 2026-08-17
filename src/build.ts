import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { getCitiesOfState, type ICity } from "@countrystatecity/countries";
import type * as XLSXType from "xlsx";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx") as typeof XLSXType;

type CountKey = `${string}|${string}`;

interface CoordinateRow {
  city: string;
  county: string;
  state: string;
  latitude: number;
  longitude: number;
  coordinateSource: string;
}

interface LocationRecord {
  city: string;
  county: string;
  state: string;
  count: number;
  lat: number;
  lng: number;
}

const ROOT = process.cwd();
const INPUT_DIR = path.join(ROOT, "input");
const OUTPUT_DIR = path.join(ROOT, "dist");
const HTML_TEMPLATE_PATH = path.join(ROOT, "src", "template.html");
const HTML_OUTPUT_PATH = path.join(OUTPUT_DIR, "index.html");
const OUTPUT_CSV_PATH = path.join(OUTPUT_DIR, "signer_city_locations.csv");
const OVERRIDES_PATH = path.join(ROOT, "data", "city_overrides.csv");

const PAPER_WORKBOOK = path.join(INPUT_DIR, "SFIW Petition.xlsx");
const ONLINE_EXPORT = path.join(INPUT_DIR, "petition_signatures_jobs_491242344_20260725164505.csv.xls");

const CITY_FIXES = new Map<string, string>([
  ["woodstock", "Woodstock"],
  ["woodstock il 60098", "Woodstock"],
  ["crystal lake", "Crystal Lake"],
  ["mchenry", "McHenry"],
  ["lake in the hills", "Lake in the Hills"],
  ["wonder lake", "Wonder Lake"],
  ["spring grove", "Spring Grove"],
  ["mundelein", "Mundelein"],
  ["delavan", "Delavan"],
  ["dekalb", "De Kalb"],
  ["racine", "Racine"],
  ["s. elgin", "South Elgin"],
  ["los angles", "Los Angeles"]
]);

const CITY_ALIASES = new Map<CountKey, CountKey>([
  ["De Kalb|IL", "DeKalb|IL"],
  ["Fuquay-Varina|NC", "Fuquay-Varina|NC"],
  ["St. Charles|IL", "Saint Charles|IL"],
  ["Village of Lakewood|IL", "Lakewood|IL"]
]);

const STATE_FIXES = new Map<string, string>([
  ["AA", "CA"]
]);

const EXCLUDED_CITY_VALUES = new Set(["daesy ruiz", "unknown"]);

function normalizeCity(raw: unknown): string {
  const city = String(raw ?? "").trim().replace(/\s+/g, " ");
  const key = city.toLowerCase();
  if (!city || EXCLUDED_CITY_VALUES.has(key)) return "";
  const fixed = CITY_FIXES.get(key);
  if (fixed) return fixed;
  if (city === city.toUpperCase() || city === city.toLowerCase()) {
    return city.replace(/\w\S*/g, (word) => word[0].toUpperCase() + word.slice(1).toLowerCase());
  }
  return city;
}

function normalizeState(raw: unknown): string {
  const state = String(raw ?? "").trim().toUpperCase();
  return STATE_FIXES.get(state) ?? state;
}

function normalizeCounty(raw: unknown): string {
  return String(raw ?? "").trim().replace(/\s+/g, " ");
}

function keyFor(city: string, state: string): CountKey {
  return `${city}|${state}`;
}

function parseSeparated(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (quoted) {
      if (char === '"' && next === '"') {
        field += '"';
        i += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        field += char;
      }
      continue;
    }

    if (char === '"') {
      quoted = true;
    } else if (char === delimiter) {
      row.push(field);
      field = "";
    } else if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (char !== "\r") {
      field += char;
    }
  }

  if (field || row.length) {
    row.push(field);
    rows.push(row);
  }

  return rows;
}

function parseDelimited(text: string, delimiter: string): Record<string, string>[] {
  const rows = parseSeparated(text, delimiter);
  const headers = rows.shift()?.map((header) => header.trim()) ?? [];
  return rows.map((values) => {
    const record: Record<string, string> = {};
    headers.forEach((header, index) => {
      record[header] = values[index] ?? "";
    });
    return record;
  });
}

function addLocation(
  counts: Map<CountKey, number>,
  counties: Map<CountKey, Map<string, number>>,
  cityRaw: unknown,
  countyRaw: unknown,
  stateRaw: unknown
): boolean {
  const city = normalizeCity(cityRaw);
  const state = normalizeState(stateRaw);
  const county = normalizeCounty(countyRaw);
  if (!city || !state || city.toLowerCase() === "city" || state.toLowerCase() === "state") return false;

  const key = keyFor(city, state);
  counts.set(key, (counts.get(key) ?? 0) + 1);
  if (county) {
    const countyCounts = counties.get(key) ?? new Map<string, number>();
    countyCounts.set(county, (countyCounts.get(county) ?? 0) + 1);
    counties.set(key, countyCounts);
  }
  return true;
}

function readPaperWorkbook(counts: Map<CountKey, number>, counties: Map<CountKey, Map<string, number>>): number {
  if (!fs.existsSync(PAPER_WORKBOOK)) return 0;

  const workbook = XLSX.readFile(PAPER_WORKBOOK);
  const sheet = workbook.Sheets.Paper;
  if (!sheet) throw new Error("Expected sheet named 'Paper' in SFIW Petition.xlsx");

  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, blankrows: false });
  let added = 0;
  for (const row of rows) {
    if (row.length < 4) continue;
    if (row[0] === "Name" && row[1] === "City" && row[2] === "County" && row[3] === "State") continue;
    if (addLocation(counts, counties, row[1], row[2], row[3])) added += 1;
  }
  return added;
}

function readOnlineExport(counts: Map<CountKey, number>, counties: Map<CountKey, Map<string, number>>): number {
  if (!fs.existsSync(ONLINE_EXPORT)) return 0;

  const text = fs.readFileSync(ONLINE_EXPORT, "utf16le");
  const records = parseDelimited(text, "\t");
  let added = 0;
  for (const record of records) {
    if (addLocation(counts, counties, record.City, "", record.State)) added += 1;
  }
  return added;
}

function readCoordinateOverrides(): Map<CountKey, CoordinateRow> {
  if (!fs.existsSync(OVERRIDES_PATH)) return new Map();

  const text = fs.readFileSync(OVERRIDES_PATH, "utf8");
  const records = parseDelimited(text, ",");
  const coordinates = new Map<CountKey, CoordinateRow>();

  for (const record of records) {
    const city = normalizeCity(record.city);
    const state = normalizeState(record.state);
    const latitude = Number(record.latitude);
    const longitude = Number(record.longitude);
    if (!city || !state || !Number.isFinite(latitude) || !Number.isFinite(longitude)) continue;
    coordinates.set(keyFor(city, state), {
      city,
      county: normalizeCounty(record.county),
      state,
      latitude,
      longitude,
      coordinateSource: record.coordinate_source ?? ""
    });
  }

  return coordinates;
}

function cityDatasetKey(city: string, state: string): CountKey {
  return CITY_ALIASES.get(keyFor(city, state)) ?? keyFor(city, state);
}

const stateCityCache = new Map<string, Promise<ICity[]>>();

async function getStateCities(state: string): Promise<ICity[]> {
  if (!stateCityCache.has(state)) {
    stateCityCache.set(state, getCitiesOfState("US", state));
  }
  return stateCityCache.get(state)!;
}

async function findDatasetCoordinate(city: string, state: string): Promise<CoordinateRow | undefined> {
  const [lookupCity, lookupState] = cityDatasetKey(city, state).split("|");
  const cities = await getStateCities(lookupState);
  const match = cities.find((candidate) => candidate.name.toLowerCase() === lookupCity.toLowerCase());
  if (!match) return undefined;

  return {
    city,
    county: "",
    state,
    latitude: Number(match.latitude),
    longitude: Number(match.longitude),
    coordinateSource: "@countrystatecity/countries"
  };
}

function mostCommonCounty(countyCounts: Map<string, number> | undefined): string {
  if (!countyCounts) return "";
  return [...countyCounts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0]?.[0] ?? "";
}

async function buildLocations(): Promise<{ locations: LocationRecord[]; sources: Record<string, number> }> {
  const counts = new Map<CountKey, number>();
  const counties = new Map<CountKey, Map<string, number>>();
  const sources = {
    paper: readPaperWorkbook(counts, counties),
    online: readOnlineExport(counts, counties)
  };

  if (!counts.size) {
    throw new Error(`No signer input rows found. Put source sheets in ${path.relative(ROOT, INPUT_DIR)}.`);
  }

  const overrides = readCoordinateOverrides();
  const missing: string[] = [];
  const locations: LocationRecord[] = [];

  for (const [key, count] of counts.entries()) {
    const [city, state] = key.split("|");
    const coordinate = overrides.get(key) ?? await findDatasetCoordinate(city, state);
    if (!coordinate) {
      missing.push(`${city}, ${state}`);
      continue;
    }
    locations.push({
      city,
      county: mostCommonCounty(counties.get(key)) || coordinate.county,
      state,
      count,
      lat: coordinate.latitude,
      lng: coordinate.longitude
    });
  }

  if (missing.length) {
    throw new Error(`Missing coordinates. Add these to data/city_overrides.csv:\n${missing.map((item) => `- ${item}`).join("\n")}`);
  }

  locations.sort((a, b) => b.count - a.count || a.city.localeCompare(b.city));
  return { locations, sources };
}

function csvEscape(value: string | number): string {
  const text = String(value);
  return /[",\n\r]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

async function writeAggregateCsv(locations: LocationRecord[]): Promise<void> {
  const overrides = readCoordinateOverrides();
  const lines = ["city,county,state,signer_count,latitude,longitude,coordinate_source"];
  for (const location of locations) {
    const coordinate = overrides.get(keyFor(location.city, location.state)) ?? await findDatasetCoordinate(location.city, location.state);
    lines.push([
      location.city,
      location.county,
      location.state,
      location.count,
      location.lat,
      location.lng,
      coordinate?.coordinateSource ?? ""
    ].map(csvEscape).join(","));
  }
  fs.writeFileSync(OUTPUT_CSV_PATH, `${lines.join("\n")}\n`);
}

function writeHtml(locations: LocationRecord[]): void {
  const html = fs.readFileSync(HTML_TEMPLATE_PATH, "utf8");
  const locationsJson = JSON.stringify(locations, null, 2);
  if (!html.includes("__LOCATIONS_JSON__")) {
    throw new Error("Could not find __LOCATIONS_JSON__ placeholder in src/template.html");
  }
  const updated = html.replace("__LOCATIONS_JSON__", locationsJson);
  fs.writeFileSync(HTML_OUTPUT_PATH, updated);
}

const { locations, sources } = await buildLocations();
fs.mkdirSync(OUTPUT_DIR, { recursive: true });
await writeAggregateCsv(locations);
writeHtml(locations);

const total = locations.reduce((sum, location) => sum + location.count, 0);
console.log(`Built ${locations.length} aggregate locations for ${total} signers.`);
console.log(`Source rows: paper=${sources.paper}, online=${sources.online}`);
