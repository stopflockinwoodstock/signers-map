import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { getCitiesOfState, type ICity } from "@countrystatecity/countries";
import type * as XLSXType from "xlsx";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx") as typeof XLSXType;

type CityKey = `${string}|${string}`;

type AggregateLocation = {
  city: string;
  state: string;
  count: number;
  counties: Map<string, number>;
};

type Coordinate = {
  lat: number;
  lng: number;
  county: string;
  source: string;
};

type MappedLocation = {
  city: string;
  county: string;
  state: string;
  count: number;
  lat: number;
  lng: number;
  coordinateSource: string;
};

type MissingLocation = {
  city: string;
  state: string;
  count: number;
};

type AddLocationResult = "accepted" | "missingCity" | "invalidState";

type SourceStats = {
  rows: number;
  accepted: number;
  missingCity: number;
  invalidState: number;
  input: string;
};

const ROOT = process.cwd();
const INPUT_DIR = path.join(ROOT, "input");
const OUTPUT_DIR = path.join(ROOT, "dist");
const TEMPLATE_PATH = path.join(ROOT, "src", "template.html");
const OVERRIDES_PATH = path.join(ROOT, "data", "city_overrides.csv");

const INPUTS = {
  paperCsv: path.join(INPUT_DIR, "SFIW Petition.csv"),
  paperXlsx: path.join(INPUT_DIR, "SFIW Petition.xlsx"),
  online: path.join(INPUT_DIR, "petition_signatures_jobs_491242344_20260818002741.csv")
};

const CITY_FIXES = new Map<string, string>([
  ["brookfeld", "Brookfield"],
  ["crystal lake", "Crystal Lake"],
  ["dekalb", "De Kalb"],
  ["hp", "Highland Park"],
  ["lake in the hills", "Lake in the Hills"],
  ["los angles", "Los Angeles"],
  ["mchenry", "McHenry"],
  ["s. elgin", "South Elgin"],
  ["spring grove", "Spring Grove"],
  ["wonder lake", "Wonder Lake"],
  ["woodstock il 60098", "Woodstock"],
  ["Charolette", "Charlotte"]
]);

const PACKAGE_ALIASES = new Map<CityKey, CityKey>([
  ["Campton Hills|IL", "Village of Campton Hills|IL"],
  ["De Kalb|IL", "DeKalb|IL"],
  ["Saint Louis|MO", "St. Louis|MO"],
  ["St. Charles|IL", "Saint Charles|IL"],
  ["Village of Lakewood|IL", "Lakewood|IL"]
]);

const STATE_FIXES = new Map([["AA", "CA"]]);
const IGNORED_CITY_VALUES = new Set(["", "city", "daesy ruiz", "illinois", "unknown"]);

function titleCase(value: string): string {
  return value.replace(/\w\S*/g, (word) => word[0].toUpperCase() + word.slice(1).toLowerCase());
}

function cityName(raw: unknown): string {
  const city = clean(raw);
  const key = city.toLowerCase();
  if (IGNORED_CITY_VALUES.has(key)) return "";
  return CITY_FIXES.get(key) ?? (city === city.toLowerCase() || city === city.toUpperCase() ? titleCase(city) : city);
}

function stateCode(raw: unknown): string {
  const state = clean(raw).toUpperCase();
  return STATE_FIXES.get(state) ?? state;
}

function clean(raw: unknown): string {
  return String(raw ?? "").trim().replace(/^"|"$/g, "").trim().replace(/\s+/g, " ");
}

function cityKey(city: string, state: string): CityKey {
  return `${city}|${state}`;
}

function addLocation(
  locations: Map<CityKey, AggregateLocation>,
  cityRaw: unknown,
  countyRaw: unknown,
  stateRaw: unknown
): AddLocationResult {
  const city = cityName(cityRaw);
  const state = stateCode(stateRaw);
  const county = clean(countyRaw);
  if (!city) return "missingCity";
  if (!/^[A-Z]{2}$/.test(state)) return "invalidState";

  const key = cityKey(city, state);
  const location = locations.get(key) ?? { city, state, count: 0, counties: new Map<string, number>() };
  location.count += 1;
  if (county) location.counties.set(county, (location.counties.get(county) ?? 0) + 1);
  locations.set(key, location);
  return "accepted";
}

function emptyStats(input: string): SourceStats {
  return {
    rows: 0,
    accepted: 0,
    missingCity: 0,
    invalidState: 0,
    input
  };
}

function trackResult(stats: SourceStats, result: AddLocationResult): void {
  stats.rows += 1;
  stats[result] += 1;
}

function readPaperCsv(locations: Map<CityKey, AggregateLocation>): SourceStats | undefined {
  if (!fs.existsSync(INPUTS.paperCsv)) return undefined;

  const stats = emptyStats(path.relative(ROOT, INPUTS.paperCsv));
  const rows = fs.readFileSync(INPUTS.paperCsv, "utf8").split(/\r?\n/).filter(Boolean);
  for (const row of rows) {
    const cells = row.split("\t");
    if (cells[0] === "Name" && cells[1] === "City") continue;
    trackResult(stats, addLocation(locations, cells[1], cells[2], cells[3]));
  }
  return stats;
}

function readPaperXlsx(locations: Map<CityKey, AggregateLocation>): SourceStats {
  const stats = emptyStats(path.relative(ROOT, INPUTS.paperXlsx));
  if (!fs.existsSync(INPUTS.paperXlsx)) return stats;

  const workbook = XLSX.readFile(INPUTS.paperXlsx);
  const sheet = workbook.Sheets?.Paper ?? workbook.Sheets?.Sheet1;
  if (!sheet) throw new Error("Expected a sheet named 'Paper' in SFIW Petition.xlsx");

  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, blankrows: false });
  for (const row of rows) {
    if (row[0] === "Name" && row[1] === "City") continue;
    trackResult(stats, addLocation(locations, row[1], row[2], row[3]));
  }
  return stats;
}

function readPaper(locations: Map<CityKey, AggregateLocation>): SourceStats {
  return readPaperCsv(locations) ?? readPaperXlsx(locations);
}

function readOnline(locations: Map<CityKey, AggregateLocation>): SourceStats {
  const stats = emptyStats(path.relative(ROOT, INPUTS.online));
  if (!fs.existsSync(INPUTS.online)) return stats;

  const rows = fs.readFileSync(INPUTS.online, "utf16le").split(/\r?\n/).filter(Boolean);
  const headers = rows.shift()?.split("\t") ?? [];
  const cityIndex = headers.indexOf("City");
  const stateIndex = headers.indexOf("State");
  if (cityIndex < 0 || stateIndex < 0) throw new Error("Online export must include City and State columns");

  for (const row of rows) {
    const cells = row.split("\t");
    trackResult(stats, addLocation(locations, cells[cityIndex], "", cells[stateIndex]));
  }
  return stats;
}

function parseOverrideCsv(): Map<CityKey, Coordinate> {
  if (!fs.existsSync(OVERRIDES_PATH)) return new Map();

  const rows = fs.readFileSync(OVERRIDES_PATH, "utf8").trim().split(/\r?\n/).slice(1);
  const overrides = new Map<CityKey, Coordinate>();
  for (const row of rows) {
    const [cityRaw, countyRaw, stateRaw, latRaw, lngRaw, sourceRaw] = row.split(",");
    const city = cityName(cityRaw);
    const state = stateCode(stateRaw);
    const lat = Number(latRaw);
    const lng = Number(lngRaw);
    if (!city || !state || !Number.isFinite(lat) || !Number.isFinite(lng)) continue;
    overrides.set(cityKey(city, state), {
      lat,
      lng,
      county: clean(countyRaw),
      source: clean(sourceRaw) || "manual_override"
    });
  }
  return overrides;
}

const packageCityCache = new Map<string, Promise<ICity[]>>();

function packageCities(state: string): Promise<ICity[]> {
  if (!packageCityCache.has(state)) {
    packageCityCache.set(state, getCitiesOfState("US", state));
  }
  return packageCityCache.get(state)!;
}

async function packageCoordinate(city: string, state: string): Promise<Coordinate | undefined> {
  const [lookupCity, lookupState] = (PACKAGE_ALIASES.get(cityKey(city, state)) ?? cityKey(city, state)).split("|");
  const match = (await packageCities(lookupState)).find((candidate) => candidate.name.toLowerCase() === lookupCity.toLowerCase());
  if (!match) return undefined;

  return {
    lat: Number(match.latitude),
    lng: Number(match.longitude),
    county: "",
    source: "@countrystatecity/countries"
  };
}

function mostCommonCounty(location: AggregateLocation): string {
  return [...location.counties.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0]?.[0] ?? "";
}

async function mapLocations(locations: Map<CityKey, AggregateLocation>): Promise<{
  mapped: MappedLocation[];
  missing: MissingLocation[];
}> {
  const overrides = parseOverrideCsv();
  const mapped: MappedLocation[] = [];
  const missing: MissingLocation[] = [];

  for (const location of locations.values()) {
    const coordinate = overrides.get(cityKey(location.city, location.state)) ?? await packageCoordinate(location.city, location.state);
    if (!coordinate) {
      missing.push({ city: location.city, state: location.state, count: location.count });
      continue;
    }

    mapped.push({
      city: location.city,
      county: mostCommonCounty(location) || coordinate.county,
      state: location.state,
      count: location.count,
      lat: coordinate.lat,
      lng: coordinate.lng,
      coordinateSource: coordinate.source
    });
  }

  return {
    mapped: mapped.sort((a, b) => b.count - a.count || a.city.localeCompare(b.city)),
    missing: missing.sort((a, b) => b.count - a.count || a.city.localeCompare(b.city))
  };
}

function csvCell(value: string | number): string {
  const text = String(value);
  return /[",\n\r]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function writeCsv(locations: MappedLocation[]): void {
  const rows = [
    ["city", "county", "state", "signer_count", "latitude", "longitude", "coordinate_source"],
    ...locations.map((location) => [
      location.city,
      location.county,
      location.state,
      location.count,
      location.lat,
      location.lng,
      location.coordinateSource
    ])
  ];
  fs.writeFileSync(path.join(OUTPUT_DIR, "signer_city_locations.csv"), `${rows.map((row) => row.map(csvCell).join(",")).join("\n")}\n`);
}

function writeMissingCsv(locations: MissingLocation[]): void {
  const rows = [
    ["city", "state", "signer_count"],
    ...locations.map((location) => [location.city, location.state, location.count])
  ];
  fs.writeFileSync(path.join(OUTPUT_DIR, "missing_coordinates.csv"), `${rows.map((row) => row.map(csvCell).join(",")).join("\n")}\n`);
}

function buildTimestamp(): string {
  const formatted = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Chicago",
    month: "long",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true
  }).format(new Date());
  return `${formatted} CT`;
}

function writeHtml(locations: MappedLocation[], totalSignatures: number, updatedAt: string): void {
  const template = fs.readFileSync(TEMPLATE_PATH, "utf8");
  if (!template.includes("__LOCATIONS_JSON__")) throw new Error("Missing __LOCATIONS_JSON__ placeholder in src/template.html");
  if (!template.includes("__TOTAL_SIGNATURES__")) throw new Error("Missing __TOTAL_SIGNATURES__ placeholder in src/template.html");
  if (!template.includes("__UPDATED_AT__")) throw new Error("Missing __UPDATED_AT__ placeholder in src/template.html");
  const html = template
    .replace("__LOCATIONS_JSON__", JSON.stringify(locations, null, 2))
    .replace("__TOTAL_SIGNATURES__", String(totalSignatures))
    .replace("__UPDATED_AT__", updatedAt);
  fs.writeFileSync(path.join(OUTPUT_DIR, "index.html"), html);
}

const aggregateLocations = new Map<CityKey, AggregateLocation>();
const sources = {
  paper: readPaper(aggregateLocations),
  online: readOnline(aggregateLocations)
};

if (!aggregateLocations.size) {
  throw new Error(`No signer input rows found. Put source sheets in ${path.relative(ROOT, INPUT_DIR)}.`);
}

const { mapped: mappedLocations, missing: missingLocations } = await mapLocations(aggregateLocations);
fs.mkdirSync(OUTPUT_DIR, { recursive: true });
writeCsv(mappedLocations);
writeMissingCsv(missingLocations);
const mappedTotal = mappedLocations.reduce((sum, location) => sum + location.count, 0);
const missingCoordinateTotal = missingLocations.reduce((sum, location) => sum + location.count, 0);
const inputTotal = sources.paper.accepted + sources.online.accepted;
const updatedAt = buildTimestamp();
writeHtml(mappedLocations, inputTotal, updatedAt);

console.log(`Built ${mappedLocations.length} aggregate locations for ${mappedTotal} mapped signers.`);
console.log(`Input rows: paper=${sources.paper.rows}, online=${sources.online.rows}`);
console.log(`Accepted signer rows with city/state: paper=${sources.paper.accepted}, online=${sources.online.accepted}, total=${inputTotal}`);
console.log(`Rows without city: paper=${sources.paper.missingCity}, online=${sources.online.missingCity}`);
console.log(`Rows with invalid/missing state: paper=${sources.paper.invalidState}, online=${sources.online.invalidState}`);
console.log(`Paper source: ${sources.paper.input}`);
console.log(`Updated at: ${updatedAt}`);
if (missingLocations.length) {
  console.log(`Skipped ${missingLocations.length} locations / ${missingCoordinateTotal} signers without coordinates. See dist/missing_coordinates.csv.`);
}
