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

const ROOT = process.cwd();
const INPUT_DIR = path.join(ROOT, "input");
const OUTPUT_DIR = path.join(ROOT, "dist");
const TEMPLATE_PATH = path.join(ROOT, "src", "template.html");
const OVERRIDES_PATH = path.join(ROOT, "data", "city_overrides.csv");

const INPUTS = {
  paper: path.join(INPUT_DIR, "SFIW Petition.xlsx"),
  online: path.join(INPUT_DIR, "petition_signatures_jobs_491242344_20260816223514.csv")
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
  ["woodstock il 60098", "Woodstock"]
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
): boolean {
  const city = cityName(cityRaw);
  const state = stateCode(stateRaw);
  const county = clean(countyRaw);
  if (!city || !/^[A-Z]{2}$/.test(state)) return false;

  const key = cityKey(city, state);
  const location = locations.get(key) ?? { city, state, count: 0, counties: new Map<string, number>() };
  location.count += 1;
  if (county) location.counties.set(county, (location.counties.get(county) ?? 0) + 1);
  locations.set(key, location);
  return true;
}

function readPaper(locations: Map<CityKey, AggregateLocation>): number {
  if (!fs.existsSync(INPUTS.paper)) return 0;

  const workbook = XLSX.readFile(INPUTS.paper);
  const sheet = workbook.Sheets?.Paper ?? workbook.Sheets?.Sheet1;
  if (!sheet) throw new Error("Expected a sheet named 'Paper' in SFIW Petition.xlsx");

  let count = 0;
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, blankrows: false });
  for (const row of rows) {
    if (row[0] === "Name" && row[1] === "City") continue;
    if (addLocation(locations, row[1], row[2], row[3])) count += 1;
  }
  return count;
}

function readOnline(locations: Map<CityKey, AggregateLocation>): number {
  if (!fs.existsSync(INPUTS.online)) return 0;

  const rows = fs.readFileSync(INPUTS.online, "utf16le").split(/\r?\n/).filter(Boolean);
  const headers = rows.shift()?.split("\t") ?? [];
  const cityIndex = headers.indexOf("City");
  const stateIndex = headers.indexOf("State");
  if (cityIndex < 0 || stateIndex < 0) throw new Error("Online export must include City and State columns");

  let count = 0;
  for (const row of rows) {
    const cells = row.split("\t");
    if (addLocation(locations, cells[cityIndex], "", cells[stateIndex])) count += 1;
  }
  return count;
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

function writeHtml(locations: MappedLocation[]): void {
  const template = fs.readFileSync(TEMPLATE_PATH, "utf8");
  if (!template.includes("__LOCATIONS_JSON__")) throw new Error("Missing __LOCATIONS_JSON__ placeholder in src/template.html");
  const html = template.replace("__LOCATIONS_JSON__", JSON.stringify(locations, null, 2));
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
writeHtml(mappedLocations);

const total = mappedLocations.reduce((sum, location) => sum + location.count, 0);
console.log(`Built ${mappedLocations.length} aggregate locations for ${total} signers.`);
console.log(`Source rows: paper=${sources.paper}, online=${sources.online}`);
if (missingLocations.length) {
  const missingTotal = missingLocations.reduce((sum, location) => sum + location.count, 0);
  console.log(`Skipped ${missingLocations.length} locations / ${missingTotal} signers without coordinates. See dist/missing_coordinates.csv.`);
}
