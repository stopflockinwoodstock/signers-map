# Signer Map

Aggregate signer locations into an OpenStreetMap-based HTML map.

## Privacy Rule

Raw signer files stay in `input/` and are ignored by Git. The build only emits aggregate city/state counts plus city coordinates.

## Rebuild

1. Put source files in `input/`:
   - `SFIW Petition.xlsx`
   - `petition_signatures_jobs_491242344_20260725164505.csv.xls`
2. Install dependencies:
   ```sh
   npm install
   ```
3. Generate outputs:
   ```sh
   npm run build
   ```

Generated outputs:

- `dist/index.html`
- `dist/signer_city_locations.csv`
- `dist/missing_coordinates.csv`

Coordinate lookup:

- `@countrystatecity/countries` for city/state coordinates
- `data/city_overrides.csv` only for package misses
- Rows that still cannot be resolved are skipped from the map and written to `dist/missing_coordinates.csv`

Tracked source files:

- `src/build.ts`
- `src/template.html`
- `data/city_overrides.csv`

## Coordinate Data Attribution

City coordinates are primarily provided by `@countrystatecity/countries`, backed by Countries States Cities Database.

Data by Countries States Cities Database: https://github.com/dr5hn/countries-states-cities-database

License: ODbL v1.0
