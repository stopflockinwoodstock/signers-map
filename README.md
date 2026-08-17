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

- `Signers Heat Map.html`
- `signer_city_locations.csv`

Coordinate dictionary:

- `data/city_coordinates.csv`
