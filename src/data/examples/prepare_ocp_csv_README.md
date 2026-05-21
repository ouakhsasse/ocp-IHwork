# Preparing the OCP Benguerir Annual CSV

Real OCP annual simulation results require an hourly annual profile with 8760 rows for a normal year or 8784 rows for a leap year.

## Expected Format

Preferred columns:

```csv
timestamp,production_mwh,consumption_mwh
2025-01-01T00:00:00,0,25
2025-01-01T01:00:00,0,25
2025-01-01T12:00:00,42.5,25
```

## Export From Excel

1. Open the real OCP PV production database in Excel.
2. Keep one row per hour.
3. Export the sheet as CSV.
4. Rename or map the timestamp column to `timestamp`.
5. Rename or map the PV hourly production column to `production_mwh`.
6. Rename or map the industrial demand column to `consumption_mwh` when it exists.

## Accepted Column Aliases

PV production can be named:

- `production_mwh`
- `pv_production_mwh`
- `pv_mwh`

Load or demand can be named:

- `consumption_mwh`
- `load_mwh`
- `demand_mwh`

If consumption is missing, the backend can generate the OCP constant 15 MW industrial load. With hourly simulation steps, this becomes 15 MWh every hour.

## Timestamp Examples

Use ordered hourly timestamps:

- `2025-01-01T00:00:00`
- `2025-01-01T00:00:00.000Z`
- `2025-01-01 00:00:00`

Avoid duplicate timestamps and missing hours. The backend validates ordering, duplicates, numeric values, negative values, and annual profile length.
