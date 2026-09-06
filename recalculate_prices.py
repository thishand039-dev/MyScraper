"""
recalculate_prices.py
----------------------
Fixes an already-downloaded vehicle's model and/or price WITHOUT
touching the browser or re-downloading anything. Use this when:

    - The wrong vehicle_model was picked for a lot (which also affects
      the tax lookup, since vehicleDetails.json is keyed by the exact
      model name) - this script renames the folder and re-runs the
      calculation under the corrected model.
    - A "Result Price" scrape issue was flagged by download_and_price.py
      (see that vehicle's SCRAPE_ISSUE.txt) - check the site manually,
      then put the correct price in this file's price_override column
      to resolve it.
    - Any other case where you need to manually correct a price without
      a full re-scrape.

INPUT: recalculate_input.xlsx, with columns:
    vehicle_url      - same URL used in input_vehicles.xlsx (used only to
                        derive vehicle_code, so the existing folder can be
                        found - no page is visited).
    vehicle_model     - the CORRECTED exact model name (matches a key in
                        vehicleDetails.json). If unchanged from before,
                        just repeat the same value.
    price_override    - optional. Leave blank to keep whatever price is
                        already stored for this vehicle. Fill in a JPY
                        number to set/replace it (e.g. to resolve a
                        flagged scrape issue, or correct a wrong value).

WHAT IT DOES, per row:
    1. Finds the existing output folder by vehicle_code (matches
       "<vehicle_code>-*" under downloads/ - works even if the folder was
       created under the wrong model name).
    2. Updates vehicle_info.json's vehicle_model (and price, if
       price_override was given - this also marks the vehicle as sold).
    3. Renames the folder if the model name changed.
    4. Regenerates description.txt from scratch (so sold/upcoming
       formatting is always consistent with the current data).
    5. Re-runs price_calculator.process_vehicle_folder() only if the
       vehicle is marked sold with a price - otherwise clears any stale
       price_breakdown.txt.
    6. Clears SCRAPE_ISSUE.txt if a price_override resolved it.

This never launches Playwright, so it's fast - safe to run as often as
needed while double-checking a batch of listings.
"""

import sys
import json
import shutil
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
from pathlib import Path

import price_calculator
import scraper_common as common

RECALC_INPUT_EXCEL = "recalculate_input.xlsx"


def find_existing_folder(vehicle_code: str) -> Path | None:
    """Matches by vehicle_code prefix, regardless of what model name the
    folder was originally created under."""
    output_dir = Path(common.OUTPUT_DIR)
    if not output_dir.exists():
        return None
    matches = list(output_dir.glob(f"{vehicle_code}-*"))
    if not matches:
        return None
    return matches[0]


def run():
    df = pd.read_excel(RECALC_INPUT_EXCEL)
    df.columns = [c.strip().lower() for c in df.columns]
    assert "vehicle_url" in df.columns and "vehicle_model" in df.columns, \
        "Excel must have 'vehicle_url' and 'vehicle_model' columns " \
        "('price_override' is optional)"

    results = []

    for _, row in df.iterrows():
        vehicle_url = str(row["vehicle_url"]).strip()
        corrected_model = str(row["vehicle_model"]).strip()
        raw_override = row.get("price_override", None)
        price_override = None
        if raw_override is not None and str(raw_override).strip() not in ("", "nan", "None"):
            price_override = common.clean_price(str(raw_override).strip())

        vehicle_code = common.extract_vehicle_code(vehicle_url)
        folder_path = find_existing_folder(vehicle_code)

        if folder_path is None:
            print(f"  ! No existing folder found for vehicle_code '{vehicle_code}' "
                  f"({vehicle_url}). Run download_and_price.py for this vehicle first.")
            results.append({
                "vehicle_url": vehicle_url,
                "vehicle_model": corrected_model,
                "status": "folder_not_found",
            })
            continue

        info_path = folder_path / "vehicle_info.json"
        if not info_path.exists():
            print(f"  ! {folder_path} has no vehicle_info.json - cannot recalculate.")
            results.append({
                "vehicle_url": vehicle_url,
                "vehicle_model": corrected_model,
                "status": "vehicle_info_missing",
            })
            continue

        with open(info_path, "r", encoding="utf-8") as f:
            vehicle_info = json.load(f)

        old_model = vehicle_info.get("vehicle_model", "")
        vehicle_info["vehicle_model"] = corrected_model

        if price_override is not None:
            vehicle_info["price"] = price_override
            vehicle_info["sold"] = True

        # Move the folder if the corrected model changes its name.
        new_folder_name = common.safe_folder_name(vehicle_code, corrected_model)
        new_folder_path = folder_path.parent / new_folder_name
        if new_folder_path != folder_path:
            if new_folder_path.exists():
                print(f"  ! Target folder {new_folder_path} already exists - "
                      f"not overwriting. Skipping this row.")
                results.append({
                    "vehicle_url": vehicle_url,
                    "vehicle_model": corrected_model,
                    "status": "target_folder_exists",
                })
                continue
            shutil.move(str(folder_path), str(new_folder_path))
            folder_path = new_folder_path
            print(f"  Renamed folder for model change "
                  f"('{old_model}' -> '{corrected_model}'): {folder_path}")

        info_path = folder_path / "vehicle_info.json"
        with open(info_path, "w", encoding="utf-8") as f:
            json.dump(vehicle_info, f, indent=2, ensure_ascii=False)

        common.write_description_file(folder_path, vehicle_info)

        sold = bool(vehicle_info.get("sold")) and vehicle_info.get("price") not in (None, "", "None")

        if sold:
            common.clear_scrape_issue_file(folder_path)
            price_calc_ok = price_calculator.process_vehicle_folder(folder_path)
            status = "ok" if price_calc_ok else "ok_price_calc_stopped"
            print(f"  [OK] {corrected_model} - price recalculated -> {folder_path}")
        else:
            common.clear_price_breakdown_file(folder_path)
            status = "upcoming"
            print(f"  [OK] {corrected_model} - marked upcoming (no price) -> {folder_path}")

        results.append({
            "vehicle_url": vehicle_url,
            "vehicle_model": corrected_model,
            "status": status,
            "folder": str(folder_path),
        })

    out_df = pd.DataFrame(results)
    out_path = Path(common.OUTPUT_DIR) / "results_recalculate.xlsx"
    out_df.to_excel(out_path, index=False)
    print(f"\nDone. Results saved to {out_path}")


if __name__ == "__main__":
    run()
