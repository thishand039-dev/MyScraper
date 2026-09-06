"""
download_and_price.py
----------------------
Reads (vehicle_url, vehicle_model) rows from Excel. For each one, logs in
(once, session reused after), navigates to the vehicle detail page,
scrapes all vehicle info, downloads images + the auction sheet, and
writes description.txt + vehicle_info.json into a
"<vehicle_code>-<vehicle_model>" folder.

SOLD vs UPCOMING is auto-detected per vehicle by checking for the
"Result Price" block on the page (see scraper_common.extract_result_price):

    - SOLD: the block is present and a price could be parsed. That JPY
      price is used directly (no Excel price needed at all) and
      price_calculator.py runs, producing price_breakdown.txt and the
      full LKR price in description.txt.
    - UPCOMING: no Result Price block yet (auction hasn't closed) -
      normal. description.txt gets only the vehicle's spec fields, no
      price/currency/WhatsApp-link lines, and no price_breakdown.txt.
    - SCRAPE ISSUE: the block exists but no price could be parsed out of
      it (page format changed, "No Sale" outcome, etc.) - flagged in
      SCRAPE_ISSUE.txt inside that vehicle's folder and in the results
      Excel, never guessed. Use recalculate_prices.py with a
      price_override to fix these once you've checked the site manually.

Re-running this script against the SAME Excel (and same vehicle_url
rows) on a later day - once auctions have closed - will pick up the new
Result Price automatically; no need to duplicate or rename the input
file for "yesterday's picks vs newly closed auctions".

SETUP:
    pip install playwright python-dotenv openpyxl pandas
    playwright install chromium

.env keys needed:
    AUTOFROMAUCTION_USERNAME=your_username
    AUTOFROMAUCTION_PASSWORD=your_password
    WHATSAPP_GROUP_LINK=https://chat.whatsapp.com/your_invite_code

Reference files needed alongside this script:
    inputSample.json, vehicleDetails.json  (see price_calculator.py)

STATUS: NOT yet run against the live site - test carefully.
"""

import sys
import pandas as pd

# Force UTF-8 output so special characters never crash the script when
# Windows redirects output to a file (cp1252 can't encode them, utf-8 can).
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
from pathlib import Path
from playwright.sync_api import sync_playwright

import price_calculator
import scraper_common as common

INPUT_EXCEL = "input_vehicles.xlsx"


def run():
    df = pd.read_excel(INPUT_EXCEL)
    df.columns = [c.strip().lower() for c in df.columns]
    assert "vehicle_url" in df.columns and "vehicle_model" in df.columns, \
        "Excel must have 'vehicle_url' and 'vehicle_model' columns"

    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=500)

        if Path(common.STORAGE_STATE_FILE).exists():
            print("Found saved session - reusing it.")
            context = browser.new_context(storage_state=common.STORAGE_STATE_FILE)
        else:
            print("No saved session found - will log in fresh.")
            context = browser.new_context()

        page = context.new_page()
        common.ensure_logged_in(page, context)

        for _, row in df.iterrows():
            vehicle_url = str(row["vehicle_url"]).strip()
            vehicle_model = str(row["vehicle_model"]).strip()
            vehicle_code = common.extract_vehicle_code(vehicle_url)

            print(f"Opening {vehicle_url} ({vehicle_model}) ...")
            detail_page = context.new_page()
            detail_page.goto(vehicle_url)
            detail_page.wait_for_load_state("networkidle")
            common.human_delay(1.0, 2.0)

            # scraped_maker_model as shown on the page is kept only as a
            # reference/sanity-check field - it is NOT used for the tax
            # calculation lookup, since the site's own naming can differ
            # from the exact dropdown name in vehicle_model. An empty
            # value here means the page didn't render at all (bad
            # URL/login) - skip this row.
            scraped_maker_model = common.extract_maker_model(detail_page)
            if not scraped_maker_model:
                print("  [FAILED] could not find vehicle name on page - "
                      "check login / URL validity")
                detail_page.close()
                results.append({**row, "status": "not_found"})
                continue

            auction_date = common.extract_auction_date(detail_page)
            detail_fields = common.extract_detail_fields(detail_page)

            price_str, had_result_block, parse_error = common.extract_result_price(detail_page)
            sold = had_result_block and not parse_error

            folder_path = Path(common.OUTPUT_DIR) / common.safe_folder_name(vehicle_code, vehicle_model)

            vehicle_info = {
                "vehicle_model": vehicle_model,
                "sold": sold,
                "price": price_str,  # None until sold
                "year": detail_fields.get("Registration Year", ""),
                "auction_date": auction_date,
                "auction_grade": detail_fields.get("Auction Grade", ""),
                "mileage": detail_fields.get("Mileage", ""),
                "color": detail_fields.get("Color", ""),
                "transmission": detail_fields.get("Transmission", ""),
                "steering": detail_fields.get("Steering Wheel", ""),
                "chassis": detail_fields.get("Chassis", ""),
                "equipment": detail_fields.get("Equipment", ""),
                "vehicle_code": vehicle_code,
                "vehicle_url": vehicle_url,
                # Reference-only fields as scraped from the site - not
                # used for the tax calculation:
                "scraped_maker_model": scraped_maker_model,
                "scraped_grade": detail_fields.get("Grade", ""),
                "scraped_fuel_type": detail_fields.get("Fuel Type", ""),
                "scraped_engine_cc": common.clean_price(detail_fields.get("Engine CC", "")),
            }

            image_urls = common.extract_images(detail_page)
            common.download_images(detail_page, image_urls, folder_path)

            sheet_url = common.extract_auction_sheet(detail_page)
            if sheet_url:
                common.download_auction_sheet(detail_page, sheet_url, folder_path)
            else:
                print("    ! no auction sheet found for this lot")

            detail_page.close()

            common.write_vehicle_info_file(folder_path, vehicle_info)
            common.write_description_file(folder_path, vehicle_info)

            if parse_error:
                # Result Price block exists but couldn't be parsed - flag
                # it clearly, never guess. Fix later with
                # recalculate_prices.py once you've checked the site.
                message = (
                    f"A 'Result Price' section was found on the vehicle page, "
                    f"but no numeric JPY price could be parsed out of it. "
                    f"This may mean the auction result was 'No Sale', or the "
                    f"page layout changed. Check {vehicle_url} manually, then "
                    f"use recalculate_prices.py with a price_override to fix "
                    f"this vehicle."
                )
                common.write_scrape_issue_file(folder_path, message)
                print(f"  ! SCRAPE ISSUE: {vehicle_model} - Result Price found "
                      f"but unparseable. See {folder_path / 'SCRAPE_ISSUE.txt'}")
                status = "price_scrape_error"

            elif sold:
                common.clear_scrape_issue_file(folder_path)
                price_calc_ok = price_calculator.process_vehicle_folder(folder_path)
                status = "ok" if price_calc_ok else "ok_price_calc_stopped"
                print(f"  [SOLD] {vehicle_model} ({vehicle_info['year']}) "
                      f"- Result Price {price_str} JPY - {len(image_urls)} image(s) "
                      f"-> {folder_path}")

            else:
                common.clear_scrape_issue_file(folder_path)
                common.clear_price_breakdown_file(folder_path)
                status = "upcoming"
                print(f"  [UPCOMING] {vehicle_model} ({vehicle_info['year']}) "
                      f"- not yet auctioned - {len(image_urls)} image(s) "
                      f"-> {folder_path}")

            results.append({
                "vehicle_url": vehicle_url,
                "vehicle_model": vehicle_model,
                "status": status,
                "images_downloaded": len(image_urls),
                "folder": str(folder_path),
            })

            common.human_delay(1.5, 3.0)

        browser.close()

    out_df = pd.DataFrame(results)
    out_df.to_excel(Path(common.OUTPUT_DIR) / "results_autofromauction.xlsx", index=False)
    print(f"\nDone. Results saved to {Path(common.OUTPUT_DIR) / 'results_autofromauction.xlsx'}")


if __name__ == "__main__":
    Path(common.OUTPUT_DIR).mkdir(exist_ok=True)
    run()
