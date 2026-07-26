"""
AutoFromAuction (www.autofromauction.com) - URL-Driven Image Downloader
--------------------------------------------------------------------------
Reads (vehicle_url, price) rows from Excel. For each one, logs in (once,
session reused after), navigates DIRECTLY to the given vehicle detail page
URL, scrapes all vehicle info, downloads images + the auction sheet, and
writes a description.txt into a "<vehicle_code>-<maker_model>" folder -
same overall output structure as the AutoAsta script.

This replaces the earlier lot-number-and-auction-house search flow, since
providing the vehicle detail URL directly sidesteps: the login form's
Angular render timing, the duplicate "Search" buttons, and auction-house
dropdown value mismatches.

After scraping each vehicle, this script also runs the local price
calculator (price_calculator.py) against the freshly-written
vehicleDetails.json, which writes price_breakdown.txt and updates
description.txt with the calculated LKR price (replacing the Excel price).

SETUP:
    pip install playwright python-dotenv openpyxl pandas
    playwright install chromium

.env keys needed:
    AUTOFROMAUCTION_USERNAME=your_username
    AUTOFROMAUCTION_PASSWORD=your_password
    CURRENCY_RATE_LKR=your_rate_here
    WHATSAPP_GROUP_LINK=https://chat.whatsapp.com/your_invite_code

Reference files needed alongside this script (see price_calculator.py):
    inputSample.json, vehicleWebValues.json, vehicleM3Values.json

STATUS: NOT yet run against the live site - test carefully. Known
uncertain points are marked with "ASSUMPTION" comments below.
"""

import sys
import os
import re
import json
import time
import random
import pandas as pd

# Force UTF-8 output so special characters (checkmarks, etc.) never crash
# the script when Windows redirects output to a file (cp1252 can't encode
# them, but utf-8 can).
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, Page

import price_calculator

load_dotenv()

USERNAME = os.getenv("AUTOFROMAUCTION_USERNAME")
PASSWORD = os.getenv("AUTOFROMAUCTION_PASSWORD")
CURRENCY_RATE_LKR = os.getenv("CURRENCY_RATE_LKR")
WHATSAPP_GROUP_LINK = os.getenv("WHATSAPP_GROUP_LINK")

_missing = [name for name, val in [
    ("AUTOFROMAUCTION_USERNAME", USERNAME),
    ("AUTOFROMAUCTION_PASSWORD", PASSWORD),
] if not val]
if _missing:
    raise RuntimeError(
        f"Missing required .env values: {', '.join(_missing)}. "
        f"Make sure .env is in the same folder as this script and has "
        f"these keys set."
    )

BASE_URL = "https://www.autofromauction.com"
LOGIN_URL = f"{BASE_URL}/"
HOME_URL = f"{BASE_URL}/home"

INPUT_EXCEL = "input_vehicles_autofromauction.xlsx"
OUTPUT_DIR = Path("downloads")
STORAGE_STATE_FILE = "auth_state_autofromauction.json"

IMAGE_ACCEPT_HEADER = "image/jpeg,image/png,image/*;q=0.8,*/*;q=0.5"


def human_delay(min_sec=1.0, max_sec=2.0):
    """Small random pause around page navigations - avoids a perfectly
    instant, obviously scripted page-to-page rhythm."""
    time.sleep(random.uniform(min_sec, max_sec))


def human_type(page: Page, selector: str, text: str):
    """Click then fill instantly (like a paste) - fast, still a real
    click+fill rather than anything unusual."""
    page.click(selector)
    page.fill(selector, text)


def login(page: Page):
    page.goto(LOGIN_URL)
    page.wait_for_load_state("networkidle")

    try:
        page.wait_for_selector("input[placeholder='User ID']", timeout=15000)
    except Exception as e:
        page.screenshot(path="af_login_page_not_found.png", full_page=True)
        print("USERNAME FIELD NOT FOUND:", repr(e))
        print("Current URL:", page.url)
        print("Page HTML snippet:")
        print(page.content()[:3000])
        raise

    # ASSUMPTION: the username field's real `name` attribute is malformed
    # (looks like a copy-pasted Angular binding snippet), so we target it by
    # placeholder text instead, which is stable.
    human_type(page, "input[placeholder='User ID']", USERNAME)
    human_type(page, "input[name='password']", PASSWORD)
    human_delay(1.0, 2.0)

    page.click("button:has-text('Sign In')")
    print("Clicked Sign In - waiting for 'Log Out' button to appear...")

    # This site doesn't change the URL after login - it stays on "/" and
    # just re-renders the navbar. Wait (up to 15s) for the "Log Out" button
    # to actually show up, rather than checking once immediately.
    try:
        # Don't assume "Log Out" is inside a real <button> tag - match by
        # visible text on any element, since it might be a styled <div>
        # or <a> instead (we hit this same trap with AutoAsta's "Reset").
        page.wait_for_selector("text=Log Out", timeout=15000)
        print("Login confirmed - 'Log Out' text found.")
    except Exception as e:
        page.screenshot(path="af_after_login_attempt.png", full_page=True)
        print("LOG OUT TEXT NEVER APPEARED:", repr(e))
        body_text = page.inner_text("body")
        print("Visible page text (first 1500 chars):")
        print(body_text[:1500])
        raise RuntimeError(
            "Login may have failed - 'Log Out' button not found after "
            "clicking Sign In. Check credentials, or inspect whether the "
            "site shows a validation error."
        )

    human_delay(1.0, 2.0)
    page.wait_for_load_state("networkidle")


def ensure_logged_in(page: Page, context):
    """Check whether we're already logged in (session reused); if not,
    log in fresh and save the session for next time."""
    page.goto(HOME_URL)
    page.wait_for_load_state("networkidle")
    if page.query_selector("text=Log Out") is None:
        print("Session expired or missing - logging in fresh.")
        login(page)
        context.storage_state(path=STORAGE_STATE_FILE)
        print(f"Session saved to {STORAGE_STATE_FILE} for next time.")
    else:
        print("Existing session confirmed valid.")


def extract_vehicle_code(vehicle_url: str) -> str:
    """Pull the trailing code from a URL like
    https://www.autofromauction.com/vehicle-details/IEM0250210324021
    -> 'IEM0250210324021'"""
    return vehicle_url.rstrip("/").split("/")[-1]


def extract_maker_model(detail_page: Page) -> str:
    """Vehicle name/title, e.g. 'SUZUKI WAGON R'."""
    el = detail_page.query_selector("h2.outfit-family-semibold")
    return el.inner_text().strip() if el else ""


def extract_auction_date(detail_page: Page) -> str:
    """The auction date/time (e.g. '2026-07-27 10:48') shown near the
    'Time Left' countdown. Found via regex over the visible page text
    rather than a fragile selector, since it's the most stable way to
    pull a date in this exact format wherever it appears on the page."""
    body_text = detail_page.inner_text("body")
    match = re.search(r"\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}", body_text)
    return match.group(0) if match else ""


def extract_detail_fields(detail_page: Page) -> dict:
    """Pull all labeled spec fields (Registration Year, Auction Grade,
    Mileage, Color, Manufacture Year, Engine CC, Steering Wheel, Fuel Type,
    Transmission, Chassis, Equipment, Grade) from the detail page's info
    card. Works by label text, so it's resilient to field re-ordering."""
    data = {}
    items = detail_page.query_selector_all(".innerlist-sec-wrap")
    for item in items:
        label_el = item.query_selector(".fs-10")
        value_el = item.query_selector(".fs-12")
        if label_el and value_el:
            label = label_el.inner_text().strip()
            value = value_el.inner_text().strip()
            data[label] = value
    return data


def extract_images(detail_page: Page) -> list[str]:
    """Full-resolution image URLs from the gallery - strips the `&w=320`
    resize parameter the thumbnails are served with."""
    imgs = detail_page.query_selector_all("gallery-thumb img.g-image-item")
    urls = []
    for img in imgs:
        src = img.get_attribute("src")
        if src:
            urls.append(re.sub(r"&w=\d+", "", src))
    return urls


def extract_auction_sheet(detail_page: Page) -> str | None:
    """Auction sheet image URL from the zoomable image viewer, same
    &w=320-stripping approach as the gallery images."""
    img = detail_page.query_selector("app-image-viewer img.zoomable-image")
    if img is None:
        return None
    src = img.get_attribute("src")
    if src is None:
        return None
    return re.sub(r"&w=\d+", "", src)


def clean_price(raw_price: str) -> str:
    """Strip currency labels, commas, and whitespace from a price string,
    keeping just the numeric value (digits + optional decimal point).
    e.g. 'LKR 5,200,000' -> '5200000', '5,200,000.50' -> '5200000.50'"""
    digits_and_dot = re.sub(r"[^\d.]", "", raw_price)
    return digits_and_dot


def _content_type_ext(content_type: str) -> str:
    ext_map = {
        "image/jpeg": "jfif",
        "image/jpg": "jfif",
        "image/png": "png",
        "image/webp": "webp",
    }
    return ext_map.get(content_type.split(";")[0].strip(), "jpg")


def download_images(page: Page, image_urls: list[str], folder: Path):
    folder.mkdir(parents=True, exist_ok=True)
    for i, url in enumerate(image_urls, start=1):
        if url.startswith("/"):
            url = BASE_URL + url

        response = page.request.get(url, headers={"Accept": IMAGE_ACCEPT_HEADER})
        if not response.ok:
            print(f"    ! failed to download image {i} ({response.status})")
            continue

        ext = _content_type_ext(response.headers.get("content-type", ""))
        (folder / f"image_{i}.{ext}").write_bytes(response.body())


def download_auction_sheet(page: Page, sheet_url: str, folder: Path):
    folder.mkdir(parents=True, exist_ok=True)

    url = sheet_url
    if url.startswith("/"):
        url = BASE_URL + url

    response = page.request.get(url, headers={"Accept": IMAGE_ACCEPT_HEADER})
    if not response.ok:
        print(f"    ! failed to download auction sheet ({response.status})")
        return

    ext = _content_type_ext(response.headers.get("content-type", ""))
    (folder / f"auction_sheet.{ext}").write_bytes(response.body())


def write_description_file(folder: Path, row_data: dict):
    """Write description.txt. The 'Grade' spec-card field (e.g.
    'Hybrid ZX') is used as the trim descriptor after the maker/model
    name, matching the desired 'SUZUKI WAGON R Hybrid ZX' style.

    NOTE: the price written here is the raw Excel price. It gets
    overwritten with the calculated LKR price by
    price_calculator.process_vehicle_folder(), called later in the loop."""
    description = (
        f"{row_data['maker_model']} {row_data['grade_trim']} {row_data['year']} - {row_data['price']}\n"
        f"Auction Date: {row_data['auction_date']}\n"
        f"Color: {row_data['color']}\n"
        f"Mileage: {row_data['mileage']}\n"
        f"Auction Grade: {row_data['auction_grade']}\n"
        f"Currency Rate: {CURRENCY_RATE_LKR} LKR (Today's government customs rate)\n"
        f"Includes All Taxes & Import Fees\n"
        f"Follow this link to join our WhatsApp group: {WHATSAPP_GROUP_LINK}\n"
    )
    (folder / "description.txt").write_text(description, encoding="utf-8")


def write_market_price_file(folder: Path, row_data: dict):
    """Write a JSON reference file for feeding into the price calculator.
    fuel_type, engine, chassis, and price come first (price explicitly as
    the 4th key), followed by the rest of the vehicle's details."""
    data = {
        "fuel_type": row_data["fuel_type"],
        "engine": row_data["engine_cc"],
        "chassis": row_data["chassis"],
        "price": row_data["price"],
        "maker_model": row_data["maker_model"],
        "grade_trim": row_data["grade_trim"],
        "year": row_data["year"],
        "auction_grade": row_data["auction_grade"],
        "mileage": row_data["mileage"],
        "color": row_data["color"],
        "transmission": row_data["transmission"],
        "steering": row_data["steering"],
        "equipment": row_data["equipment"],
        "auction_date": row_data["auction_date"],
        "vehicle_code": row_data["vehicle_code"],
        "vehicle_url": row_data["vehicle_url"],
    }
    with open(folder / "vehicleDetails.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def safe_folder_name(vehicle_code: str, maker_model: str) -> str:
    raw = f"{vehicle_code}-{maker_model}"
    return "".join(c for c in raw if c not in '<>:"/\\|?*').strip()


def run():
    df = pd.read_excel(INPUT_EXCEL)
    df.columns = [c.strip().lower() for c in df.columns]
    assert "vehicle_url" in df.columns and "price" in df.columns, \
        "Excel must have 'vehicle_url' and 'price' columns"

    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=500)

        if Path(STORAGE_STATE_FILE).exists():
            print("Found saved session - reusing it.")
            context = browser.new_context(storage_state=STORAGE_STATE_FILE)
        else:
            print("No saved session found - will log in fresh.")
            context = browser.new_context()

        page = context.new_page()
        ensure_logged_in(page, context)

        for _, row in df.iterrows():
            vehicle_url = str(row["vehicle_url"]).strip()
            price = clean_price(str(row["price"]).strip())
            vehicle_code = extract_vehicle_code(vehicle_url)

            print(f"Opening {vehicle_url} ...")
            detail_page = context.new_page()
            detail_page.goto(vehicle_url)
            detail_page.wait_for_load_state("networkidle")
            human_delay(1.0, 2.0)

            maker_model = extract_maker_model(detail_page)
            if not maker_model:
                print("  [FAILED] could not find vehicle name on page - "
                      "check login / URL validity")
                detail_page.close()
                results.append({**row, "status": "not_found", "images_downloaded": 0})
                continue

            auction_date = extract_auction_date(detail_page)
            detail_fields = extract_detail_fields(detail_page)

            folder_path = OUTPUT_DIR / safe_folder_name(vehicle_code, maker_model)

            row_data = {
                "maker_model": maker_model,
                "grade_trim": detail_fields.get("Grade", ""),
                "year": detail_fields.get("Registration Year", ""),
                "auction_date": auction_date,
                "color": detail_fields.get("Color", ""),
                "mileage": detail_fields.get("Mileage", ""),
                "auction_grade": detail_fields.get("Auction Grade", ""),
                "engine_cc": re.sub(r"[^\d]", "", detail_fields.get("Engine CC", "")),
                "fuel_type": detail_fields.get("Fuel Type", ""),
                "transmission": detail_fields.get("Transmission", ""),
                "steering": detail_fields.get("Steering Wheel", ""),
                "chassis": detail_fields.get("Chassis", ""),
                "equipment": detail_fields.get("Equipment", ""),
                "vehicle_code": vehicle_code,
                "vehicle_url": vehicle_url,
                "price": price,
            }

            image_urls = extract_images(detail_page)
            download_images(detail_page, image_urls, folder_path)

            sheet_url = extract_auction_sheet(detail_page)
            if sheet_url:
                download_auction_sheet(detail_page, sheet_url, folder_path)
            else:
                print("    ! no auction sheet found for this lot")

            detail_page.close()

            write_description_file(folder_path, row_data)
            write_market_price_file(folder_path, row_data)

            # --- Price calculation step ---
            # Runs the JPY->LKR import-cost formula against the
            # vehicleDetails.json we just wrote, writes price_breakdown.txt,
            # and overwrites description.txt's price line with the
            # calculated value. If a reference value is missing or out of
            # range, this stops itself and flags it in price_breakdown.txt
            # (and in description.txt) rather than guessing - it never
            # raises, so the scraper loop always continues to the next lot.
            price_calc_ok = price_calculator.process_vehicle_folder(folder_path)

            sheet_note = "+ auction sheet" if sheet_url else "(no auction sheet)"
            print(f"  [OK] {maker_model} {row_data['grade_trim']} ({row_data['year']}) "
                  f"- {len(image_urls)} image(s) {sheet_note} + description.txt -> {folder_path}")
            results.append({
                **row,
                "status": "ok" if price_calc_ok else "ok_price_calc_stopped",
                "images_downloaded": len(image_urls),
                "folder": str(folder_path),
            })

            human_delay(1.5, 3.0)

        browser.close()

    out_df = pd.DataFrame(results)
    out_df.to_excel(OUTPUT_DIR / "results_autofromauction.xlsx", index=False)
    print(f"\nDone. Results saved to {OUTPUT_DIR / 'results_autofromauction.xlsx'}")


if __name__ == "__main__":
    OUTPUT_DIR.mkdir(exist_ok=True)
    run()
