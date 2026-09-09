"""
scraper_common.py
------------------
Shared helpers for the AutoFromAuction workflow: login/session handling,
page scraping (including the "Result Price" sold-price detection),
image/auction-sheet downloads, and the two per-vehicle output files
(vehicle_info.json, description.txt).

Used by:
    download_and_price.py   - the full scrape-and-download flow
    recalculate_prices.py   - fixes an existing vehicle's model/price
                               without touching the browser at all

SOLD vs UPCOMING detection
---------------------------
A vehicle detail page only shows a "Result Price" block once its auction
has actually closed. extract_result_price() checks for that block and
returns one of three states:

    (price_str, True,  False)  - sold: block found, price parsed OK
    (None,      False, False)  - upcoming: no block yet (normal, no error)
    (None,      True,  True)   - the block exists but no number could be
                                  parsed out of it (page format changed,
                                  "No Sale" outcome, etc.) - this is
                                  flagged as a scrape issue, never guessed.

Vehicles in the second and third states get no price, no price
breakdown, and a shorter description.txt (see write_description_file).
"""

import os
import re
import json
import time
import random
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, Page

load_dotenv()

USERNAME = os.getenv("AUTOFROMAUCTION_USERNAME")
PASSWORD = os.getenv("AUTOFROMAUCTION_PASSWORD")
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


def _load_jpy_rate() -> str:
    """Reads jpyRate from inputSample.json - the same reference file
    price_calculator.py uses for the tax calculation - so the 'Currency
    Rate' line in description.txt always matches the rate actually used,
    with no separate .env value to keep in sync."""
    try:
        with open(Path(__file__).resolve().parent / "inputSample.json", "r", encoding="utf-8") as f:
            data = json.load(f)
        return str(data.get("jpyRate", ""))
    except Exception as e:
        print(f"    ! could not read jpyRate from inputSample.json: {e}")
        return ""


CURRENCY_RATE_LKR = _load_jpy_rate()

BASE_URL = "https://www.autofromauction.com"
LOGIN_URL = f"{BASE_URL}/"
HOME_URL = f"{BASE_URL}/home"

OUTPUT_DIR = Path("downloads")
STORAGE_STATE_FILE = "auth_state_autofromauction.json"
IMAGE_ACCEPT_HEADER = "image/jpeg,image/png,image/*;q=0.8,*/*;q=0.5"


# ---------------------------------------------------------------------
# Login / session
# ---------------------------------------------------------------------

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


# ---------------------------------------------------------------------
# Page scraping
# ---------------------------------------------------------------------

def extract_vehicle_code(vehicle_url: str) -> str:
    """Pull the trailing code from a URL like
    https://www.autofromauction.com/vehicle-details/IEM0250210324021
    -> 'IEM0250210324021'. Pure string parsing - no browser needed, so
    this is also used by recalculate_prices.py."""
    return vehicle_url.rstrip("/").split("/")[-1]


def extract_maker_model(detail_page: Page) -> str:
    """Vehicle name/title, e.g. 'SUZUKI WAGON R'. Used only as a
    reference/sanity-check field - NOT for the tax calculation lookup,
    which uses the exact Excel vehicle_model instead."""
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


def extract_result_price(detail_page: Page):
    """Detects whether this vehicle has already been auctioned, by
    looking for the 'Result Price' block that only appears once an
    auction has closed. Returns a (price_str, had_result_block,
    parse_error) tuple:

        ("927000", True,  False)  - SOLD: block found, price parsed OK.
        (None,     False, False)  - UPCOMING: no block yet. Normal, not
                                     an error - the auction just hasn't
                                     happened.
        (None,     True,  True)   - block exists but no number could be
                                     parsed, OR the exact label appears
                                     more than once on the page (page
                                     format changed, "No Sale" outcome,
                                     an unexpected duplicate section,
                                     etc.) - flag this, never guess a
                                     price.

    IMPORTANT: matching is case-sensitive and requires the literal colon
    ("Result Price:") - the site's pages were found to contain several
    other occurrences of "result price" elsewhere (lowercase/no colon -
    e.g. a sort/filter option, or similar-vehicles text), which an
    earlier case-insensitive version could match by mistake, grabbing an
    unrelated number. Only the genuine sold-price label uses this exact
    capitalization and punctuation."""
    body_text = detail_page.inner_text("body")

    if "Result Price:" not in body_text:
        return None, False, False

    matches = re.findall(r"Result Price:\s*([\d,]+(?:\.\d+)?)\s*JPY", body_text)

    if len(matches) == 1:
        return clean_price(matches[0]), True, False

    if len(matches) > 1:
        # More than one exact "Result Price:" label with a parseable
        # number - ambiguous, don't guess which one is correct.
        return None, True, True

    # Exact label text is present but not immediately followed by a
    # parseable "<number> JPY" (e.g. "No Sale") - flag, don't guess.
    return None, True, True


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
    return re.sub(r"[^\d.]", "", raw_price)


# ---------------------------------------------------------------------
# Downloads
# ---------------------------------------------------------------------

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


# ---------------------------------------------------------------------
# Per-vehicle output files
# ---------------------------------------------------------------------

def safe_folder_name(vehicle_code: str, vehicle_model: str) -> str:
    raw = f"{vehicle_code}-{vehicle_model}"
    return "".join(c for c in raw if c not in '<>:"/\\|?*').strip()


def _date_only(auction_date: str) -> str:
    """Auction dates sometimes include a time component (e.g.
    '2026-08-31 11:39'). Only the date is shown in description.txt -
    this strips any trailing time, falling back to the original string
    unchanged if it doesn't start with a YYYY-MM-DD pattern."""
    match = re.match(r"(\d{4}-\d{2}-\d{2})", (auction_date or "").strip())
    return match.group(1) if match else (auction_date or "")


def write_description_file(folder: Path, info: dict):
    """Writes description.txt from a vehicle_info-shaped dict. Used by
    BOTH download_and_price.py (right after scraping) and
    recalculate_prices.py (regenerating from a corrected vehicle_info.json)
    - always rebuilds the file from scratch, so sold/upcoming status is
    reflected consistently either way.

    SOLD (info['sold'] is True and info.get('price') is truthy): full
    listing with price | *price*, Currency Rate, the all-inclusive note,
    and the WhatsApp link. The price shown here is the raw JPY figure -
    price_calculator.process_vehicle_folder() overwrites it with the
    calculated LKR total right after this is called.

    UPCOMING or unresolved scrape issue: shorter listing with just the
    vehicle's spec fields - no price line, no Currency Rate line, no
    all-inclusive note, no WhatsApp link, per instructions."""
    sold = bool(info.get("sold")) and info.get("price") not in (None, "", "None")
    auction_date = _date_only(info.get("auction_date", ""))

    if sold:
        sinhala_note = (
            f"මෙය {auction_date} දින ජපාන වෙන්දේසියේදී අලෙවි වූ මෝටර් රථයකි.\n"
            f"මෙහි දැක්වෙන්නේ මෙම වාහනය ලංකාවට ආනයනය කිරීම සඳහා අවශ්‍ය වන සම්පූර්ණ පිරිවැයයි.\n"
            f"මෙමගින් හෙට දිනයේ වෙන්දේසියේදී අලෙවි වීමට නියමිත වාහනවල මිල "
            f"ගණන් පිළිබඳව ඔබට නිවැරදි අවබෝධයක් ලබාගත හැකි වේ.\nඅද දින අලෙවි වූ "
            f"වාහනවල සත්‍ය මිල ගණන් මෙන්ම, හෙට දින ජපාන වෙන්දේසියට එක්වන වාහනවල "
            f"විස්තර දැනගැනීම සඳහා අපගේ WhatsApp සමූහයට එක්වන්න."
        )
        description = (
            f"🚗 {info['vehicle_model']} {info.get('year', '')} | *{info['price']}*\n\n"
            f"📅 Auction Date: {auction_date}\n"
            f"🎨 Color: {info.get('color', '')}\n"
            f"🛣️ Mileage: {info.get('mileage', '')}\n"
            f"⭐ Auction Grade: {info.get('auction_grade', '')}\n"
            f"💱 Currency Rate: {CURRENCY_RATE_LKR} LKR (Local customs rate)\n"
            f"✅ All-inclusive take home price\n\n"
            f"{sinhala_note}\n\n"
            f"📲 Follow this link to join our WhatsApp group: {WHATSAPP_GROUP_LINK}\n"
        )
    else:
        description = (
            f"🚗 {info['vehicle_model']} {info.get('year', '')}\n\n"
            f"📅 Auction Date: {auction_date}\n"
            f"🎨 Color: {info.get('color', '')}\n"
            f"🛣️ Mileage: {info.get('mileage', '')}\n"
            f"⭐ Auction Grade: {info.get('auction_grade', '')}\n"
        )
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "description.txt").write_text(description, encoding="utf-8")


def write_vehicle_info_file(folder: Path, info: dict):
    """Writes vehicle_info.json as-is. 'vehicle_model' and 'price' are the
    authoritative fields price_calculator.py uses (matched exactly
    against vehicleDetails.json). 'sold' controls whether
    description.txt includes pricing. Fields prefixed 'scraped_' are
    reference/debugging only - never used in the calculation."""
    folder.mkdir(parents=True, exist_ok=True)
    with open(folder / "vehicle_info.json", "w", encoding="utf-8") as f:
        json.dump(info, f, indent=2, ensure_ascii=False)


def write_scrape_issue_file(folder: Path, message: str):
    """Writes a clear, internal-only note when the Result Price block was
    found but couldn't be parsed - never surfaced in description.txt
    (which is customer-facing), only here for whoever reviews the
    listing before it's used."""
    (folder / "SCRAPE_ISSUE.txt").write_text(
        "SCRAPE ISSUE - MANUAL REVIEW NEEDED\n"
        + "=" * 60 + "\n"
        + message + "\n",
        encoding="utf-8",
    )


def clear_scrape_issue_file(folder: Path):
    issue_path = folder / "SCRAPE_ISSUE.txt"
    if issue_path.exists():
        issue_path.unlink()


def clear_price_breakdown_file(folder: Path):
    """Removes a stale price_breakdown.txt, e.g. if a vehicle_model
    correction reverts a vehicle to a state where it no longer has a
    valid price to calculate."""
    breakdown_path = folder / "price_breakdown.txt"
    if breakdown_path.exists():
        breakdown_path.unlink()