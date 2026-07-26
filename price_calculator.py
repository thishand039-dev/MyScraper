"""
price_calculator.py
--------------------
Calculates the landed local-market (LKR) price of a vehicle from its
auction unit price (JPY), using the fixed import-cost formula, and writes
a full price_breakdown.txt into the vehicle's output folder. Also updates
that folder's description.txt with the calculated price (replacing the
raw Excel price).

INPUT FILES (all expected in the same directory as this script, i.e. next
to the scraper script that imports it):

    inputSample.json       - fixed constants (rates & fees), e.g.:
        {
          "jpyRate": 2.1126,
          "velTax": 15000,
          "bankingFees": 55000,
          "clearingFees": 45000,
          "portFees": 35000,
          "stampFees": 1750,
          "serviceChargers": "200000"
        }

    vehicleWebValues.json  - manually maintained "web value" (x) per
        vehicle, keyed by "<maker_model> <grade_trim>", e.g.:
        { "SUZUKI WAGON R Hybrid ZX": 1700000 }

    vehicleM3Values.json   - manually maintained M3/volume value per
        vehicle, same key format, e.g.:
        { "SUZUKI WAGON R Hybrid ZX": 8.55 }

PER-VEHICLE INPUT (read from the vehicle's own output folder):
    vehicleDetails.json     - written by the scraper (price, fuel_type,
                               engine_cc, maker_model, grade_trim, etc.)

OUTPUTS (written into the same vehicle folder):
    price_breakdown.txt      - full JPY->LKR cost breakdown, values shown
                                to 2 decimal places.
    description.txt          - price line updated (via the caller passing
                                us the same row_data dict used to build
                                the original description, plus our own
                                write_description_file-compatible call).

STOP-ON-ISSUE BEHAVIOUR:
    If any required reference value is missing, or a value falls outside
    every defined bracket (M3 out of range, unsupported fuel/CC
    combination, unit price with no undervalue-amount bracket, etc.),
    calculation stops at that point. Whatever was already computed is
    still written to price_breakdown.txt, followed by a clear message
    describing what stopped it. This function never raises/crashes the
    caller's loop - it always returns a (success: bool, breakdown_text)
    tuple.
"""

import json
import re
from pathlib import Path

CONFIG_DIR = Path(__file__).resolve().parent
INPUT_VALUES_FILE = CONFIG_DIR / "inputSample.json"
WEB_VALUES_FILE = CONFIG_DIR / "vehicleWebValues.json"
M3_VALUES_FILE = CONFIG_DIR / "vehicleM3Values.json"


class CalculationStopped(Exception):
    """Raised when a required reference value is missing or out of range.
    The message is written verbatim into price_breakdown.txt."""


def _load_json(path: Path) -> dict:
    if not path.exists():
        raise CalculationStopped(f"Required reference file not found: {path.name}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _lookup_key(vehicle_details: dict) -> str:
    """Matches entries in vehicleWebValues.json / vehicleM3Values.json,
    which are keyed by '<maker_model> <grade_trim>' e.g.
    'SUZUKI WAGON R Hybrid ZX'."""
    maker_model = str(vehicle_details.get("maker_model", "")).strip()
    grade_trim = str(vehicle_details.get("grade_trim", "")).strip()
    return f"{maker_model} {grade_trim}".strip()


# ---------------------------------------------------------------------
# Bracket tables (JPY unless noted). Clean cutoffs: [low, next_low).
# ---------------------------------------------------------------------

def agent_fee(b1: float) -> float:
    """B3"""
    if b1 < 1_000_000:
        return 140_000.0
    elif b1 < 2_000_000:
        return 150_000.0
    elif b1 < 3_000_000:
        return 160_000.0
    elif b1 < 4_000_000:
        return 280_000.0
    else:
        return b1 * 0.10


def undervalue_amount(b1: float) -> float:
    """B9. No bracket is defined for unit price >= 4,000,000 - stop."""
    if b1 < 1_000_000:
        return 120_000.0
    elif b1 < 2_000_000:
        return 130_000.0
    elif b1 < 3_000_000:
        return 140_000.0
    elif b1 < 4_000_000:
        return 240_000.0
    else:
        raise CalculationStopped(
            "Undervalue Amount (B9): unit price is 4,000,000 JPY or higher, "
            "which has no defined bracket. Calculation stopped - needs "
            "manual review."
        )


def freight_fee(b5_m3: float) -> float:
    """B7. Table only covers 0 - 16.00 M3."""
    brackets = [
        (0.0, 10.00, 110_000.0),
        (10.01, 11.00, 120_000.0),
        (11.01, 12.00, 130_000.0),
        (12.01, 13.00, 140_000.0),
        (13.01, 14.00, 150_000.0),
        (14.01, 15.00, 160_000.0),
        (15.01, 16.00, 190_000.0),
    ]
    for low, high, fee in brackets:
        if low <= b5_m3 <= high:
            return fee
    raise CalculationStopped(
        f"Freight (B7): M3 value {b5_m3} is not in the supported range "
        f"(0 - 16.00 M3). Calculation stopped."
    )


def xid_amount(fuel_type: str, engine_cc: int) -> float:
    """XID. 660cc or less is a fixed value regardless of fuel type.
    ASSUMPTION: table boundaries overlap at 1000/1300/1500cc (e.g. Petrol
    lists both '0-1000' and '1000-1299'); brackets are checked in order
    and the first (lower) match wins at an exact boundary value."""
    fuel = (fuel_type or "").strip().lower()

    if engine_cc <= 660:
        return 1_992_000.0

    if fuel == "petrol":
        brackets = [
            (0, 1000, 2450.0),
            (1000, 1299, 3850.0),
            (1300, 1499, 4450.0),
            (1500, 1599, 5150.0),
        ]
    elif fuel == "hybrid":
        brackets = [
            (0, 1299, 2750.0),
            (1300, 1499, 3450.0),
            (1500, 1599, 4800.0),
        ]
    else:
        raise CalculationStopped(
            f"XID: unsupported fuel type '{fuel_type}' - needs manual "
            f"review. Calculation stopped."
        )

    for low, high, rate in brackets:
        if low <= engine_cc <= high:
            return rate * engine_cc

    raise CalculationStopped(
        f"XID: unsupported fuel type/CC combination ('{fuel_type}', "
        f"{engine_cc}cc) - needs manual review. Calculation stopped."
    )


# ---------------------------------------------------------------------
# Main calculation
# ---------------------------------------------------------------------

def calculate(vehicle_details: dict) -> dict:
    """Runs the full formula. Returns a dict of every intermediate value
    plus 'total_vehicle_cost'. Raises CalculationStopped part-way through
    if a required reference value is missing/out of range - the caller
    (process_vehicle_folder) catches this and writes whatever is in
    `partial` so far, so keep this dict updated at every step."""

    partial = {}

    inputs = _load_json(INPUT_VALUES_FILE)
    web_values = _load_json(WEB_VALUES_FILE)
    m3_values = _load_json(M3_VALUES_FILE)

    key = _lookup_key(vehicle_details)

    b1 = float(vehicle_details.get("price", 0) or 0)
    b2 = 0.0  # Auction House Fee - defaulted to 0 until a per-auction-house
              # json is added later.
    partial["B1_unit_price"] = b1
    partial["B2_auction_house_fee"] = b2

    b3 = agent_fee(b1)
    partial["B3_agent_fee"] = b3

    b4 = 30_000.0
    partial["B4_inspection"] = b4

    if key not in m3_values:
        raise _stopped_with(
            partial,
            f"Vehicle M3 value not found for '{key}' in vehicleM3Values.json. "
            f"Calculation stopped."
        )
    b5 = float(m3_values[key])
    partial["B5_m3"] = b5

    b6 = (b1 + b2 + b3) * b5 * 0.001 * 1.05
    partial["B6_insurance"] = b6

    try:
        b7 = freight_fee(b5)
    except CalculationStopped as e:
        raise _stopped_with(partial, str(e))
    partial["B7_freight"] = b7

    b8 = b1 + b2 + b3 + b4 + b6 + b7
    partial["B8_full_lc_cif"] = b8

    try:
        b9 = undervalue_amount(b1)
    except CalculationStopped as e:
        raise _stopped_with(partial, str(e))
    partial["B9_undervalue_amount"] = b9

    b10 = b8 - b9
    partial["B10_proforma_invoice_cif"] = b10

    if key not in web_values:
        raise _stopped_with(
            partial,
            f"Vehicle web value not found for '{key}' in "
            f"vehicleWebValues.json. Calculation stopped."
        )
    web_value = float(web_values[key])
    y = web_value * (100 / 110) * (85 / 100)
    partial["web_value_x"] = web_value
    partial["Y_web_value_adjusted"] = y

    tax_base = max(y, b10)
    partial["tax_base"] = tax_base

    jpy_rate = float(inputs["jpyRate"])
    partial["jpy_rate"] = jpy_rate

    cif = jpy_rate * tax_base
    cid = cif * 0.45
    partial["CIF"] = cif
    partial["CID"] = cid

    fuel_type = vehicle_details.get("fuel_type", "")
    engine_cc_raw = vehicle_details.get("engine_cc", "")
    try:
        engine_cc = int(re.sub(r"[^\d]", "", str(engine_cc_raw)))
    except ValueError:
        raise _stopped_with(
            partial,
            f"Could not parse engine CC from '{engine_cc_raw}'. "
            f"Calculation stopped."
        )

    try:
        xid = xid_amount(fuel_type, engine_cc)
    except CalculationStopped as e:
        raise _stopped_with(partial, str(e))
    partial["XID"] = xid

    sscl = (cif * 1.10 + cid + xid) * 0.025
    vat = (cif * 1.10 + cid + xid) * 0.18
    vel = float(inputs["velTax"])
    partial["SSCL_TAX"] = sscl
    partial["VAT"] = vat
    partial["VEL"] = vel

    total_tax = cid + xid + sscl + vat + vel
    partial["total_tax"] = total_tax

    b33 = b9 * jpy_rate * 1.10
    vehicle_value = b10 * jpy_rate + b33
    partial["B33_undervalue_and_charges"] = b33
    partial["vehicle_value"] = vehicle_value

    bank_port_clearing = (
        float(inputs["bankingFees"]) + float(inputs["clearingFees"])
        + float(inputs["portFees"]) + float(inputs["stampFees"])
    )
    partial["bank_port_clearing"] = bank_port_clearing

    landing_cost = total_tax + vehicle_value + bank_port_clearing
    partial["landing_cost"] = landing_cost

    service_charges = float(inputs["serviceChargers"])
    partial["service_charges"] = service_charges

    total_vehicle_cost = landing_cost + service_charges
    partial["total_vehicle_cost"] = total_vehicle_cost

    return partial


def _stopped_with(partial: dict, message: str) -> CalculationStopped:
    exc = CalculationStopped(message)
    exc.partial = dict(partial)
    return exc


# ---------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------

def _money(v: float) -> str:
    return f"{v:,.2f}"


def round_to_lakhs(value_lkr: float) -> str:
    """Rounds to the nearest 25,000 LKR, then expresses in Lakhs
    (divide by 100,000). E.g. 7,257,533.756 -> nearest 25,000 is
    7,250,000 -> 72.5 Lakhs."""
    rounded_25k = round(value_lkr / 25_000) * 25_000
    lakhs = rounded_25k / 100_000
    s = f"{lakhs:.2f}".rstrip("0").rstrip(".")
    return s if s else "0"


def format_breakdown(vehicle_details: dict, result: dict, stopped_message: str = None) -> str:
    key = _lookup_key(vehicle_details)
    lines = []
    lines.append(f"Price Breakdown - {key}")
    lines.append(f"Vehicle Code: {vehicle_details.get('vehicle_code', '')}")
    lines.append("=" * 60)

    if "CID" in result:
        lines.append(f"CID:                          {_money(result['CID'])}")
    if "XID" in result:
        lines.append(f"XID:                          {_money(result['XID'])}")
    if "SSCL_TAX" in result:
        lines.append(f"SSCL TAX (2.5%):              {_money(result['SSCL_TAX'])}")
    if "VAT" in result:
        lines.append(f"VAT (18%):                    {_money(result['VAT'])}")
    if "VEL" in result:
        lines.append(f"VEL:                          {_money(result['VEL'])}")
    lines.append("LXT:                          Not included (see note below)")
    if "total_tax" in result:
        lines.append(f"Total Tax:                    {_money(result['total_tax'])}")
    if "vehicle_value" in result:
        lines.append(f"Vehicle Value:                {_money(result['vehicle_value'])}")
    if "bank_port_clearing" in result:
        lines.append(f"Bank, Port & Clearing:        {_money(result['bank_port_clearing'])}")
    if "landing_cost" in result:
        lines.append(f"Landing Cost:                 {_money(result['landing_cost'])}")
    if "service_charges" in result:
        lines.append(f"Service Charges:              {_money(result['service_charges'])}")
    if "total_vehicle_cost" in result:
        lines.append(f"Total Vehicle Cost:           {_money(result['total_vehicle_cost'])}")

    lines.append("=" * 60)
    lines.append("Note: Luxury Tax (LXT) is not included in this calculation.")

    if stopped_message:
        lines.append("")
        lines.append("!" * 60)
        lines.append("CALCULATION STOPPED - MANUAL REVIEW NEEDED")
        lines.append(stopped_message)
        lines.append("!" * 60)

    return "\n".join(lines) + "\n"


def update_description_price(description_path: Path, new_price_line: str):
    """Replaces the first line of description.txt (the
    '<maker_model> <grade_trim> <year> - <price>' line) with the same
    text but the calculated price, leaving all other lines untouched."""
    if not description_path.exists():
        return
    text = description_path.read_text(encoding="utf-8")
    lines = text.splitlines()
    if not lines:
        return
    first_line = lines[0]
    if " - " in first_line:
        prefix = first_line.rsplit(" - ", 1)[0]
        lines[0] = f"{prefix} - {new_price_line}"
    else:
        lines[0] = f"{first_line} - {new_price_line}"
    description_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


# ---------------------------------------------------------------------
# Public entry point used by the scraper
# ---------------------------------------------------------------------

def process_vehicle_folder(folder: Path) -> bool:
    """Reads <folder>/vehicleDetails.json, runs the calculation, writes
    <folder>/price_breakdown.txt, and updates <folder>/description.txt's
    price line. Returns True on a full successful calculation, False if
    it stopped early (details are in price_breakdown.txt either way)."""
    details_path = folder / "vehicleDetails.json"
    if not details_path.exists():
        print(f"    ! price_calculator: vehicleDetails.json not found in {folder}")
        return False

    with open(details_path, "r", encoding="utf-8") as f:
        vehicle_details = json.load(f)

    try:
        result = calculate(vehicle_details)
        breakdown_text = format_breakdown(vehicle_details, result)
        (folder / "price_breakdown.txt").write_text(breakdown_text, encoding="utf-8")

        price_label = f"LKR {round_to_lakhs(result['total_vehicle_cost'])} Lakhs"
        update_description_price(folder / "description.txt", price_label)

        print(f"    Price calculated: {price_label} "
              f"(Total Vehicle Cost: {_money(result['total_vehicle_cost'])} LKR)")
        return True

    except CalculationStopped as e:
        partial = getattr(e, "partial", {})
        breakdown_text = format_breakdown(vehicle_details, partial, stopped_message=str(e))
        (folder / "price_breakdown.txt").write_text(breakdown_text, encoding="utf-8")

        update_description_price(
            folder / "description.txt",
            "PRICE CALCULATION STOPPED - see price_breakdown.txt"
        )
        print(f"    ! price calculation stopped: {e}")
        return False
