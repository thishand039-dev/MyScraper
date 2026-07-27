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

    inputSample.json     - fixed constants (rates & fees), e.g.:
        {
          "jpyRate": 2.1126,
          "velTax": 15000,
          "bankingFees": 55000,
          "clearingFees": 45000,
          "portFees": 35000,
          "stampFees": 1750,
          "serviceChargers": "200000"
        }

    vehicleDetails.json  - manually maintained reference config, one entry
        per exact vehicle model name as it appears in the site's model
        dropdown (this is also what goes in the Excel 'vehicle_model'
        column, so the two match exactly), e.g.:
        {
          "SUZUKI WAGON R ZX Hybrid": {
            "webValue": 1736900,
            "fuelType": "hybrid",
            "engineCc": 660,
            "m3Value": 8.26
          }
        }
        This is the ONLY source of fuelType/engineCc/webValue/m3Value used
        for the tax calculation - scraped site fields are NOT used here,
        since different auction sites/pages can label the same vehicle
        differently.

PER-VEHICLE INPUT (read from the vehicle's own output folder):
    vehicle_info.json    - written by the scraper. Must contain at least
                            "vehicle_model" (the exact Excel dropdown
                            value) and "price" (the Excel unit price in
                            JPY). May also carry scraped fields for
                            reference/debugging only.

OUTPUTS (written into the same vehicle folder):
    price_breakdown.txt   - full JPY->LKR cost breakdown, values shown to
                             2 decimal places.
    description.txt       - price line updated in place.

STOP-ON-ISSUE BEHAVIOUR:
    If the vehicle_model isn't found in vehicleDetails.json, or any
    computed value falls outside every defined bracket (M3 out of range,
    unsupported fuel/CC combination, unit price with no undervalue-amount
    bracket, etc.), calculation stops at that point. Whatever was already
    computed is still written to price_breakdown.txt, followed by a clear
    message describing what stopped it. This function never
    raises/crashes the caller's loop - it always returns a bool.
"""

import json
from pathlib import Path

CONFIG_DIR = Path(__file__).resolve().parent
INPUT_VALUES_FILE = CONFIG_DIR / "inputSample.json"
VEHICLE_DETAILS_CONFIG_FILE = CONFIG_DIR / "vehicleDetails.json"


class CalculationStopped(Exception):
    """Raised when a required reference value is missing or out of range.
    The message is written verbatim into price_breakdown.txt."""


def _load_json(path: Path) -> dict:
    if not path.exists():
        raise CalculationStopped(f"Required reference file not found: {path.name}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _stopped_with(partial: dict, message: str) -> CalculationStopped:
    exc = CalculationStopped(message)
    exc.partial = dict(partial)
    return exc


# ---------------------------------------------------------------------
# Bracket tables (JPY unless noted). Clean cutoffs: [low, next_low).
# Named "calculate_..." (rather than e.g. "agent_fee") so they don't
# shadow the descriptive result variables of the same concept used in
# calculate() below (agent_fee, freight, etc.).
# ---------------------------------------------------------------------

def calculate_agent_fee(unit_price: float) -> float:
    """Agent Fee, based on unit price."""
    if unit_price < 1_000_000:
        return 140_000.0
    elif unit_price < 2_000_000:
        return 150_000.0
    elif unit_price < 3_000_000:
        return 160_000.0
    elif unit_price < 4_000_000:
        return 280_000.0
    else:
        return unit_price * 0.10


def calculate_under_value_amount(unit_price: float) -> float:
    """Under Value Amount, based on unit price. No bracket is defined for
    unit price >= 4,000,000 - stop."""
    if unit_price < 1_000_000:
        return 120_000.0
    elif unit_price < 2_000_000:
        return 130_000.0
    elif unit_price < 3_000_000:
        return 140_000.0
    elif unit_price < 4_000_000:
        return 240_000.0
    else:
        raise CalculationStopped(
            "Under Value Amount: unit price is 4,000,000 JPY or higher, "
            "which has no defined bracket. Calculation stopped - needs "
            "manual review."
        )


def calculate_freight(vehicle_m3_value: float) -> float:
    """Freight, based on Vehicle M3 Value. Table only covers 0 - 16.00 M3."""
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
        if low <= vehicle_m3_value <= high:
            return fee
    raise CalculationStopped(
        f"Freight: Vehicle M3 Value {vehicle_m3_value} is not in the "
        f"supported range (0 - 16.00 M3). Calculation stopped."
    )


def calculate_xid(fuel_type: str, engine_cc: int) -> float:
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

def calculate(vehicle_model: str, price: float) -> dict:
    """Runs the full formula for `vehicle_model` (must exactly match a key
    in vehicleDetails.json) at unit price `price` (JPY). Returns a dict of
    every intermediate value plus 'total_vehicle_cost'. Raises
    CalculationStopped part-way through if a required reference value is
    missing/out of range - the caller (process_vehicle_folder) catches
    this and writes whatever is in `partial` so far, so keep this dict
    updated at every step."""

    partial = {}

    inputs = _load_json(INPUT_VALUES_FILE)
    details_config = _load_json(VEHICLE_DETAILS_CONFIG_FILE)

    if vehicle_model not in details_config:
        raise _stopped_with(
            partial,
            f"Vehicle model '{vehicle_model}' not found in "
            f"vehicleDetails.json. Check the exact dropdown spelling, or "
            f"add this model to the config. Calculation stopped."
        )
    cfg = details_config[vehicle_model]

    unit_price = float(price or 0)
    auction_house_fee = 0.0  # defaulted to 0 until a per-auction-house
                              # json is added later.
    partial["unit_price"] = unit_price
    partial["auction_house_fee"] = auction_house_fee

    agent_fee = calculate_agent_fee(unit_price)
    partial["agent_fee"] = agent_fee

    inspection = 30_000.0
    partial["inspection"] = inspection

    if "m3Value" not in cfg:
        raise _stopped_with(
            partial,
            f"'m3Value' missing for '{vehicle_model}' in vehicleDetails.json. "
            f"Calculation stopped."
        )
    vehicle_m3_value = float(cfg["m3Value"])
    partial["vehicle_m3_value"] = vehicle_m3_value

    insurance = (unit_price + auction_house_fee + agent_fee) * vehicle_m3_value * 0.001 * 1.05
    partial["insurance"] = insurance

    try:
        freight = calculate_freight(vehicle_m3_value)
    except CalculationStopped as e:
        raise _stopped_with(partial, str(e))
    partial["freight"] = freight

    full_lc_cif = unit_price + auction_house_fee + agent_fee + inspection + insurance + freight
    partial["full_lc_cif"] = full_lc_cif

    try:
        under_value_amount = calculate_under_value_amount(unit_price)
    except CalculationStopped as e:
        raise _stopped_with(partial, str(e))
    partial["under_value_amount"] = under_value_amount

    preforma_invoice_cif = full_lc_cif - under_value_amount
    partial["preforma_invoice_cif"] = preforma_invoice_cif

    if "webValue" not in cfg:
        raise _stopped_with(
            partial,
            f"'webValue' missing for '{vehicle_model}' in vehicleDetails.json. "
            f"Calculation stopped."
        )
    web_value = float(cfg["webValue"])
    web_value_adjusted = web_value * (100 / 110) * (85 / 100)
    partial["web_value"] = web_value
    partial["web_value_adjusted"] = web_value_adjusted

    tax_base = max(web_value_adjusted, preforma_invoice_cif)
    partial["tax_base"] = tax_base

    jpy_rate = float(inputs["jpyRate"])
    partial["jpy_rate"] = jpy_rate

    cif = jpy_rate * tax_base
    cid = cif * 0.45
    partial["CIF"] = cif
    partial["CID"] = cid

    fuel_type = cfg.get("fuelType", "")
    try:
        engine_cc = int(cfg.get("engineCc", 0))
    except (TypeError, ValueError):
        raise _stopped_with(
            partial,
            f"'engineCc' for '{vehicle_model}' in vehicleDetails.json is not "
            f"a valid number. Calculation stopped."
        )

    try:
        xid = calculate_xid(fuel_type, engine_cc)
    except CalculationStopped as e:
        raise _stopped_with(partial, str(e))
    partial["XID"] = xid

    sscl_tax = (cif * 1.10 + cid + xid) * 0.025
    vat = (cif * 1.10 + cid + xid) * 0.18
    vel = float(inputs["velTax"])
    partial["SSCL_TAX"] = sscl_tax
    partial["VAT"] = vat
    partial["VEL"] = vel

    total_tax = cid + xid + sscl_tax + vat + vel
    partial["total_tax"] = total_tax

    under_value_and_charges = under_value_amount * jpy_rate * 1.10
    vehicle_value = preforma_invoice_cif * jpy_rate + under_value_and_charges
    partial["under_value_and_charges"] = under_value_and_charges
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


def format_breakdown(vehicle_model: str, vehicle_code: str, result: dict,
                      stopped_message: str = None) -> str:
    lines = []
    lines.append(f"Price Breakdown - {vehicle_model}")
    lines.append(f"Vehicle Code: {vehicle_code}")
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
    '<vehicle_model> <year> - <price>' line) with the same text but the
    calculated price, leaving all other lines untouched."""
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
    """Reads <folder>/vehicle_info.json (must contain 'vehicle_model' and
    'price'), runs the calculation, writes <folder>/price_breakdown.txt,
    and updates <folder>/description.txt's price line. Returns True on a
    full successful calculation, False if it stopped early (details are
    in price_breakdown.txt either way)."""
    info_path = folder / "vehicle_info.json"
    if not info_path.exists():
        print(f"    ! price_calculator: vehicle_info.json not found in {folder}")
        return False

    with open(info_path, "r", encoding="utf-8") as f:
        vehicle_info = json.load(f)

    vehicle_model = vehicle_info.get("vehicle_model", "")
    vehicle_code = vehicle_info.get("vehicle_code", "")
    price = vehicle_info.get("price", 0)

    try:
        result = calculate(vehicle_model, price)
        breakdown_text = format_breakdown(vehicle_model, vehicle_code, result)
        (folder / "price_breakdown.txt").write_text(breakdown_text, encoding="utf-8")

        price_label = f"LKR {round_to_lakhs(result['total_vehicle_cost'])} Lakhs"
        update_description_price(folder / "description.txt", price_label)

        print(f"    Price calculated: {price_label} "
              f"(Total Vehicle Cost: {_money(result['total_vehicle_cost'])} LKR)")
        return True

    except CalculationStopped as e:
        partial = getattr(e, "partial", {})
        breakdown_text = format_breakdown(vehicle_model, vehicle_code, partial, stopped_message=str(e))
        (folder / "price_breakdown.txt").write_text(breakdown_text, encoding="utf-8")

        update_description_price(
            folder / "description.txt",
            "PRICE CALCULATION STOPPED - see price_breakdown.txt"
        )
        print(f"    ! price calculation stopped: {e}")
        return False