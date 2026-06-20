from __future__ import annotations

import argparse
import csv
import dataclasses
import datetime as dt
import hashlib
import json
import math
import os
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlencode, urlparse
from urllib.request import Request, urlopen


CITIES = {"surrey", "langley", "white rock", "abbotsford", "chilliwack"}


@dataclasses.dataclass
class Assumptions:
    mortgage_interest_rate: float = 0.0525
    amortization_years: int = 25
    down_payment_pct: float = 0.20
    monthly_buffer: float = 250.0
    vacancy_pct: float = 0.03
    strong_appreciation_threshold: float = 0.08
    near_positive_cash_flow_threshold: float = -250.0


@dataclasses.dataclass
class PropertyListing:
    address: str
    city: str
    mls_number: str
    list_price: float
    property_type: str = ""
    bedrooms: float | None = None
    bathrooms: float | None = None
    square_feet: float | None = None
    lot_size: str = ""
    property_tax: float | None = None
    strata_fee: float | None = None
    year_built: int | None = None
    listing_date: str = ""
    source_url: str = ""

    @property
    def listing_id(self) -> str:
        raw = self.mls_number or f"{self.address}|{self.city}|{self.list_price}"
        return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12].upper()


@dataclasses.dataclass
class ZealtyRecord:
    sold_price_history: str = ""
    sale_dates: str = ""
    price_1y_ago: float | None = None
    price_3y_ago: float | None = None
    price_5y_ago: float | None = None
    previous_listing_prices: str = ""
    price_change_history: str = ""
    days_on_market: int | None = None
    comparable_sales: list[dict[str, Any]] = dataclasses.field(default_factory=list)
    source_url: str = ""
    source_notes: str = ""


@dataclasses.dataclass
class RentalEstimate:
    estimated_rent: float | None = None
    comps: list[dict[str, Any]] = dataclasses.field(default_factory=list)
    source_notes: str = ""


@dataclasses.dataclass
class LocationSignals:
    transit_minutes: float | None = None
    school_score: float | None = None
    development_score: float | None = None
    notes: str = ""


REALTOR_MAP_SEARCH_URL = (
    "https://api2.realtor.ca/Listing.svc/PropertySearch_Post"
)


class RealtorSavedSearchProvider:
    """Provider boundary for Realtor.ca saved searches.

    In production, wire this to either a browser session export or a compliant
    listing feed. The CSV path is the deterministic local mode used for testing.
    """

    def __init__(self, csv_path: Path | None = None, search_url: str | None = None) -> None:
        self.csv_path = csv_path
        self.search_url = search_url

    def fetch(self) -> list[PropertyListing]:
        if self.search_url:
            return self.fetch_from_realtor_url(self.search_url)
        if not self.csv_path:
            raise RuntimeError("Set --realtor-url to a Realtor.ca map search or --realtor-csv to a saved-search export.")
        rows = list(csv.DictReader(self.csv_path.open("r", encoding="utf-8-sig")))
        listings: list[PropertyListing] = []
        for row in rows:
            city = clean(row.get("City", ""))
            if city.lower() not in CITIES:
                continue
            listings.append(
                PropertyListing(
                    address=clean(row.get("Address", "")),
                    city=city,
                    mls_number=clean(row.get("MLS Number", row.get("MLS", ""))),
                    list_price=money(row.get("List Price", row.get("Price", ""))) or 0,
                    property_type=clean(row.get("Property Type", "")),
                    bedrooms=number(row.get("Bedrooms", "")),
                    bathrooms=number(row.get("Bathrooms", "")),
                    square_feet=number(row.get("Square Feet", row.get("Sq Ft", ""))),
                    lot_size=clean(row.get("Lot Size", "")),
                    property_tax=money(row.get("Property Tax", "")),
                    strata_fee=money(row.get("Strata Fee", "")),
                    year_built=integer(row.get("Year Built", "")),
                    listing_date=clean(row.get("Listing Date", "")),
                    source_url=clean(row.get("Source URL", row.get("URL", ""))),
                )
            )
        return listings

    def fetch_from_realtor_url(self, search_url: str) -> list[PropertyListing]:
        params = realtor_map_params(search_url)
        listings: list[PropertyListing] = []
        page = 1
        while True:
            payload = realtor_payload(params, page)
            response = realtor_post(payload, search_url)
            page_results = response.get("Results") or []
            for result in page_results:
                listing = listing_from_realtor_result(result)
                if listing and listing.city.lower() in CITIES:
                    listings.append(listing)
            paging = response.get("Paging") or {}
            total_pages = int(paging.get("TotalPages") or 1)
            if page >= total_pages:
                break
            page += 1
        return listings


def realtor_map_params(search_url: str) -> dict[str, str]:
    parsed = urlparse(search_url)
    if "/real-estate/" in parsed.path.lower():
        raise ValueError(
            "This is a single Realtor.ca property page. Paste a Realtor.ca map search URL instead: "
            "open Realtor.ca map view, set your area/filters, then copy the URL that starts with "
            "https://www.realtor.ca/map# and contains LatitudeMax, LatitudeMin, LongitudeMax, and LongitudeMin."
        )
    fragment = parsed.fragment or ""
    if fragment.startswith("/"):
        fragment = fragment[1:]
    if "?" in fragment:
        fragment = fragment.split("?", 1)[1]
    values = {k: unquote(v[-1]) for k, v in parse_qs(fragment).items() if v}
    if not values:
        raise ValueError(
            "The Realtor.ca URL did not contain map search parameters. Paste the map search URL, not an individual "
            "property page. A valid URL starts with https://www.realtor.ca/map# and includes LatitudeMax, "
            "LatitudeMin, LongitudeMax, and LongitudeMin."
        )
    required = ["LatitudeMax", "LongitudeMax", "LatitudeMin", "LongitudeMin"]
    missing = [key for key in required if key not in values]
    if missing:
        raise ValueError(f"The Realtor.ca URL is missing required map bounds: {', '.join(missing)}")
    return values


def realtor_payload(params: dict[str, str], page: int) -> dict[str, str]:
    payload = {
        "ZoomLevel": params.get("ZoomLevel", "12"),
        "LatitudeMax": params["LatitudeMax"],
        "LongitudeMax": params["LongitudeMax"],
        "LatitudeMin": params["LatitudeMin"],
        "LongitudeMin": params["LongitudeMin"],
        "Sort": params.get("Sort", "6-D"),
        "PropertyTypeGroupID": params.get("PropertyTypeGroupID", "1"),
        "PropertySearchTypeId": params.get("PropertySearchTypeId", "1"),
        "TransactionTypeId": params.get("TransactionTypeId", "2"),
        "OwnershipTypeGroupId": params.get("OwnershipTypeGroupId", "1"),
        "Currency": params.get("Currency", "CAD"),
        "RecordsPerPage": params.get("RecordsPerPage", "200"),
        "ApplicationId": "1",
        "CultureId": "1",
        "Version": "7.0",
        "CurrentPage": str(page),
    }
    if "Center" in params:
        center = params["Center"].split(",")
        if len(center) == 2:
            payload["Latitude"] = center[0]
            payload["Longitude"] = center[1]
    return payload


def realtor_post(payload: dict[str, str], referer: str) -> dict[str, Any]:
    body = urlencode(payload).encode("utf-8")
    request = Request(
        REALTOR_MAP_SEARCH_URL,
        data=body,
        headers={
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Origin": "https://www.realtor.ca",
            "Referer": referer,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/126.0 Safari/537.36",
            "Accept": "application/json, text/javascript, */*; q=0.01",
        },
        method="POST",
    )
    try:
        with urlopen(request, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        raise RuntimeError(
            "Realtor.ca blocked or failed the map-search request. "
            "Use --realtor-csv for now, or run from an authenticated/browser-backed session."
        ) from exc


def listing_from_realtor_result(result: dict[str, Any]) -> PropertyListing | None:
    property_info = result.get("Property") or {}
    building = result.get("Building") or property_info.get("Building") or {}
    land = result.get("Land") or property_info.get("Land") or {}
    address = property_info.get("Address") or {}
    price_text = clean(property_info.get("PriceUnformattedValue") or property_info.get("Price"))
    mls_number = clean(result.get("MlsNumber") or result.get("MLSNumber"))
    address_text = clean(address.get("AddressText"))
    if "|" in address_text:
        street, city_part = [part.strip() for part in address_text.split("|", 1)]
        city = city_part.split(",")[0].strip()
    else:
        street = address_text
        city = clean(address.get("City"))
    if not street or not mls_number:
        return None
    beds = building.get("Bedrooms")
    baths = building.get("BathroomTotal")
    sqft = building.get("SizeInterior")
    return PropertyListing(
        address=street,
        city=city,
        mls_number=mls_number,
        list_price=money(price_text) or 0,
        property_type=clean(building.get("Type") or property_info.get("Type")),
        bedrooms=number(beds),
        bathrooms=number(baths),
        square_feet=parse_sqft(sqft),
        lot_size=clean(land.get("SizeTotal") or land.get("SizeFrontage")),
        property_tax=money(property_info.get("PropertyTax")),
        strata_fee=money(property_info.get("MaintenanceFee")),
        year_built=integer(building.get("ConstructedDate")),
        listing_date=realtor_date(result.get("InsertedDateUTC") or result.get("PhotoChangeDateUTC") or result.get("TimeOnRealtor")),
        source_url="https://www.realtor.ca" + clean(result.get("RelativeDetailsURL")),
    )


def realtor_date(value: Any) -> str:
    text = clean(value)
    if not text:
        return ""
    if text.isdigit() and len(text) >= 17:
        ticks = int(text)
        seconds = ticks / 10_000_000 - 62135596800
        try:
            return dt.datetime.fromtimestamp(seconds, tz=dt.timezone.utc).date().isoformat()
        except (OverflowError, OSError, ValueError):
            return text
    for fmt in ("%Y-%m-%d %I:%M:%S %p", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return text


def parse_sqft(value: Any) -> float | None:
    text = clean(value).lower().replace(",", "")
    if not text:
        return None
    for token in text.replace("-", " ").split():
        parsed = number(token)
        if parsed:
            return parsed
    return None


class ZealtyProvider:
    def __init__(self, json_path: Path | None = None) -> None:
        self.records = load_lookup(json_path)

    def fetch(self, listing: PropertyListing) -> ZealtyRecord:
        record = lookup_record(self.records, listing)
        if record:
            comps = record.get("comparable_sales") or record.get("Comparable Sales") or []
            return ZealtyRecord(
                sold_price_history=clean(record.get("sold_price_history", record.get("Sold Price History", ""))),
                sale_dates=clean(record.get("sale_dates", record.get("Sale Dates", ""))),
                price_1y_ago=money(record.get("price_1y_ago", record.get("Price 1Y Ago", ""))),
                price_3y_ago=money(record.get("price_3y_ago", record.get("Price 3Y Ago", ""))),
                price_5y_ago=money(record.get("price_5y_ago", record.get("Price 5Y Ago", ""))),
                previous_listing_prices=clean(record.get("previous_listing_prices", record.get("Previous Listing Prices", ""))),
                price_change_history=clean(record.get("price_change_history", record.get("Price Change History", ""))),
                days_on_market=integer(record.get("days_on_market", record.get("Days on Market", ""))),
                comparable_sales=comps if isinstance(comps, list) else [],
                source_url=clean(record.get("source_url", record.get("Zealty URL", ""))),
                source_notes=clean(record.get("source_notes", record.get("Source Notes", "Loaded from Zealty JSON"))),
            )
        # TODO: Connect with an authenticated Zealty search workflow.
        # Keep this provider isolated because Zealty access often depends on a logged-in browser session.
        return ZealtyRecord(source_notes="Pending Zealty login/search connection")


class RentalProvider:
    def __init__(self, json_path: Path | None = None) -> None:
        self.records = load_lookup(json_path)

    def fetch(self, listing: PropertyListing) -> RentalEstimate:
        record = lookup_record(self.records, listing)
        if record:
            comps = record.get("comps") or record.get("rental_comps") or record.get("Rental Comps") or []
            return RentalEstimate(
                estimated_rent=money(record.get("estimated_rent", record.get("Estimated Rent", ""))),
                comps=comps if isinstance(comps, list) else [],
                source_notes=clean(record.get("source_notes", record.get("Source Notes", "Loaded from rental JSON"))),
            )
        # TODO: Connect Rentals.ca, Craigslist, Facebook Marketplace export, or a paid rental feed.
        return RentalEstimate(source_notes="Pending rental listing source connection")


class SignalProvider:
    def __init__(self, json_path: Path | None = None) -> None:
        self.records = load_lookup(json_path)

    def fetch(self, listing: PropertyListing) -> LocationSignals:
        record = lookup_record(self.records, listing)
        if record:
            return LocationSignals(
                transit_minutes=number(record.get("transit_minutes", record.get("Transit Minutes", ""))),
                school_score=number(record.get("school_score", record.get("School Score", ""))),
                development_score=number(record.get("development_score", record.get("Development Score", ""))),
                notes=clean(record.get("notes", record.get("Notes", "Loaded from signal JSON"))),
            )
        # TODO: Connect Google Maps/Transit and school score sources.
        return LocationSignals(notes="Pending transit/school/development data connection")


def load_lookup(json_path: Path | None) -> dict[str, dict[str, Any]]:
    if not json_path:
        return {}
    payload = json.loads(json_path.read_text(encoding="utf-8"))
    if isinstance(payload, dict):
        records = payload.get("records", payload)
        if isinstance(records, dict):
            return {clean(k).lower(): v for k, v in records.items() if isinstance(v, dict)}
        if isinstance(records, list):
            return index_records(records)
    if isinstance(payload, list):
        return index_records(payload)
    raise ValueError(f"Unsupported lookup JSON shape: {json_path}")


def index_records(records: list[Any]) -> dict[str, dict[str, Any]]:
    indexed: dict[str, dict[str, Any]] = {}
    for record in records:
        if not isinstance(record, dict):
            continue
        for key_name in ("listing_id", "Listing ID", "mls_number", "MLS Number", "MLS", "address", "Address"):
            key = clean(record.get(key_name)).lower()
            if key:
                indexed[key] = record
    return indexed


def lookup_record(records: dict[str, dict[str, Any]], listing: PropertyListing) -> dict[str, Any] | None:
    keys = [listing.listing_id, listing.mls_number, listing.address]
    for key in keys:
        found = records.get(clean(key).lower())
        if found:
            return found
    return None


def clean(value: Any) -> str:
    return str(value or "").strip()


def money(value: Any) -> float | None:
    text = clean(value).replace("$", "").replace(",", "")
    return number(text)


def number(value: Any) -> float | None:
    text = clean(value)
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def integer(value: Any) -> int | None:
    parsed = number(value)
    return int(parsed) if parsed is not None else None


def mortgage_payment(price: float, assumptions: Assumptions) -> float:
    principal = price * (1 - assumptions.down_payment_pct)
    monthly_rate = assumptions.mortgage_interest_rate / 12
    months = assumptions.amortization_years * 12
    if monthly_rate == 0:
        return principal / months
    return principal * monthly_rate * (1 + monthly_rate) ** months / ((1 + monthly_rate) ** months - 1)


def pct_change(current: float, prior: float | None) -> float | None:
    if not prior:
        return None
    return (current - prior) / prior


def clamp(value: float, low: float = 0, high: float = 100) -> float:
    return max(low, min(high, value))


def score_property(
    listing: PropertyListing,
    zealty: ZealtyRecord,
    rental: RentalEstimate,
    signals: LocationSignals,
    assumptions: Assumptions,
) -> dict[str, Any]:
    price_per_sqft = listing.list_price / listing.square_feet if listing.square_feet else None
    change_1y = pct_change(listing.list_price, zealty.price_1y_ago)
    change_3y = pct_change(listing.list_price, zealty.price_3y_ago)
    change_5y = pct_change(listing.list_price, zealty.price_5y_ago)
    comp_prices = [money(c.get("sale_price")) for c in zealty.comparable_sales]
    comp_prices = [p for p in comp_prices if p]
    avg_comp = sum(comp_prices) / len(comp_prices) if comp_prices else None
    comp_discount = (listing.list_price - avg_comp) / avg_comp if avg_comp else None

    monthly_mortgage = mortgage_payment(listing.list_price, assumptions)
    monthly_tax = (listing.property_tax or 0) / 12
    monthly_strata = listing.strata_fee or 0
    estimated_rent = rental.estimated_rent
    total_monthly_cost = monthly_mortgage + monthly_tax + monthly_strata + assumptions.monthly_buffer
    if estimated_rent is not None:
        total_monthly_cost += estimated_rent * assumptions.vacancy_pct
    cash_flow = estimated_rent - total_monthly_cost if estimated_rent is not None else None

    appreciation_score = clamp(((change_1y or 0) / assumptions.strong_appreciation_threshold) * 100)
    cash_flow_score = clamp(50 + ((cash_flow or -1000) / 1000) * 50)
    comp_score = clamp((-(comp_discount or 0) / 0.10) * 100)
    transit_score = clamp(((12 - (signals.transit_minutes or 12)) / 12) * 100)
    school_score = clamp((signals.school_score or 0) * 10)
    development_score = clamp((signals.development_score or 0) * 10)

    investment_score = round(
        appreciation_score * 0.30
        + cash_flow_score * 0.25
        + comp_score * 0.20
        + transit_score * 0.10
        + school_score * 0.05
        + development_score * 0.10
    )

    flags = []
    if comp_discount is not None and comp_discount <= -0.03:
        flags.append("Below comps")
    if change_1y is not None and change_1y >= assumptions.strong_appreciation_threshold:
        flags.append("Strong appreciation")
    if cash_flow is not None and cash_flow >= assumptions.near_positive_cash_flow_threshold:
        flags.append("Near-positive cash flow")
    if investment_score >= 70 and cash_flow is not None and cash_flow >= assumptions.near_positive_cash_flow_threshold:
        flags.append("Mortgage-investment fit")

    notes = build_notes(listing, change_1y, change_5y, comp_discount, cash_flow, flags)

    return {
        "Listing ID": listing.listing_id,
        "Source URL": listing.source_url,
        "Address": listing.address,
        "City": listing.city,
        "MLS Number": listing.mls_number,
        "List Price": listing.list_price,
        "Property Type": listing.property_type,
        "Bedrooms": listing.bedrooms,
        "Bathrooms": listing.bathrooms,
        "Square Feet": listing.square_feet,
        "Lot Size": listing.lot_size,
        "Property Tax": listing.property_tax,
        "Strata Fee": listing.strata_fee,
        "Year Built": listing.year_built,
        "Listing Date": listing.listing_date,
        "Estimated Rent": estimated_rent,
        "Price / Sq Ft": price_per_sqft,
        "1-Year Change %": change_1y,
        "3-Year Change %": change_3y,
        "5-Year Change %": change_5y,
        "Avg Comparable Sale": avg_comp,
        "Comp Discount %": comp_discount,
        "Transit Minutes": signals.transit_minutes,
        "School Score": signals.school_score,
        "Development Score": signals.development_score,
        "Monthly Mortgage": monthly_mortgage,
        "Monthly Property Tax": monthly_tax,
        "Monthly Strata": monthly_strata,
        "Total Monthly Cost": total_monthly_cost,
        "Estimated Cash Flow": cash_flow,
        "Investment Score": investment_score,
        "Flags": ", ".join(flags),
        "Notes": notes,
        "Last Seen": dt.datetime.now().isoformat(timespec="seconds"),
        "_zealty": dataclasses.asdict(zealty),
        "_rental": dataclasses.asdict(rental),
        "_signals": dataclasses.asdict(signals),
    }


PUBLIC_HEADERS = [
    "Listing ID",
    "Source URL",
    "Address",
    "City",
    "MLS Number",
    "List Price",
    "Property Type",
    "Bedrooms",
    "Bathrooms",
    "Square Feet",
    "Lot Size",
    "Property Tax",
    "Strata Fee",
    "Year Built",
    "Listing Date",
    "Estimated Rent",
    "Price / Sq Ft",
    "1-Year Change %",
    "3-Year Change %",
    "5-Year Change %",
    "Avg Comparable Sale",
    "Comp Discount %",
    "Transit Minutes",
    "School Score",
    "Development Score",
    "Monthly Mortgage",
    "Monthly Property Tax",
    "Monthly Strata",
    "Total Monthly Cost",
    "Estimated Cash Flow",
    "Investment Score",
    "Flags",
    "Notes",
    "Last Seen",
]


def build_notes(
    listing: PropertyListing,
    change_1y: float | None,
    change_5y: float | None,
    comp_discount: float | None,
    cash_flow: float | None,
    flags: list[str],
) -> str:
    parts = []
    if flags:
        parts.append("; ".join(flags))
    if change_1y is not None:
        parts.append(f"1Y change {change_1y:.1%}")
    if change_5y is not None:
        parts.append(f"5Y change {change_5y:.1%}")
    if comp_discount is not None:
        parts.append(f"{abs(comp_discount):.1%} {'below' if comp_discount < 0 else 'above'} comparable average")
    if cash_flow is not None:
        parts.append(f"estimated cash flow ${cash_flow:,.0f}/mo")
    if listing.square_feet:
        parts.append(f"${listing.list_price / listing.square_feet:,.0f}/sf")
    return ". ".join(parts)


def load_state(state_file: Path | None) -> dict[str, dict[str, Any]]:
    if not state_file or not state_file.exists():
        return {}
    rows = json.loads(state_file.read_text(encoding="utf-8"))
    if not isinstance(rows, list):
        return {}
    return {clean(row.get("Listing ID")).lower(): row for row in rows if isinstance(row, dict) and row.get("Listing ID")}


def merge_with_state(rows: list[dict[str, Any]], previous: dict[str, dict[str, Any]]) -> tuple[list[dict[str, Any]], dict[str, int]]:
    now_ids = {clean(row.get("Listing ID")).lower() for row in rows}
    new_count = 0
    updated_count = 0
    merged = []
    for row in rows:
        key = clean(row.get("Listing ID")).lower()
        old = previous.get(key)
        if not old:
            new_count += 1
            merged.append(row)
            continue
        updated_count += int(any(row.get(k) != old.get(k) for k in ("List Price", "Estimated Rent", "Investment Score", "Flags")))
        merged.append(row)
    stale = [row for key, row in previous.items() if key not in now_ids]
    for row in stale:
        archived = dict(row)
        archived["Flags"] = append_flag(clean(archived.get("Flags")), "Not in latest search")
        merged.append(archived)
    return merged, {"new": new_count, "updated": updated_count, "stale": len(stale)}


def append_flag(flags: str, flag: str) -> str:
    parts = [part.strip() for part in flags.split(",") if part.strip()]
    if flag not in parts:
        parts.append(flag)
    return ", ".join(parts)


def write_outputs(rows: list[dict[str, Any]], out_dir: Path, run_stats: dict[str, int]) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    db_path = out_dir / "property_database_rows.csv"
    json_path = out_dir / "property_database_rows.json"
    top_path = out_dir / "top_10_investment_opportunities.csv"
    zealty_path = out_dir / "zealty_history_rows.csv"
    rentals_path = out_dir / "rental_comps_rows.csv"
    comps_path = out_dir / "comparable_sales_rows.csv"
    log_path = out_dir / "update_log.csv"

    with db_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=PUBLIC_HEADERS)
        writer.writeheader()
        writer.writerows([{k: row.get(k) for k in PUBLIC_HEADERS} for row in rows])

    json_path.write_text(json.dumps(rows, indent=2, default=str), encoding="utf-8")

    top_fields = [
        "Address",
        "List Price",
        "1-Year Change %",
        "5-Year Change %",
        "Estimated Rent",
        "Investment Score",
        "Notes",
    ]
    ranked = sorted(rows, key=lambda r: r.get("Investment Score") or -math.inf, reverse=True)[:10]
    with top_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=top_fields)
        writer.writeheader()
        writer.writerows([{k: row.get(k) for k in top_fields} for row in ranked])

    write_zealty_rows(rows, zealty_path)
    write_rental_rows(rows, rentals_path)
    write_comp_rows(rows, comps_path)
    with log_path.open("w", newline="", encoding="utf-8") as f:
        fields = ["Run ID", "Started At", "Finished At", "Listings Found", "New Listings", "Updated Listings", "Stale Listings", "Notes"]
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerow(
            {
                "Run ID": run_stats["run_id"],
                "Started At": run_stats["started_at"],
                "Finished At": dt.datetime.now().isoformat(timespec="seconds"),
                "Listings Found": run_stats["found"],
                "New Listings": run_stats["new"],
                "Updated Listings": run_stats["updated"],
                "Stale Listings": run_stats["stale"],
                "Notes": "Generated by local research agent",
            }
        )


def write_zealty_rows(rows: list[dict[str, Any]], path: Path) -> None:
    fields = ["Listing ID", "Address", "MLS Number", "Sold Price History", "Sale Dates", "Price 1Y Ago", "Price 3Y Ago", "Price 5Y Ago", "Previous Listing Prices", "Price Change History", "Days on Market", "Zealty URL", "Collected At", "Source Notes"]
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            z = row.get("_zealty", {})
            writer.writerow(
                {
                    "Listing ID": row.get("Listing ID"),
                    "Address": row.get("Address"),
                    "MLS Number": row.get("MLS Number"),
                    "Sold Price History": z.get("sold_price_history"),
                    "Sale Dates": z.get("sale_dates"),
                    "Price 1Y Ago": z.get("price_1y_ago"),
                    "Price 3Y Ago": z.get("price_3y_ago"),
                    "Price 5Y Ago": z.get("price_5y_ago"),
                    "Previous Listing Prices": z.get("previous_listing_prices"),
                    "Price Change History": z.get("price_change_history"),
                    "Days on Market": z.get("days_on_market"),
                    "Zealty URL": z.get("source_url"),
                    "Collected At": row.get("Last Seen"),
                    "Source Notes": z.get("source_notes"),
                }
            )


def write_rental_rows(rows: list[dict[str, Any]], path: Path) -> None:
    fields = ["Listing ID", "Property Address", "Rental Address", "Distance Km", "Beds", "Baths", "Sq Ft", "Monthly Rent", "Rent / Sq Ft", "Source URL", "Collected At", "Notes"]
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            rental = row.get("_rental", {})
            comps = rental.get("comps") or [{}]
            for comp in comps:
                rent = money(comp.get("monthly_rent", comp.get("Monthly Rent", "")))
                sqft = number(comp.get("square_feet", comp.get("Sq Ft", "")))
                writer.writerow(
                    {
                        "Listing ID": row.get("Listing ID"),
                        "Property Address": row.get("Address"),
                        "Rental Address": comp.get("address", comp.get("Rental Address", "")),
                        "Distance Km": comp.get("distance_km", comp.get("Distance Km", "")),
                        "Beds": comp.get("beds", comp.get("Beds", "")),
                        "Baths": comp.get("baths", comp.get("Baths", "")),
                        "Sq Ft": sqft,
                        "Monthly Rent": rent,
                        "Rent / Sq Ft": rent / sqft if rent and sqft else "",
                        "Source URL": comp.get("source_url", comp.get("Source URL", "")),
                        "Collected At": row.get("Last Seen"),
                        "Notes": comp.get("notes", rental.get("source_notes", "")),
                    }
                )


def write_comp_rows(rows: list[dict[str, Any]], path: Path) -> None:
    fields = ["Listing ID", "Subject Address", "Comparable Address", "Distance Km", "Sale Date", "Sale Price", "Beds", "Baths", "Sq Ft", "Price / Sq Ft", "Source URL", "Notes"]
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            z = row.get("_zealty", {})
            comp_rows = z.get("comparable_sales") or [{}]
            for comp in comp_rows:
                sale_price = money(comp.get("sale_price", comp.get("Sale Price", "")))
                sqft = number(comp.get("square_feet", comp.get("Sq Ft", "")))
                writer.writerow(
                    {
                        "Listing ID": row.get("Listing ID"),
                        "Subject Address": row.get("Address"),
                        "Comparable Address": comp.get("address", comp.get("Comparable Address", "")),
                        "Distance Km": comp.get("distance_km", comp.get("Distance Km", "")),
                        "Sale Date": comp.get("sale_date", comp.get("Sale Date", "")),
                        "Sale Price": sale_price,
                        "Beds": comp.get("beds", comp.get("Beds", "")),
                        "Baths": comp.get("baths", comp.get("Baths", "")),
                        "Sq Ft": sqft,
                        "Price / Sq Ft": sale_price / sqft if sale_price and sqft else "",
                        "Source URL": comp.get("source_url", comp.get("Source URL", "")),
                        "Notes": comp.get("notes", z.get("source_notes", "")),
                    }
                )


def zealty_sheet_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output = []
    for row in rows:
        z = row.get("_zealty", {})
        output.append(
            {
                "Listing ID": row.get("Listing ID"),
                "Address": row.get("Address"),
                "MLS Number": row.get("MLS Number"),
                "Sold Price History": z.get("sold_price_history"),
                "Sale Dates": z.get("sale_dates"),
                "Price 1Y Ago": z.get("price_1y_ago"),
                "Price 3Y Ago": z.get("price_3y_ago"),
                "Price 5Y Ago": z.get("price_5y_ago"),
                "Previous Listing Prices": z.get("previous_listing_prices"),
                "Price Change History": z.get("price_change_history"),
                "Days on Market": z.get("days_on_market"),
                "Zealty URL": z.get("source_url"),
                "Collected At": row.get("Last Seen"),
                "Source Notes": z.get("source_notes"),
            }
        )
    return output


def rental_sheet_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output = []
    for row in rows:
        rental = row.get("_rental", {})
        comps = rental.get("comps") or [{}]
        for comp in comps:
            rent = money(comp.get("monthly_rent", comp.get("Monthly Rent", "")))
            sqft = number(comp.get("square_feet", comp.get("Sq Ft", "")))
            output.append(
                {
                    "Listing ID": row.get("Listing ID"),
                    "Property Address": row.get("Address"),
                    "Rental Address": comp.get("address", comp.get("Rental Address", "")),
                    "Distance Km": comp.get("distance_km", comp.get("Distance Km", "")),
                    "Beds": comp.get("beds", comp.get("Beds", "")),
                    "Baths": comp.get("baths", comp.get("Baths", "")),
                    "Sq Ft": sqft,
                    "Monthly Rent": rent,
                    "Rent / Sq Ft": rent / sqft if rent and sqft else "",
                    "Source URL": comp.get("source_url", comp.get("Source URL", "")),
                    "Collected At": row.get("Last Seen"),
                    "Notes": comp.get("notes", rental.get("source_notes", "")),
                }
            )
    return output


def comparable_sheet_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output = []
    for row in rows:
        z = row.get("_zealty", {})
        comp_rows = z.get("comparable_sales") or [{}]
        for comp in comp_rows:
            sale_price = money(comp.get("sale_price", comp.get("Sale Price", "")))
            sqft = number(comp.get("square_feet", comp.get("Sq Ft", "")))
            output.append(
                {
                    "Listing ID": row.get("Listing ID"),
                    "Subject Address": row.get("Address"),
                    "Comparable Address": comp.get("address", comp.get("Comparable Address", "")),
                    "Distance Km": comp.get("distance_km", comp.get("Distance Km", "")),
                    "Sale Date": comp.get("sale_date", comp.get("Sale Date", "")),
                    "Sale Price": sale_price,
                    "Beds": comp.get("beds", comp.get("Beds", "")),
                    "Baths": comp.get("baths", comp.get("Baths", "")),
                    "Sq Ft": sqft,
                    "Price / Sq Ft": sale_price / sqft if sale_price and sqft else "",
                    "Source URL": comp.get("source_url", comp.get("Source URL", "")),
                    "Notes": comp.get("notes", z.get("source_notes", "")),
                }
            )
    return output


def sync_google_sheet(rows: list[dict[str, Any]]) -> None:
    spreadsheet_id = os.getenv("PROPERTY_DATABASE_SHEET_ID")
    service_account_json = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if not spreadsheet_id or not service_account_json:
        print("Google Sheets sync skipped: set PROPERTY_DATABASE_SHEET_ID and GOOGLE_APPLICATION_CREDENTIALS.")
        return
    try:
        import gspread  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Install gspread in the runtime used for scheduled sync.") from exc

    client = gspread.service_account(filename=service_account_json)
    spreadsheet = client.open_by_key(spreadsheet_id)
    update_worksheet(spreadsheet, "Property Database", PUBLIC_HEADERS, rows)
    top_fields = ["Address", "List Price", "1-Year Change %", "5-Year Change %", "Estimated Rent", "Investment Score", "Notes"]
    ranked = sorted(rows, key=lambda r: r.get("Investment Score") or -math.inf, reverse=True)[:10]
    update_worksheet(spreadsheet, "Top 10 Investment Opportunities", top_fields, ranked)
    zealty_fields = ["Listing ID", "Address", "MLS Number", "Sold Price History", "Sale Dates", "Price 1Y Ago", "Price 3Y Ago", "Price 5Y Ago", "Previous Listing Prices", "Price Change History", "Days on Market", "Zealty URL", "Collected At", "Source Notes"]
    rental_fields = ["Listing ID", "Property Address", "Rental Address", "Distance Km", "Beds", "Baths", "Sq Ft", "Monthly Rent", "Rent / Sq Ft", "Source URL", "Collected At", "Notes"]
    comp_fields = ["Listing ID", "Subject Address", "Comparable Address", "Distance Km", "Sale Date", "Sale Price", "Beds", "Baths", "Sq Ft", "Price / Sq Ft", "Source URL", "Notes"]
    update_worksheet(spreadsheet, "Zealty History", zealty_fields, zealty_sheet_rows(rows))
    update_worksheet(spreadsheet, "Rental Comps", rental_fields, rental_sheet_rows(rows))
    update_worksheet(spreadsheet, "Comparable Sales", comp_fields, comparable_sheet_rows(rows))


def update_worksheet(spreadsheet: Any, name: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
    try:
        sheet = spreadsheet.worksheet(name)
    except Exception:
        sheet = spreadsheet.add_worksheet(title=name, rows=max(100, len(rows) + 5), cols=len(headers) + 2)
    values = [headers] + [[serialize_cell(row.get(h, "")) for h in headers] for row in rows]
    sheet.clear()
    sheet.update(values)


def serialize_cell(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (int, float, str, bool)):
        return value
    return json.dumps(value, default=str)


def run(args: argparse.Namespace) -> None:
    started_at = dt.datetime.now().isoformat(timespec="seconds")
    assumptions = Assumptions()
    realtor = RealtorSavedSearchProvider(Path(args.realtor_csv) if args.realtor_csv else None, args.realtor_url)
    zealty = ZealtyProvider(Path(args.zealty_json) if args.zealty_json else None)
    rentals = RentalProvider(Path(args.rentals_json) if args.rentals_json else None)
    signals = SignalProvider(Path(args.signals_json) if args.signals_json else None)

    listings = realtor.fetch()
    rows = []
    for listing in listings:
        z = zealty.fetch(listing)
        r = rentals.fetch(listing)
        s = signals.fetch(listing)
        rows.append(score_property(listing, z, r, s, assumptions))

    out_dir = Path(args.output_dir)
    previous = load_state(Path(args.state_file) if args.state_file else out_dir / "property_database_rows.json")
    merged_rows, counts = merge_with_state(rows, previous)
    run_stats = {
        "run_id": dt.datetime.now().strftime("%Y%m%d%H%M%S"),
        "started_at": started_at,
        "found": len(rows),
        **counts,
    }
    write_outputs(merged_rows, out_dir, run_stats)
    if args.sync_google:
        sync_google_sheet(merged_rows)
    print(
        f"Processed {len(rows)} current listings "
        f"({counts['new']} new, {counts['updated']} updated, {counts['stale']} stale). "
        f"Outputs saved to {out_dir}."
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Real estate investment research agent")
    parser.add_argument("--realtor-url", help="Realtor.ca map search URL")
    parser.add_argument("--realtor-csv", help="CSV export from saved Realtor.ca search")
    parser.add_argument("--zealty-json", help="Optional Zealty enrichment JSON keyed by listing ID, MLS, or address")
    parser.add_argument("--rentals-json", help="Optional rental enrichment JSON keyed by listing ID, MLS, or address")
    parser.add_argument("--signals-json", help="Optional transit/school/development JSON keyed by listing ID, MLS, or address")
    parser.add_argument("--state-file", help="Previous property_database_rows.json for incremental updates")
    parser.add_argument("--output-dir", default="outputs/agent_runs")
    parser.add_argument("--sync-google", action="store_true", help="Sync rows to Google Sheet when credentials are configured")
    return parser.parse_args()


if __name__ == "__main__":
    run(parse_args())
