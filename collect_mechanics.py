"""
Three-Wheeler Mechanic Data Collection Pipeline
===============================================
Collects business name, phone number, address, and location for independent
three-wheeler (auto rickshaw / tempo) repair mechanics in target Indian
cities/towns, using the official Google Places API (Text Search + Place
Details).

SETUP (one-time)
-----------------
1. Get a Google Cloud API key with "Places API" enabled:
   https://console.cloud.google.com/  -> APIs & Services -> Enable "Places API"
   -> Credentials -> Create API Key
   (Google gives $200/month free credit, which covers a good number of calls.
    Track your usage on the Cloud Console to avoid surprise charges.)

2. Install dependencies:
   pip install requests --break-system-packages

3. Set your API key as an environment variable (safer than hardcoding it):
   export GOOGLE_PLACES_API_KEY="your_key_here"        (Mac/Linux)
   set GOOGLE_PLACES_API_KEY="your_key_here"            (Windows cmd)

4. Edit the CONFIG section below with your target locations and search terms.

5. Run:
   python collect_mechanics.py

Output: one CSV per city (e.g. mechanics_three_wheeler_bengaluru.csv),
written to the same folder.
"""

import os
import csv
import time
import requests

# ============================================================
# CONFIG - edit this section for your use case
# ============================================================

API_KEY = os.environ.get("GOOGLE_PLACES_API_KEY", "")

# Target locations, grouped by city: use city/town or locality names (Places
# API will geocode them via text search). Note: "tehsil"/"taluk" admin units
# don't map cleanly onto the big metros below (Kolkata's district has none at
# all; Bengaluru's and Navi Mumbai's are generic/non-place names that
# geocode poorly), so those cities use well-known localities spread across
# the city instead, for better Places API hit rate and geographic coverage.
# Each city gets its own output CSV - see OUTPUT_FILE_TEMPLATE below.
TARGET_LOCATIONS_BY_CITY = {
    # Dehradun district, Uttarakhand - city + every tehsil
    # (per the official district site, dehradun.nic.in/tehsils/)
    # Disabled for now - uncomment to include Dehradun in the run.
    # "Dehradun": [
    #     "Dehradun Sadar, Dehradun, Uttarakhand",
    #     "Doiwala, Dehradun, Uttarakhand",
    #     "Rishikesh, Dehradun, Uttarakhand",
    #     "Vikasnagar, Dehradun, Uttarakhand",
    #     "Chakrata, Dehradun, Uttarakhand",
    #     "Kalsi, Dehradun, Uttarakhand",
    #     "Tyuni, Dehradun, Uttarakhand",
    #     "Mussoorie, Dehradun, Uttarakhand",
    # ],

    # Bengaluru - well-known localities across the city
    "Bengaluru": [
        "Indiranagar, Bengaluru",
        "Koramangala, Bengaluru",
        "Whitefield, Bengaluru",
        "Jayanagar, Bengaluru",
        "Malleshwaram, Bengaluru",
        "Electronic City, Bengaluru",
        "Marathahalli, Bengaluru",
        "Hebbal, Bengaluru",
        "Banashankari, Bengaluru",
        "HSR Layout, Bengaluru",
        "Yelahanka, Bengaluru",
        "RT Nagar, Bengaluru",
    ],

    # Chennai - well-known localities across the city
    "Chennai": [
        "T Nagar, Chennai",
        "Anna Nagar, Chennai",
        "Adyar, Chennai",
        "Velachery, Chennai",
        "Tambaram, Chennai",
        "Porur, Chennai",
        "Guindy, Chennai",
        "Mylapore, Chennai",
        "Perambur, Chennai",
        "Ambattur, Chennai",
        "Sholinganallur, Chennai",
        "Egmore, Chennai",
    ],

    # Kolkata - well-known localities across the city
    "Kolkata": [
        "Salt Lake, Kolkata",
        "Park Street, Kolkata",
        "Gariahat, Kolkata",
        "Behala, Kolkata",
        "Dum Dum, Kolkata",
        "Garia, Kolkata",
        "Tollygunge, Kolkata",
        "Rajarhat New Town, Kolkata",
        "Jadavpur, Kolkata",
        "Shyambazar, Kolkata",
        "Ballygunge, Kolkata",
    ],

    # Navi Mumbai - well-known planned "nodes" across the city
    "Navi Mumbai": [
        "Vashi, Navi Mumbai",
        "Nerul, Navi Mumbai",
        "Belapur, Navi Mumbai",
        "Kharghar, Navi Mumbai",
        "Kopar Khairane, Navi Mumbai",
        "Airoli, Navi Mumbai",
        "Panvel, Navi Mumbai",
        "Kalamboli, Navi Mumbai",
        "Sanpada, Navi Mumbai",
        "Turbhe, Navi Mumbai",
    ],

    # Delhi - well-known localities across the city
    "Delhi": [
        "Karol Bagh, Delhi",
        "Connaught Place, Delhi",
        "Lajpat Nagar, Delhi",
        "Dwarka, Delhi",
        "Rohini, Delhi",
        "Rajouri Garden, Delhi",
        "Saket, Delhi",
        "Pitampura, Delhi",
        "Janakpuri, Delhi",
        "Vasant Kunj, Delhi",
        "Laxmi Nagar, Delhi",
        "Mayur Vihar, Delhi",
    ],
}

# Search terms to try per location. Multiple phrasings improve coverage
# since business listings are inconsistently worded. Kept auto-rickshaw
# specific (dropped the bare "three wheeler mechanic" phrasing - Places'
# fuzzy text matching was pulling in plain two-wheeler garages on that one).
SEARCH_TERMS = [
    "auto rickshaw repair",
    "auto rickshaw garage",
    "auto rickshaw mechanic",
    "e-rickshaw repair",
    "CNG auto repair",
    "three wheeler repair garage",
]

# Delay between API calls (seconds) - be a good API citizen.
REQUEST_DELAY = 0.2

# One CSV per city, written to the same folder. {city} is replaced with a
# lowercase, underscore-separated slug of the city name.
OUTPUT_FILE_TEMPLATE = "mechanics_three_wheeler_{city}.csv"

# Keywords that suggest a listing is an AUTHORIZED dealer/service center
# rather than an independent mechanic - used for a simple heuristic flag.
DEALER_KEYWORDS = [
    "authorised", "authorized", "showroom", "pvt ltd", "private limited",
    "bajaj", "piaggio", "tvs", "mahindra", "atul", "force motors", "dealer",
]

# Places' text search returns a lot of plain two-wheeler garages even for
# three-wheeler-specific queries, so listings are filtered post-fetch: a
# result is dropped if its name reads as two-wheeler-only (has a bike/
# scooter/motorcycle word) and doesn't also mention anything three-wheeler
# related. Ambiguous names (e.g. "Sharma Motor Works", no vehicle-type word
# at all) are kept, since they may well service autos too.
TWO_WHEELER_EXCLUDE_KEYWORDS = [
    "bike", "motorbike", "motorcycle", "moto ", "scooter", "two wheeler",
    "2 wheeler", "moped", "cycle repair", "cycle works", "cycle garage",
]
THREE_WHEELER_KEYWORDS = [
    "auto rickshaw", "rickshaw", "three wheeler", "3 wheeler", "tempo",
    "e-rickshaw", "erickshaw", "toto", "auto garage", "auto works",
    "auto repair", "auto service", "cng auto",
]


def looks_like_two_wheeler_only(name):
    """True if the name signals a two-wheeler-only shop with no hint of
    doing three-wheeler work - these get dropped from the output."""
    name_lower = name.lower()
    has_two_wheeler_word = any(k in name_lower for k in TWO_WHEELER_EXCLUDE_KEYWORDS)
    has_three_wheeler_word = any(k in name_lower for k in THREE_WHEELER_KEYWORDS)
    return has_two_wheeler_word and not has_three_wheeler_word

# ============================================================
# PIPELINE CODE - shouldn't need to edit below this line
# ============================================================

TEXT_SEARCH_URL = "https://maps.googleapis.com/maps/api/place/textsearch/json"
DETAILS_URL = "https://maps.googleapis.com/maps/api/place/details/json"


def text_search(query, next_page_token=None):
    """Call Places Text Search API. Returns the raw JSON response."""
    params = {"key": API_KEY}
    if next_page_token:
        params["pagetoken"] = next_page_token
    else:
        params["query"] = query
    resp = requests.get(TEXT_SEARCH_URL, params=params, timeout=15)
    resp.raise_for_status()
    return resp.json()


def get_place_details(place_id):
    """Fetch phone number and formatted address for a given place_id."""
    params = {
        "key": API_KEY,
        "place_id": place_id,
        "fields": "name,formatted_phone_number,international_phone_number,"
                   "formatted_address,geometry,rating,user_ratings_total,business_status",
    }
    resp = requests.get(DETAILS_URL, params=params, timeout=15)
    resp.raise_for_status()
    return resp.json().get("result", {})


def search_all_pages(query):
    """Text Search returns up to 20 results per page, max 3 pages (60 results).
    Google requires a short delay before a next_page_token becomes valid."""
    results = []
    data = text_search(query)
    results.extend(data.get("results", []))

    next_token = data.get("next_page_token")
    pages_fetched = 1
    while next_token and pages_fetched < 3:
        time.sleep(2)  # required delay before token is valid
        data = text_search(query, next_page_token=next_token)
        results.extend(data.get("results", []))
        next_token = data.get("next_page_token")
        pages_fetched += 1

    return results


def looks_like_dealer(name):
    name_lower = name.lower()
    return any(keyword in name_lower for keyword in DEALER_KEYWORDS)


def main():
    if not API_KEY:
        print("ERROR: Set the GOOGLE_PLACES_API_KEY environment variable first.")
        print('  export GOOGLE_PLACES_API_KEY="your_key_here"')
        return

    for city, locations in TARGET_LOCATIONS_BY_CITY.items():
        print(f"\n=== {city} ===")
        seen_place_ids = set()
        city_records = []
        skipped_two_wheeler = 0

        for location in locations:
            for term in SEARCH_TERMS:
                query = f"{term} in {location}"
                print(f"Searching: {query}")
                try:
                    results = search_all_pages(query)
                except requests.RequestException as e:
                    print(f"  Request failed: {e}")
                    continue

                print(f"  -> {len(results)} raw results")

                for r in results:
                    place_id = r.get("place_id")
                    if not place_id or place_id in seen_place_ids:
                        continue
                    seen_place_ids.add(place_id)

                    # Filter on the Text Search name before spending a
                    # Details call on a listing that's clearly two-wheeler-only.
                    if looks_like_two_wheeler_only(r.get("name", "")):
                        skipped_two_wheeler += 1
                        continue

                    time.sleep(REQUEST_DELAY)
                    try:
                        details = get_place_details(place_id)
                    except requests.RequestException as e:
                        print(f"  Details fetch failed for {place_id}: {e}")
                        continue

                    if looks_like_two_wheeler_only(details.get("name", "")):
                        skipped_two_wheeler += 1
                        continue

                    name = details.get("name", r.get("name", ""))
                    record = {
                        "name": name,
                        "phone": details.get("formatted_phone_number", ""),
                        "phone_international": details.get("international_phone_number", ""),
                        "address": details.get("formatted_address", r.get("formatted_address", "")),
                        "search_location": location,
                        "search_term": term,
                        "rating": details.get("rating", ""),
                        "num_ratings": details.get("user_ratings_total", ""),
                        "business_status": details.get("business_status", ""),
                        "likely_authorized_dealer": looks_like_dealer(name),
                        "place_id": place_id,
                        "maps_url": f"https://www.google.com/maps/place/?q=place_id:{place_id}",
                    }
                    city_records.append(record)

                time.sleep(REQUEST_DELAY)

        if skipped_two_wheeler:
            print(f"  (skipped {skipped_two_wheeler} listings that looked two-wheeler-only)")
        write_city_csv(city, city_records)


def write_city_csv(city, records):
    """Write one city's records to its own CSV (named per OUTPUT_FILE_TEMPLATE)."""
    output_file = OUTPUT_FILE_TEMPLATE.format(city=city.lower().replace(" ", "_"))

    if not records:
        print(f"No records collected for {city}.")
        return

    fieldnames = list(records[0].keys())
    with open(output_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(records)

    print(f"Done with {city}. {len(records)} unique listings written to {output_file}")
    with_phone = sum(1 for r in records if r["phone"])
    print(f"  {with_phone} of them have a phone number.")
    likely_independent = sum(1 for r in records if not r["likely_authorized_dealer"])
    print(f"  {likely_independent} flagged as likely independent (not authorized dealers).")


if __name__ == "__main__":
    main()
