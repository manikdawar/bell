"""
Two-Wheeler Mechanic Data Collection Pipeline
===============================================
Collects business name, phone number, address, and location for independent
two-wheeler / bike repair mechanics in target Indian cities, using the
official Google Places API (Text Search + Place Details).

Target locations, search terms, and dealer keywords are all read from text
files (not hardcoded here) - edit two_wheeler_locations.txt or
two_wheeler_search_config.txt to change what gets searched for, without
touching this script. Cities covered by default: Saharanpur, Delhi NCR,
Chandigarh, Muzaffarnagar, Meerut, Ambala, Ludhiana, Jullundur (Jalandhar),
Amritsar, Jammu.

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

4. Edit two_wheeler_locations.txt with your target cities/localities, and
   two_wheeler_search_config.txt with your search terms / dealer keywords.

5. Run:
   python collect_two_wheeler_mechanics.py

Output: one CSV per city (e.g. mechanics_two_wheeler_saharanpur.csv),
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

# Locations file: "== City Name ==" headers followed by one locality per
# line. See two_wheeler_locations.txt for the format and current contents.
LOCATIONS_FILE = "two_wheeler_locations.txt"

# Search config file: "== Search Terms ==" and "== Dealer Keywords =="
# sections, one entry per line. See two_wheeler_search_config.txt for the
# format and current contents.
SEARCH_CONFIG_FILE = "two_wheeler_search_config.txt"

# Delay between API calls (seconds) - be a good API citizen.
REQUEST_DELAY = 0.2

# One CSV per city, written to the same folder. {city} is replaced with a
# lowercase, underscore-separated slug of the city name.
OUTPUT_FILE_TEMPLATE = "mechanics_two_wheeler_{city}.csv"

# ============================================================
# PIPELINE CODE - shouldn't need to edit below this line
# ============================================================

TEXT_SEARCH_URL = "https://maps.googleapis.com/maps/api/place/textsearch/json"
DETAILS_URL = "https://maps.googleapis.com/maps/api/place/details/json"


def parse_sectioned_file(path):
    """Parse a "== Section Name ==" delimited text file into
    {section_name: [lines]}, in file order. Used for both the locations
    file (sections = cities) and the search config file (sections =
    "Search Terms" / "Dealer Keywords"). Blank lines and lines starting
    with # are ignored.
    """
    sections = {}
    current_section = None
    with open(path, encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("==") and line.endswith("=="):
                current_section = line.strip("=").strip()
                sections.setdefault(current_section, [])
                continue
            if current_section is None:
                raise ValueError(
                    f"Line found before any '== Section ==' header in "
                    f"{path}: {line!r}"
                )
            sections[current_section].append(line)
    return sections


def load_locations(path):
    """Parse a locations file into {city: [locations]}. See
    two_wheeler_locations.txt for the format and current contents."""
    return parse_sectioned_file(path)


def load_search_config(path):
    """Parse the search config file into (search_terms, dealer_keywords).
    See two_wheeler_search_config.txt for the format and current contents."""
    sections = parse_sectioned_file(path)
    search_terms = sections.get("Search Terms", [])
    dealer_keywords = sections.get("Dealer Keywords", [])
    return search_terms, dealer_keywords


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


def looks_like_dealer(name, dealer_keywords):
    name_lower = name.lower()
    return any(keyword.lower() in name_lower for keyword in dealer_keywords)


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


def main():
    if not API_KEY:
        print("ERROR: Set the GOOGLE_PLACES_API_KEY environment variable first.")
        print('  export GOOGLE_PLACES_API_KEY="your_key_here"')
        return

    try:
        target_locations_by_city = load_locations(LOCATIONS_FILE)
        search_terms, dealer_keywords = load_search_config(SEARCH_CONFIG_FILE)
    except FileNotFoundError as e:
        print(f"ERROR: File not found: {e.filename}")
        return
    except ValueError as e:
        print(f"ERROR: {e}")
        return

    if not target_locations_by_city:
        print(f"No cities found in {LOCATIONS_FILE}.")
        return
    if not search_terms:
        print(f"No search terms found in {SEARCH_CONFIG_FILE} "
              f"(expected a '== Search Terms ==' section).")
        return

    for city, locations in target_locations_by_city.items():
        print(f"\n=== {city} ===")
        seen_place_ids = set()
        city_records = []

        for location in locations:
            for term in search_terms:
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

                    time.sleep(REQUEST_DELAY)
                    try:
                        details = get_place_details(place_id)
                    except requests.RequestException as e:
                        print(f"  Details fetch failed for {place_id}: {e}")
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
                        "likely_authorized_dealer": looks_like_dealer(name, dealer_keywords),
                        "place_id": place_id,
                        "maps_url": f"https://www.google.com/maps/place/?q=place_id:{place_id}",
                    }
                    city_records.append(record)

                time.sleep(REQUEST_DELAY)

        write_city_csv(city, city_records)


if __name__ == "__main__":
    main()
