# Two-Wheeler Mechanic Data Collection

Collects name, phone number, and address for two-wheeler mechanics in
target Indian locations, using the official Google Places API.

## Setup (do this once)

1. **Get a Google Places API key**
   - Go to https://console.cloud.google.com/
   - Create a project (or use an existing one)
   - Enable the **Places API**
   - Go to Credentials -> Create Credentials -> API Key
   - Google gives new accounts free monthly credit, but keep an eye on the
     billing dashboard so you don't get an unexpected bill.

2. **Install Python dependency**
   ```
   pip install requests --break-system-packages
   ```

3. **Set your API key** as an environment variable so it's not hardcoded
   into the script:
   ```
   export GOOGLE_PLACES_API_KEY="your_key_here"
   ```

4. **Edit `collect_mechanics.py`**
   - `TARGET_LOCATIONS`: list the cities/neighborhoods you want to cover.
     Being specific (e.g. "Andheri, Mumbai" instead of just "Mumbai") gets
     better, less-clipped results since each query only returns up to 60
     results.
   - `SEARCH_TERMS`: phrasings to search for. Add more if you find gaps
     (e.g. "puncture shop", "vehicle service center").

## Run it

```
python collect_mechanics.py
```

This writes `mechanics_output.csv` with columns: name, phone, address,
rating, an "likely_authorized_dealer" flag (heuristic based on the name -
check it manually, it won't be perfect), and a Google Maps link for
each listing.

## Notes on scaling up

- Each location x search term combination costs a handful of API calls
  (1 Text Search call per page, up to 3 pages, + 1 Details call per unique
  result). Budget accordingly if you're covering many cities.
- Phone numbers for small businesses in India are often personal mobile
  numbers. If you plan to contact these numbers for marketing/outreach,
  India's TRAI regulations (DND registry, commercial communication rules)
  apply - keep messaging clearly identified, low-frequency, and easy to
  opt out of.
- The "likely_authorized_dealer" flag is a simple keyword heuristic on the
  business name. It will misclassify some listings - treat it as a
  starting filter, not ground truth.
- For denser coverage of a single large city, consider replacing the
  "Text Search by city name" approach with a grid of Nearby Search calls
  across lat/lng points - happy to add that if you need it.
