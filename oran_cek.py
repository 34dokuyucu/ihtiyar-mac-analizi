import json
import math
import os
import statistics
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd
import streamlit as st
from streamlit.errors import StreamlitSecretNotFoundError


API_KEY_ENV = "ODDS_API_KEY"
BASE_URL = "https://api.the-odds-api.com/v4/sports/{sport}/odds/"
OUTPUT_NAME = "piyasa_oranlari.xlsx"
REGION = "eu"
MARKETS = "h2h,totals"
SPORTS = [
    "soccer_turkey_super_league",
    "soccer_epl",
    "soccer_spain_la_liga",
    "soccer_italy_serie_a",
    "soccer_germany_bundesliga",
    "soccer_france_ligue_one",
    "soccer_uefa_champs_league",
]
OUTPUT_COLUMNS = [
    "EvSahibi",
    "Deplasman",
    "Ort_MS1",
    "Ort_MSX",
    "Ort_MS2",
    "Ort_Ust25",
]


def get_api_key():
    api_key = os.environ.get(API_KEY_ENV, "").strip()
    if api_key:
        return api_key

    try:
        return str(st.secrets.get(API_KEY_ENV, "")).strip()
    except StreamlitSecretNotFoundError:
        return ""


def fetch_odds(sport, api_key):
    query = urlencode(
        {
            "apiKey": api_key,
            "regions": REGION,
            "markets": MARKETS,
            "oddsFormat": "decimal",
            "dateFormat": "iso",
        }
    )
    request = Request(
        f"{BASE_URL.format(sport=sport)}?{query}",
        headers={
            "Accept": "application/json",
            "User-Agent": "football-odds-export/1.0",
        },
    )
    try:
        with urlopen(request, timeout=30) as response:
            data = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"{sport} icin Odds API HTTP {exc.code} hatasi verdi: {detail}"
        ) from exc
    except URLError as exc:
        raise RuntimeError(
            f"{sport} icin Odds API istegi basarisiz: {exc.reason}"
        ) from exc

    if not isinstance(data, list):
        raise ValueError(f"{sport} API yaniti mac listesi degil.")
    return data


def valid_price(value):
    try:
        price = float(value)
    except (TypeError, ValueError):
        return None
    return price if math.isfinite(price) and price > 0 else None


def normalized_name(value):
    return " ".join(str(value).strip().casefold().split())


def average_event_odds(event):
    home_name = event.get("home_team")
    away_name = event.get("away_team")
    if not home_name or not away_name:
        raise ValueError("API mac kaydinda ev/deplasman takim adi eksik.")

    prices = {"home": [], "draw": [], "away": [], "over25": []}
    home_key = normalized_name(home_name)
    away_key = normalized_name(away_name)

    for bookmaker in event.get("bookmakers", []):
        for market in bookmaker.get("markets", []):
            outcomes = market.get("outcomes", [])
            if market.get("key") == "h2h":
                for outcome in outcomes:
                    outcome_name = normalized_name(outcome.get("name", ""))
                    price = valid_price(outcome.get("price"))
                    if price is None:
                        continue
                    if outcome_name == home_key:
                        prices["home"].append(price)
                    elif outcome_name == away_key:
                        prices["away"].append(price)
                    elif outcome_name in {"draw", "tie"}:
                        prices["draw"].append(price)

            elif market.get("key") == "totals":
                for outcome in outcomes:
                    if normalized_name(outcome.get("name", "")) != "over":
                        continue
                    try:
                        point = float(outcome.get("point"))
                    except (TypeError, ValueError):
                        continue
                    if not math.isclose(point, 2.5, abs_tol=1e-9):
                        continue
                    price = valid_price(outcome.get("price"))
                    if price is not None:
                        prices["over25"].append(price)

    def mean_or_blank(values):
        return statistics.fmean(values) if values else None

    return {
        "EvSahibi": home_name,
        "Deplasman": away_name,
        "Ort_MS1": mean_or_blank(prices["home"]),
        "Ort_MSX": mean_or_blank(prices["draw"]),
        "Ort_MS2": mean_or_blank(prices["away"]),
        "Ort_Ust25": mean_or_blank(prices["over25"]),
    }


def main():
    api_key = get_api_key()
    if not api_key:
        raise RuntimeError(
            "ODDS_API_KEY ortam değişkeninde veya Streamlit secrets içinde tanımlı değil."
        )

    rows = []
    for sport in SPORTS:
        events = fetch_odds(sport, api_key)
        rows.extend(average_event_odds(event) for event in events)
        print(f"{sport}: {len(events)} mac alindi.")

    output_path = Path(__file__).resolve().parent / OUTPUT_NAME
    pd.DataFrame(rows, columns=OUTPUT_COLUMNS).to_excel(output_path, index=False)
    print(f"Toplam {len(rows)} mac orani '{OUTPUT_NAME}' dosyasina kaydedildi.")


if __name__ == "__main__":
    main()
