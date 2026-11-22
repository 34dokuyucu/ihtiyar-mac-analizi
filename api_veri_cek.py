import json
import os
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import pandas as pd
import streamlit as st
from streamlit.errors import StreamlitSecretNotFoundError


API_KEY_ENV = "FOOTBALL_DATA_API_KEY"
BASE_URL = "https://api.football-data.org/v4/competitions"
OUTPUT_NAME = "api_tum_ligler.xlsx"
COMPETITIONS = ["PL", "ELC", "PD", "SA", "BL1", "FL1", "DED", "PPL", "CL"]
OUTPUT_COLUMNS = [
    "Lig",
    "Tarih",
    "Durum",
    "EvSahibi",
    "Deplasman",
    "EvGol",
    "DepGol",
    "Kazanan",
]


def get_api_key():
    api_key = os.environ.get(API_KEY_ENV, "").strip()
    if api_key:
        return api_key

    try:
        return str(st.secrets.get(API_KEY_ENV, "")).strip()
    except StreamlitSecretNotFoundError:
        return ""


def get_matches(competition_code, api_key):
    url = f"{BASE_URL}/{competition_code}/matches"
    request = Request(
        url,
        headers={
            "X-Auth-Token": api_key,
            "Accept": "application/json",
            "User-Agent": "football-match-export/1.0",
        },
    )
    try:
        with urlopen(request, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"{competition_code} API istegi HTTP {exc.code} hatasi verdi: {detail}"
        ) from exc
    except URLError as exc:
        raise RuntimeError(
            f"{competition_code} API istegi basarisiz: {exc.reason}"
        ) from exc


def to_row(competition_code, match):
    score = match.get("score") or {}
    full_time = score.get("fullTime") or {}
    home_team = match.get("homeTeam") or {}
    away_team = match.get("awayTeam") or {}
    status = match.get("status")

    return {
        "Lig": competition_code,
        "Tarih": match.get("utcDate"),
        "Durum": status,
        "EvSahibi": home_team.get("name"),
        "Deplasman": away_team.get("name"),
        "EvGol": full_time.get("home"),
        "DepGol": full_time.get("away"),
        "Kazanan": score.get("winner"),
    }


def main():
    api_key = get_api_key()
    if not api_key:
        raise RuntimeError(
            "FOOTBALL_DATA_API_KEY ortam değişkeninde veya Streamlit secrets içinde tanımlı değil."
        )

    rows = []
    for index, competition_code in enumerate(COMPETITIONS):
        payload = get_matches(competition_code, api_key)
        matches = payload.get("matches")
        if not isinstance(matches, list):
            raise ValueError(
                f"{competition_code} API yanitinda 'matches' listesi bulunamadi."
            )

        rows.extend(to_row(competition_code, match) for match in matches)
        print(f"{competition_code}: {len(matches)} mac alindi.")

        if index < len(COMPETITIONS) - 1:
            time.sleep(7)

    output_path = Path(__file__).resolve().parent / OUTPUT_NAME
    pd.DataFrame(rows, columns=OUTPUT_COLUMNS).to_excel(output_path, index=False)
    print(f"Toplam {len(rows)} mac '{OUTPUT_NAME}' dosyasina kaydedildi.")


if __name__ == "__main__":
    main()
