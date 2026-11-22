import os
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
from streamlit.errors import StreamlitSecretNotFoundError


API_KEY_ENV = "API_FOOTBALL_KEY"
BASE_URL = "https://v3.football.api-sports.io/fixtures"
OUTPUT_PATH = Path(__file__).resolve().parent / "api_tum_ligler.xlsx"
SEASON = 2026
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
FINISHED_STATUSES = {"FT", "AET", "PEN"}
SCHEDULED_STATUSES = {"NS", "TBD", "PST"}
LEAGUES = [
    {"id": 98, "code": "J1", "name": "Japonya J1 League"},
    {"id": 292, "code": "K1", "name": "Güney Kore K League 1"},
    {"id": 301, "code": "UAE", "name": "BAE Pro League"},
    {"id": 233, "code": "EGY", "name": "Mısır Premier League"},
    {"id": 288, "code": "RSA", "name": "Güney Afrika Premier Division"},
    {"id": 702, "code": "PL2", "name": "İngiltere Premier League 2 U21"},
    {"id": 48, "code": "UYL", "name": "UEFA Youth League U19"},
]


def get_api_key():
    api_key = os.environ.get(API_KEY_ENV, "").strip()
    if api_key:
        return api_key

    try:
        api_key = str(st.secrets.get(API_KEY_ENV, "") or "").strip()
    except StreamlitSecretNotFoundError:
        api_key = ""

    if not api_key:
        api_key = input("API_FOOTBALL_KEY değerini girin: ").strip()
    if not api_key:
        raise RuntimeError("API_FOOTBALL_KEY boş olamaz.")
    return api_key


def get_fixture_page(league_id, page, api_key):
    response = requests.get(
        BASE_URL,
        headers={"x-apisports-key": api_key},
        params={"league": league_id, "season": SEASON, "page": page},
        timeout=30,
    )
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        raise RuntimeError(
            f"Lig ID {league_id} API isteği HTTP {response.status_code} hatası verdi: "
            f"{response.text}"
        ) from exc

    try:
        payload = response.json()
    except ValueError as exc:
        raise ValueError(f"Lig ID {league_id} API yanıtı geçerli JSON değil.") from exc

    if not isinstance(payload, dict):
        raise ValueError(f"Lig ID {league_id} API yanıt biçimi geçersiz.")

    errors = payload.get("errors")
    if errors:
        raise RuntimeError(f"Lig ID {league_id} API hatası: {errors}")

    fixtures = payload.get("response")
    if not isinstance(fixtures, list):
        raise ValueError(f"Lig ID {league_id} yanıtında maç listesi bulunamadı.")

    paging = payload.get("paging") or {}
    try:
        total_pages = max(1, int(paging.get("total", 1)))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Lig ID {league_id} sayfalama bilgisi geçersiz.") from exc
    return fixtures, total_pages


def fetch_league_fixtures(league, api_key):
    fixtures, total_pages = get_fixture_page(league["id"], 1, api_key)
    all_fixtures = list(fixtures)

    for page in range(2, total_pages + 1):
        fixtures, _ = get_fixture_page(league["id"], page, api_key)
        all_fixtures.extend(fixtures)

    return all_fixtures


def fixture_to_row(league, fixture, now):
    fixture_info = fixture.get("fixture") or {}
    status_info = fixture_info.get("status") or {}
    status_code = status_info.get("short")
    kickoff = pd.to_datetime(fixture_info.get("date"), errors="coerce", utc=True)
    if pd.isna(kickoff):
        raise ValueError(f"Lig ID {league['id']} maçında geçersiz tarih var.")

    if status_code in FINISHED_STATUSES:
        match_status = "FINISHED"
    elif status_code in SCHEDULED_STATUSES and kickoff >= now:
        match_status = "SCHEDULED"
    else:
        return None

    teams = fixture.get("teams") or {}
    home_team = teams.get("home") or {}
    away_team = teams.get("away") or {}
    home_name = home_team.get("name")
    away_name = away_team.get("name")
    if not home_name or not away_name:
        raise ValueError(f"Lig ID {league['id']} maçında takım adı eksik.")

    goals = fixture.get("goals") or {}
    home_goals = goals.get("home")
    away_goals = goals.get("away")
    winner = None
    if match_status == "FINISHED":
        if home_goals is None or away_goals is None:
            raise ValueError(
                f"{home_name} - {away_name} bitmiş görünüyor ancak skor eksik."
            )
        if home_goals > away_goals:
            winner = "HOME_TEAM"
        elif away_goals > home_goals:
            winner = "AWAY_TEAM"
        else:
            winner = "DRAW"

    return {
        "Lig": league["code"],
        "Tarih": kickoff.isoformat(),
        "Durum": match_status,
        "EvSahibi": home_name,
        "Deplasman": away_name,
        "EvGol": home_goals,
        "DepGol": away_goals,
        "Kazanan": winner,
    }


def update_workbook(rows):
    if OUTPUT_PATH.is_file():
        existing = pd.read_excel(OUTPUT_PATH)
        missing_columns = set(OUTPUT_COLUMNS).difference(existing.columns)
        if missing_columns:
            raise ValueError(
                "api_tum_ligler.xlsx dosyasında gerekli sütunlar eksik: "
                + ", ".join(sorted(missing_columns))
            )
        existing = existing[OUTPUT_COLUMNS].copy()
    else:
        existing = pd.DataFrame(columns=OUTPUT_COLUMNS)

    league_codes = {league["code"] for league in LEAGUES}
    existing = existing.loc[
        ~existing["Lig"].astype(str).isin(league_codes)
    ].copy()
    refreshed = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    combined = pd.concat([existing, refreshed], ignore_index=True)
    combined = combined.drop_duplicates(
        subset=["Lig", "Tarih", "EvSahibi", "Deplasman"],
        keep="last",
    )
    combined.to_excel(OUTPUT_PATH, index=False)
    return len(existing), len(refreshed), len(combined)


def main():
    api_key = get_api_key()
    now = pd.Timestamp(datetime.now(timezone.utc))
    rows = []

    for league in LEAGUES:
        fixtures = fetch_league_fixtures(league, api_key)
        league_rows = []
        for fixture in fixtures:
            row = fixture_to_row(league, fixture, now)
            if row is not None:
                league_rows.append(row)
        rows.extend(league_rows)
        print(
            f"{league['name']}: {len(league_rows)} bitmiş veya gelecek maç alındı."
        )

    preserved_count, refreshed_count, total_count = update_workbook(rows)
    print(
        f"Tamamlandı: {refreshed_count} maç güncellendi; "
        f"{preserved_count} diğer kayıt korundu. "
        f"Dosyada toplam {total_count} maç var: {OUTPUT_PATH.name}"
    )


if __name__ == "__main__":
    main()
