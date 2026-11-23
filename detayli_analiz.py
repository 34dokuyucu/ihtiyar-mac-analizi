import json
import math
import os
import re
import unicodedata
from datetime import timedelta
from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd
import requests
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


INPUT_NAME = "api_tum_ligler.xlsx"
REPORT_NAME = "haftalik_analiz_raporu.xlsx"
CHART_NAME = "lig_karsilastirma.png"
H2H_CACHE_NAME = "h2h_gecmis.xlsx"
API_FOOTBALL_KEY_ENV = "API_FOOTBALL_KEY"
API_FOOTBALL_BASE_URL = "https://v3.football.api-sports.io"
REQUIRED_COLUMNS = {
    "Lig",
    "Tarih",
    "Durum",
    "EvSahibi",
    "Deplasman",
    "EvGol",
    "DepGol",
}
UPCOMING_COLUMNS = [
    "Tarih (UTC)",
    "Lig",
    "Durum",
    "Ev Sahibi",
    "Ev Form (G/B/M)",
    "Ev Form Mac",
    "Ev Son 5 Puan",
    "Ev Attigi Gol Ort. (Son 5)",
    "Ev Yedigi Gol Ort. (Son 5)",
    "Ev Attigi Gol Ort. (Sezon)",
    "Ev Yedigi Gol Ort. (Sezon)",
    "Ev Oynanan Maç",
    "Ev Sürpriz Yapma %",
    "Ev Sürpriz Yeme %",
    "Ev Sürpriz Yapma Sayısı",
    "Ev Sürpriz Yeme Sayısı",
    "Ev İlk Yarı Gol Ort.",
    "Ev İlk Yarı Veri Sayısı",
    "Ev İkinci Yarı Gol Ort.",
    "Ev İkinci Yarı Veri Sayısı",
    "Deplasman",
    "Dep Form (G/B/M)",
    "Dep Form Mac",
    "Dep Son 5 Puan",
    "Dep Attigi Gol Ort. (Son 5)",
    "Dep Yedigi Gol Ort. (Son 5)",
    "Dep Attigi Gol Ort. (Sezon)",
    "Dep Yedigi Gol Ort. (Sezon)",
    "Dep Oynanan Maç",
    "Dep Sürpriz Yapma %",
    "Dep Sürpriz Yeme %",
    "Dep Sürpriz Yapma Sayısı",
    "Dep Sürpriz Yeme Sayısı",
    "Dep İlk Yarı Gol Ort.",
    "Dep İlk Yarı Veri Sayısı",
    "Dep İkinci Yarı Gol Ort.",
    "Dep İkinci Yarı Veri Sayısı",
    "Sezon Takim Gol Ort. Toplami",
    "Son 5 Takim Gol Ort. Toplami",
    "Form Avantajli Takim",
    "Istatistiksel_Etiketler",
    "Lig_IcSaha_Gol_Ort",
    "Lig_DisSaha_Gol_Ort",
    "Ev_IcSaha_Hucum_Gucu",
    "Ev_IcSaha_Savunma_Gucu",
    "Dep_DisSaha_Hucum_Gucu",
    "Dep_DisSaha_Savunma_Gucu",
    "Ev_xG_Lambda",
    "Dep_xG_Lambda",
    "MS1_%",
    "MSX_%",
    "MS2_%",
    "Ust_25_%",
    "KG_Var_%",
    "MS1_ve_Ust25_%",
    "MS2_ve_Ust25_%",
    "MS1_ve_KG_Var_%",
    "MS2_ve_KG_Var_%",
    "Ust25_ve_KG_Var_%",
    "Ust_25_KG_Var_%",
    "En_Olasi_Skor",
    "Korner_Tahmini",
    "Kart_Tahmini",
    "Oyuncu_Tahmini",
    "H2H_Gecmis_JSON",
    "H2H_API_Durumu",
]
H2H_MATCH_COLUMNS = [
    "PairKey",
    "TarihUTC",
    "Lig",
    "EvSahibi",
    "Deplasman",
    "EvGol",
    "DepGol",
    "Kaynak",
]
H2H_QUERY_COLUMNS = ["PairKey", "Durum", "SorguUTC", "Hata"]
TEAM_ID_COLUMNS = ["TeamKey", "Lig", "Takim", "TeamID"]
API_LEAGUE_IDS = {
    "TSL": 203,
    "T1L": 204,
    "PL": 39,
    "ELC": 40,
    "PD": 140,
    "SA": 135,
    "BL1": 78,
    "FL1": 61,
    "DED": 88,
    "PPL": 94,
    "CL": 2,
    "J1": 98,
    "K1": 292,
    "UAE": 301,
    "EGY": 233,
    "RSA": 288,
    "PL2": 702,
    "UYL": 48,
}
HALF_TIME_HOME_COLUMNS = ("IlkYariEvGol", "HTEvGol", "HTHG", "HomeHTGoals")
HALF_TIME_AWAY_COLUMNS = ("IlkYariDepGol", "HTDepGol", "HTAG", "AwayHTGoals")
PERCENTAGE_COLUMNS = {
    "Ev Sahibi Kazanma %",
    "2.5 Ust %",
    "KG Var %",
    "MS1_%",
    "MSX_%",
    "MS2_%",
    "Ust_25_%",
    "KG_Var_%",
    "MS1_ve_Ust25_%",
    "MS2_ve_Ust25_%",
    "MS1_ve_KG_Var_%",
    "MS2_ve_KG_Var_%",
    "Ust25_ve_KG_Var_%",
    "Ust_25_KG_Var_%",
    "Ev Sürpriz Yapma %",
    "Ev Sürpriz Yeme %",
    "Dep Sürpriz Yapma %",
    "Dep Sürpriz Yeme %",
}


def build_league_summary(finished):
    if finished.empty:
        return pd.DataFrame(
            columns=[
                "Lig",
                "Oynanan Mac",
                "Mac Basi Gol Ort.",
                "Ev Sahibi Kazanma %",
                "2.5 Ust %",
                "KG Var %",
            ]
        )

    matches = finished.copy()
    matches["ToplamGol"] = matches["EvGol"] + matches["DepGol"]
    matches["EvSahibiKazandi"] = matches["EvGol"] > matches["DepGol"]
    matches["Ust25"] = matches["ToplamGol"] > 2.5
    matches["KGVar"] = (matches["EvGol"] > 0) & (matches["DepGol"] > 0)
    summary = (
        matches.groupby("Lig", sort=True)
        .agg(
            OynananMac=("Lig", "size"),
            MacBasiGolOrt=("ToplamGol", "mean"),
            EvSahibiKazanma=("EvSahibiKazandi", "mean"),
            Ust25=("Ust25", "mean"),
            KGVar=("KGVar", "mean"),
        )
        .reset_index()
        .rename(
            columns={
                "OynananMac": "Oynanan Mac",
                "MacBasiGolOrt": "Mac Basi Gol Ort.",
                "EvSahibiKazanma": "Ev Sahibi Kazanma %",
                "Ust25": "2.5 Ust %",
                "KGVar": "KG Var %",
            }
        )
    )
    return summary


def build_standings(finished):
    columns = [
        "Lig",
        "Saha",
        "Sira",
        "Takim",
        "O",
        "G",
        "B",
        "M",
        "AG",
        "YG",
        "AV",
        "Puan",
    ]
    if finished.empty:
        return pd.DataFrame(columns=columns)

    stats = {}
    teams_by_league = {}
    for match in finished.itertuples(index=False):
        league = match.Lig
        home, away = match.EvSahibi, match.Deplasman
        home_goals, away_goals = int(match.EvGol), int(match.DepGol)
        teams_by_league.setdefault(league, set()).update((home, away))

        for venue, team, goals_for, goals_against in (
            ("Ic Saha", home, home_goals, away_goals),
            ("Deplasman", away, away_goals, home_goals),
        ):
            key = (league, venue, team)
            row = stats.setdefault(
                key, {"O": 0, "G": 0, "B": 0, "M": 0, "AG": 0, "YG": 0, "Puan": 0}
            )
            row["O"] += 1
            row["AG"] += goals_for
            row["YG"] += goals_against
            if goals_for > goals_against:
                row["G"] += 1
                row["Puan"] += 3
            elif goals_for == goals_against:
                row["B"] += 1
                row["Puan"] += 1
            else:
                row["M"] += 1

    rows = []
    for league, teams in teams_by_league.items():
        for venue in ("Ic Saha", "Deplasman"):
            for team in teams:
                row = stats.get(
                    (league, venue, team),
                    {"O": 0, "G": 0, "B": 0, "M": 0, "AG": 0, "YG": 0, "Puan": 0},
                )
                rows.append(
                    {
                        "Lig": league,
                        "Saha": venue,
                        "Takim": team,
                        **row,
                        "AV": row["AG"] - row["YG"],
                    }
                )

    standings = pd.DataFrame(rows)
    standings = standings.sort_values(
        ["Lig", "Saha", "Puan", "AV", "AG", "Takim"],
        ascending=[True, True, False, False, False, True],
        kind="stable",
    )
    standings["Sira"] = standings.groupby(["Lig", "Saha"]).cumcount() + 1
    return standings[columns].reset_index(drop=True)


def normalize_team_name(value):
    translation = str.maketrans(
        {"ı": "i", "İ": "i", "ş": "s", "Ş": "s", "ğ": "g", "Ğ": "g",
         "ü": "u", "Ü": "u", "ö": "o", "Ö": "o", "ç": "c", "Ç": "c"}
    )
    normalized = unicodedata.normalize(
        "NFKD", str(value).translate(translation)
    )
    without_marks = "".join(
        character for character in normalized
        if not unicodedata.combining(character)
    )
    return re.sub(r"[^a-z0-9]+", " ", without_marks.casefold()).strip()


def pair_cache_key(home_team, away_team):
    return "||".join(sorted((normalize_team_name(home_team), normalize_team_name(away_team))))


def get_api_football_key():
    api_key = os.environ.get(API_FOOTBALL_KEY_ENV, "").strip()
    if api_key:
        return api_key

    try:
        import streamlit as st
        from streamlit.errors import StreamlitSecretNotFoundError

        return str(st.secrets.get(API_FOOTBALL_KEY_ENV, "") or "").strip()
    except StreamlitSecretNotFoundError:
        return ""


def find_column(frame, candidates):
    columns = {str(column).casefold(): column for column in frame.columns}
    for candidate in candidates:
        found = columns.get(candidate.casefold())
        if found is not None:
            return found
    return None


def build_team_history(finished):
    history = {}
    standings = {}
    if finished.empty:
        return history

    ordered = finished.loc[finished["TarihUTC"].notna()].sort_values(
        ["Lig", "TarihUTC"], kind="stable"
    )
    home_half_column = find_column(finished, HALF_TIME_HOME_COLUMNS)
    away_half_column = find_column(finished, HALF_TIME_AWAY_COLUMNS)
    points_by_result = {"G": 3, "B": 1, "M": 0}

    for (league, kickoff), same_time_matches in ordered.groupby(
        ["Lig", "TarihUTC"], sort=True, dropna=True
    ):
        pending_updates = []
        for match in same_time_matches.itertuples(index=False):
            home_key = (league, match.EvSahibi)
            away_key = (league, match.Deplasman)
            home_played, home_points = standings.get(home_key, (0, 0))
            away_played, away_points = standings.get(away_key, (0, 0))
            home_goals, away_goals = int(match.EvGol), int(match.DepGol)

            if home_goals > away_goals:
                home_result, away_result = "G", "M"
            elif home_goals < away_goals:
                home_result, away_result = "M", "G"
            else:
                home_result = away_result = "B"

            comparable = home_played > 0 and away_played > 0
            home_underdog_win = (
                comparable
                and home_result == "G"
                and home_points / home_played < away_points / away_played
            )
            away_underdog_win = (
                comparable
                and away_result == "G"
                and away_points / away_played < home_points / home_played
            )

            home_half_goals = (
                pd.to_numeric(
                    pd.Series([getattr(match, home_half_column)]),
                    errors="coerce",
                ).iloc[0]
                if home_half_column is not None
                else None
            )
            away_half_goals = (
                pd.to_numeric(
                    pd.Series([getattr(match, away_half_column)]),
                    errors="coerce",
                ).iloc[0]
                if away_half_column is not None
                else None
            )
            valid_halftime = (
                pd.notna(home_half_goals)
                and pd.notna(away_half_goals)
                and 0 <= home_half_goals <= home_goals
                and 0 <= away_half_goals <= away_goals
            )
            if not valid_halftime:
                home_half_goals = away_half_goals = None

            for team, result, scored, conceded, upset_made, upset_conceded, first_half in (
                (
                    match.EvSahibi,
                    home_result,
                    home_goals,
                    away_goals,
                    home_underdog_win,
                    comparable and home_result == "M" and home_points / home_played > away_points / away_played,
                    home_half_goals,
                ),
                (
                    match.Deplasman,
                    away_result,
                    away_goals,
                    home_goals,
                    away_underdog_win,
                    comparable and away_result == "M" and away_points / away_played > home_points / home_played,
                    away_half_goals,
                ),
            ):
                history.setdefault((league, team), []).append(
                    {
                        "timestamp": kickoff,
                        "result": result,
                        "scored": scored,
                        "conceded": conceded,
                        "upset_made": bool(upset_made),
                        "upset_conceded": bool(upset_conceded),
                        "first_half_scored": first_half,
                        "second_half_scored": (
                            scored - first_half
                            if first_half is not None
                            else None
                        ),
                    }
                )

            pending_updates.extend(
                [
                    (home_key, home_played + 1, home_points + points_by_result[home_result]),
                    (away_key, away_played + 1, away_points + points_by_result[away_result]),
                ]
            )

        for key, played, points in pending_updates:
            standings[key] = (played, points)

    for matches in history.values():
        matches.sort(key=lambda item: item["timestamp"])
    return history


def get_team_stats(history, league, team, kickoff):
    season_matches = [
        match
        for match in history.get((league, team), [])
        if match["timestamp"] < kickoff
    ]
    recent = season_matches[-5:]
    form = "-".join(match["result"] for match in recent) if recent else "Yok"
    recent_points = sum(
        {"G": 3, "B": 1, "M": 0}[match["result"]] for match in recent
    )

    def average(matches, key):
        values = [match[key] for match in matches if match[key] is not None]
        return sum(values) / len(values) if values else None

    upset_made_count = sum(match["upset_made"] for match in season_matches)
    upset_conceded_count = sum(match["upset_conceded"] for match in season_matches)
    first_half_count = sum(
        match["first_half_scored"] is not None for match in season_matches
    )
    second_half_count = sum(
        match["second_half_scored"] is not None for match in season_matches
    )
    match_count = len(season_matches)
    return {
        "form": form,
        "recent_count": len(recent),
        "recent_points": recent_points,
        "recent_scored": average(recent, "scored"),
        "recent_conceded": average(recent, "conceded"),
        "season_count": match_count,
        "season_scored": average(season_matches, "scored"),
        "season_conceded": average(season_matches, "conceded"),
        "upset_made_count": upset_made_count,
        "upset_conceded_count": upset_conceded_count,
        "upset_made_rate": upset_made_count / match_count if match_count else None,
        "upset_conceded_rate": (
            upset_conceded_count / match_count if match_count else None
        ),
        "first_half_count": first_half_count,
        "first_half_average": average(season_matches, "first_half_scored"),
        "second_half_count": second_half_count,
        "second_half_average": average(season_matches, "second_half_scored"),
    }


def empty_h2h_cache(cache_path):
    return {
        "path": cache_path,
        "matches": pd.DataFrame(columns=H2H_MATCH_COLUMNS),
        "queries": pd.DataFrame(columns=H2H_QUERY_COLUMNS),
        "team_ids": pd.DataFrame(columns=TEAM_ID_COLUMNS),
    }


def load_h2h_cache(cache_path):
    cache = empty_h2h_cache(cache_path)
    if not cache_path.is_file():
        return cache

    workbook = pd.ExcelFile(cache_path)
    sheets = {
        "H2H_Maclar": ("matches", H2H_MATCH_COLUMNS),
        "Sorgular": ("queries", H2H_QUERY_COLUMNS),
        "Takim_ID": ("team_ids", TEAM_ID_COLUMNS),
    }
    for sheet_name, (cache_key, columns) in sheets.items():
        if sheet_name in workbook.sheet_names:
            frame = pd.read_excel(cache_path, sheet_name=sheet_name)
            missing = set(columns).difference(frame.columns)
            if missing:
                raise ValueError(
                    f"{H2H_CACHE_NAME}/{sheet_name} sütunları eksik: "
                    + ", ".join(sorted(missing))
                )
            cache[cache_key] = frame[columns].copy()
    return cache


def save_h2h_cache(cache):
    matches = cache["matches"].copy()
    if not matches.empty:
        matches["TarihUTC"] = pd.to_datetime(
            matches["TarihUTC"], errors="coerce", utc=True
        ).dt.tz_localize(None)
    queries = cache["queries"].copy()
    if not queries.empty:
        queries["SorguUTC"] = pd.to_datetime(
            queries["SorguUTC"], errors="coerce", utc=True
        ).dt.tz_localize(None)
    with pd.ExcelWriter(cache["path"], engine="openpyxl") as writer:
        matches.to_excel(writer, sheet_name="H2H_Maclar", index=False)
        queries.to_excel(writer, sheet_name="Sorgular", index=False)
        cache["team_ids"].to_excel(writer, sheet_name="Takim_ID", index=False)


def api_football_get(endpoint, params, api_key):
    response = requests.get(
        f"{API_FOOTBALL_BASE_URL}/{endpoint.lstrip('/')}",
        headers={"x-apisports-key": api_key},
        params=params,
        timeout=20,
    )
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        raise RuntimeError(
            f"API-Football HTTP {response.status_code}: {response.text[:300]}"
        ) from exc

    try:
        payload = response.json()
    except ValueError as exc:
        raise ValueError("API-Football yanıtı geçerli JSON değil.") from exc
    if not isinstance(payload, dict):
        raise ValueError("API-Football yanıt biçimi geçersiz.")
    errors = payload.get("errors")
    if errors:
        raise RuntimeError(f"API-Football hatası: {errors}")
    result = payload.get("response")
    if not isinstance(result, list):
        raise ValueError("API-Football yanıtında sonuç listesi bulunamadı.")
    return result


def get_api_team_id(team_name, league, api_key, cache):
    team_key = f"{league}|{normalize_team_name(team_name)}"
    existing = cache["team_ids"].loc[
        cache["team_ids"]["TeamKey"].astype(str).eq(team_key)
    ]
    if not existing.empty:
        return int(existing.iloc[-1]["TeamID"])

    params = {"search": team_name}
    league_id = API_LEAGUE_IDS.get(str(league))
    if league_id is not None:
        params.update({"league": league_id, "season": 2026})
    teams = api_football_get("teams", params, api_key)
    matches = [
        team
        for team in teams
        if normalize_team_name((team.get("team") or {}).get("name", ""))
        == normalize_team_name(team_name)
    ]
    if not matches and len(teams) == 1:
        matches = teams
    if len(matches) != 1:
        raise ValueError(
            f"{team_name} için tekil API-Football takım eşleşmesi bulunamadı."
        )
    team_id = int((matches[0].get("team") or {})["id"])
    cache["team_ids"] = pd.concat(
        [
            cache["team_ids"],
            pd.DataFrame(
                [
                    {
                        "TeamKey": team_key,
                        "Lig": league,
                        "Takim": team_name,
                        "TeamID": team_id,
                    }
                ]
            ),
        ],
        ignore_index=True,
    )
    return team_id


def api_h2h_matches(home_team, away_team, league, api_key, cache, pair_key):
    home_id = get_api_team_id(home_team, league, api_key, cache)
    away_id = get_api_team_id(away_team, league, api_key, cache)
    fixtures = api_football_get(
        "fixtures/headtohead",
        {"h2h": f"{home_id}-{away_id}", "last": 5},
        api_key,
    )
    rows = []
    for fixture in fixtures:
        fixture_info = fixture.get("fixture") or {}
        status = (fixture_info.get("status") or {}).get("short")
        goals = fixture.get("goals") or {}
        teams = fixture.get("teams") or {}
        home, away = teams.get("home") or {}, teams.get("away") or {}
        kickoff = pd.to_datetime(fixture_info.get("date"), errors="coerce", utc=True)
        if (
            status not in {"FT", "AET", "PEN"}
            or pd.isna(kickoff)
            or goals.get("home") is None
            or goals.get("away") is None
            or not home.get("name")
            or not away.get("name")
        ):
            continue
        rows.append(
            {
                "PairKey": pair_key,
                "TarihUTC": kickoff,
                "Lig": (fixture.get("league") or {}).get("name", ""),
                "EvSahibi": home["name"],
                "Deplasman": away["name"],
                "EvGol": int(goals["home"]),
                "DepGol": int(goals["away"]),
                "Kaynak": "API-Football",
            }
        )
    return rows


def cache_query_status(cache, pair_key, status, error=""):
    queries = cache["queries"].loc[
        cache["queries"]["PairKey"].astype(str).ne(pair_key)
    ].copy()
    query = pd.DataFrame(
        [
            {
                "PairKey": pair_key,
                "Durum": status,
                "SorguUTC": pd.Timestamp.now(tz="UTC"),
                "Hata": error,
            }
        ]
    )
    cache["queries"] = pd.concat([queries, query], ignore_index=True)


def get_head_to_head(finished, league, home_team, away_team, kickoff, cache, api_key):
    pair_key = pair_cache_key(home_team, away_team)
    local = finished.loc[
        finished["TarihUTC"].notna()
        & (finished["TarihUTC"] < kickoff)
        & (
            (
                finished["EvSahibi"].astype(str).map(normalize_team_name)
                == normalize_team_name(home_team)
            )
            & (
                finished["Deplasman"].astype(str).map(normalize_team_name)
                == normalize_team_name(away_team)
            )
            | (
                finished["EvSahibi"].astype(str).map(normalize_team_name)
                == normalize_team_name(away_team)
            )
            & (
                finished["Deplasman"].astype(str).map(normalize_team_name)
                == normalize_team_name(home_team)
            )
        )
    ]
    rows = [
        {
            "TarihUTC": match.TarihUTC,
            "Lig": match.Lig,
            "EvSahibi": match.EvSahibi,
            "Deplasman": match.Deplasman,
            "EvGol": int(match.EvGol),
            "DepGol": int(match.DepGol),
            "Kaynak": "Veri seti",
        }
        for match in local.itertuples(index=False)
    ]
    status_text = "API anahtarı tanımlı değil; yerel kayıtlar kullanılıyor."

    cached_queries = cache["queries"].loc[
        cache["queries"]["PairKey"].astype(str).eq(pair_key)
    ]
    if api_key and cached_queries.empty:
        cache_query_status(cache, pair_key, "BEKLIYOR")
        save_h2h_cache(cache)
        try:
            api_rows = api_h2h_matches(
                home_team, away_team, league, api_key, cache, pair_key
            )
            if api_rows:
                cached_rows = cache["matches"].loc[
                    cache["matches"]["PairKey"].astype(str).ne(pair_key)
                ]
                cache["matches"] = pd.concat(
                    [cached_rows, pd.DataFrame(api_rows, columns=H2H_MATCH_COLUMNS)],
                    ignore_index=True,
                )
                cache_query_status(cache, pair_key, "OK")
            else:
                cache_query_status(cache, pair_key, "BOS")
            status_text = "API-Football önbelleği güncel."
        except (requests.RequestException, RuntimeError, ValueError, KeyError) as exc:
            cache_query_status(cache, pair_key, "HATA", str(exc))
            status_text = f"API-Football sorgusu başarısız: {exc}"
        save_h2h_cache(cache)
    elif api_key and not cached_queries.empty:
        query = cached_queries.iloc[-1]
        status = str(query["Durum"])
        if status == "HATA":
            status_text = f"Önbellekteki API hatası: {query['Hata']}"
        elif status == "BOS":
            status_text = "API-Football önbelleğinde geçmiş karşılaşma bulunamadı."
        else:
            status_text = "API-Football önbelleği kullanılıyor."

    cached_matches = cache["matches"].loc[
        cache["matches"]["PairKey"].astype(str).eq(pair_key)
    ]
    for match in cached_matches.itertuples(index=False):
        match_date = pd.to_datetime(match.TarihUTC, errors="coerce", utc=True)
        if pd.isna(match_date) or match_date >= kickoff:
            continue
        rows.append(
            {
                "TarihUTC": match_date,
                "Lig": match.Lig,
                "EvSahibi": match.EvSahibi,
                "Deplasman": match.Deplasman,
                "EvGol": int(match.EvGol),
                "DepGol": int(match.DepGol),
                "Kaynak": match.Kaynak,
            }
        )

    deduplicated = {}
    for match in rows:
        match_date = pd.to_datetime(match["TarihUTC"], errors="coerce", utc=True)
        key = (
            match_date.strftime("%Y-%m-%d"),
            normalize_team_name(match["EvSahibi"]),
            normalize_team_name(match["Deplasman"]),
        )
        deduplicated[key] = {
            **match,
            "TarihUTC": match_date.strftime("%Y-%m-%d"),
        }
    result = sorted(
        deduplicated.values(), key=lambda item: item["TarihUTC"], reverse=True
    )[:5]
    return result, status_text


def poisson_probabilities(expected_goals, max_goals=6):
    probabilities = []
    for goals in range(max_goals + 1):
        if expected_goals == 0:
            probability = 1.0 if goals == 0 else 0.0
        else:
            probability = (
                math.exp(-expected_goals)
                * expected_goals**goals
                / math.factorial(goals)
            )
        probabilities.append(probability)
    return probabilities


def build_poisson_model(finished, league, home_team, away_team, kickoff):
    matches = finished.loc[
        (finished["Lig"] == league)
        & finished["TarihUTC"].notna()
        & (finished["TarihUTC"] < kickoff)
    ]
    if matches.empty:
        raise ValueError(f"{league} icin Poisson modeli kuracak bitmis mac yok.")

    league_home_average = matches["EvGol"].mean()
    league_away_average = matches["DepGol"].mean()
    if league_home_average <= 0 or league_away_average <= 0:
        raise ValueError(
            f"{league} icin ic saha veya deplasman gol ortalamasi sifir."
        )

    def team_strengths(team):
        home_matches = matches.loc[matches["EvSahibi"] == team]
        away_matches = matches.loc[matches["Deplasman"] == team]
        if home_matches.empty or away_matches.empty:
            raise ValueError(
                f"{league} / {team} icin ic saha ve deplasman gecmisi "
                "Poisson modeli kurmaya yetmiyor."
            )
        return {
            "home_attack": home_matches["EvGol"].mean() / league_home_average,
            "home_defense": home_matches["DepGol"].mean() / league_away_average,
            "away_attack": away_matches["DepGol"].mean() / league_away_average,
            "away_defense": away_matches["EvGol"].mean() / league_home_average,
        }

    home_strength = team_strengths(home_team)
    away_strength = team_strengths(away_team)
    home_lambda = (
        league_home_average
        * home_strength["home_attack"]
        * away_strength["away_defense"]
    )
    away_lambda = (
        league_away_average
        * away_strength["away_attack"]
        * home_strength["home_defense"]
    )

    home_goals = poisson_probabilities(home_lambda)
    away_goals = poisson_probabilities(away_lambda)
    score_matrix = [
        [home_probability * away_probability for away_probability in away_goals]
        for home_probability in home_goals
    ]
    matrix_total = sum(sum(row) for row in score_matrix)
    if matrix_total <= 0 or not math.isfinite(matrix_total):
        raise ValueError(
            f"{league} / {home_team} - {away_team} icin Poisson matrisi gecersiz."
        )
    score_matrix = [
        [probability / matrix_total for probability in row]
        for row in score_matrix
    ]

    home_win = sum(
        score_matrix[home_score][away_score]
        for home_score in range(7)
        for away_score in range(7)
        if home_score > away_score
    )
    draw = sum(score_matrix[score][score] for score in range(7))
    away_win = sum(
        score_matrix[home_score][away_score]
        for home_score in range(7)
        for away_score in range(7)
        if home_score < away_score
    )
    over_25 = sum(
        score_matrix[home_score][away_score]
        for home_score in range(7)
        for away_score in range(7)
        if home_score + away_score > 2.5
    )
    both_score = sum(
        score_matrix[home_score][away_score]
        for home_score in range(1, 7)
        for away_score in range(1, 7)
    )
    home_win_over_25 = sum(
        score_matrix[home_score][away_score]
        for home_score in range(7)
        for away_score in range(7)
        if home_score > away_score and home_score + away_score >= 3
    )
    away_win_over_25 = sum(
        score_matrix[home_score][away_score]
        for home_score in range(7)
        for away_score in range(7)
        if away_score > home_score and home_score + away_score >= 3
    )
    home_win_both_score = sum(
        score_matrix[home_score][away_score]
        for home_score in range(7)
        for away_score in range(7)
        if home_score > away_score and away_score >= 1
    )
    away_win_both_score = sum(
        score_matrix[home_score][away_score]
        for home_score in range(7)
        for away_score in range(7)
        if away_score > home_score and home_score >= 1
    )
    over_25_both_score = sum(
        score_matrix[home_score][away_score]
        for home_score in range(1, 7)
        for away_score in range(1, 7)
        if home_score + away_score >= 3
    )
    most_likely_home, most_likely_away = max(
        (
            (home_score, away_score)
            for home_score in range(7)
            for away_score in range(7)
        ),
        key=lambda score: score_matrix[score[0]][score[1]],
    )

    return {
        "Lig_IcSaha_Gol_Ort": league_home_average,
        "Lig_DisSaha_Gol_Ort": league_away_average,
        "Ev_IcSaha_Hucum_Gucu": home_strength["home_attack"],
        "Ev_IcSaha_Savunma_Gucu": home_strength["home_defense"],
        "Dep_DisSaha_Hucum_Gucu": away_strength["away_attack"],
        "Dep_DisSaha_Savunma_Gucu": away_strength["away_defense"],
        "Ev_xG_Lambda": home_lambda,
        "Dep_xG_Lambda": away_lambda,
        "MS1_%": home_win,
        "MSX_%": draw,
        "MS2_%": away_win,
        "Ust_25_%": over_25,
        "KG_Var_%": both_score,
        "MS1_ve_Ust25_%": home_win_over_25,
        "MS2_ve_Ust25_%": away_win_over_25,
        "MS1_ve_KG_Var_%": home_win_both_score,
        "MS2_ve_KG_Var_%": away_win_both_score,
        "Ust25_ve_KG_Var_%": over_25_both_score,
        "Ust_25_KG_Var_%": over_25_both_score,
        "En_Olasi_Skor": f"{most_likely_home}-{most_likely_away}",
    }


def build_upcoming_matches(data, finished, now, h2h_cache, api_key):
    week_end = now + timedelta(days=7)
    upcoming = data.loc[
        data["Durum"].isin(["SCHEDULED", "TIMED"])
        & data["TarihUTC"].notna()
        & (data["TarihUTC"] >= now)
        & (data["TarihUTC"] < week_end)
    ].sort_values("TarihUTC")
    history = build_team_history(finished)

    rows = []
    for match in upcoming.itertuples(index=False):
        home = get_team_stats(history, match.Lig, match.EvSahibi, match.TarihUTC)
        away = get_team_stats(history, match.Lig, match.Deplasman, match.TarihUTC)
        poisson = build_poisson_model(
            finished,
            match.Lig,
            match.EvSahibi,
            match.Deplasman,
            match.TarihUTC,
        )
        h2h_matches, h2h_api_status = get_head_to_head(
            finished,
            match.Lig,
            match.EvSahibi,
            match.Deplasman,
            match.TarihUTC,
            h2h_cache,
            api_key,
        )

        season_goal_sum = (
            home["season_scored"] + away["season_scored"]
            if home["season_scored"] is not None
            and away["season_scored"] is not None
            else None
        )
        recent_goal_sum = (
            home["recent_scored"] + away["recent_scored"]
            if home["recent_count"] == 5
            and away["recent_count"] == 5
            else None
        )

        labels = []
        if (
            season_goal_sum is not None and season_goal_sum >= 3.0
        ) or (recent_goal_sum is not None and recent_goal_sum >= 3.0):
            labels.append("2.5 Üst Potansiyeli")

        season_btts = all(
            value is not None and value >= threshold
            for value, threshold in (
                (home["season_scored"], 1.2),
                (home["season_conceded"], 1.1),
                (away["season_scored"], 1.2),
                (away["season_conceded"], 1.1),
            )
        )
        recent_btts = (
            home["recent_count"] == 5
            and away["recent_count"] == 5
            and home["recent_scored"] >= 1.2
            and home["recent_conceded"] >= 1.1
            and away["recent_scored"] >= 1.2
            and away["recent_conceded"] >= 1.1
        )
        if season_btts or recent_btts:
            labels.append("KG Var Adayı")

        form_advantage_team = None
        if (
            home["recent_count"] == 5
            and away["recent_count"] == 5
            and abs(home["recent_points"] - away["recent_points"]) >= 7
        ):
            labels.append("Güçlü Form Avantajı")
            form_advantage_team = (
                match.EvSahibi
                if home["recent_points"] > away["recent_points"]
                else match.Deplasman
            )

        rows.append(
            {
                "Tarih (UTC)": match.TarihUTC.strftime("%Y-%m-%d %H:%M"),
                "Lig": match.Lig,
                "Durum": match.Durum,
                "Ev Sahibi": match.EvSahibi,
                "Ev Form (G/B/M)": home["form"],
                "Ev Form Mac": home["recent_count"],
                "Ev Son 5 Puan": home["recent_points"],
                "Ev Attigi Gol Ort. (Son 5)": home["recent_scored"],
                "Ev Yedigi Gol Ort. (Son 5)": home["recent_conceded"],
                "Ev Attigi Gol Ort. (Sezon)": home["season_scored"],
                "Ev Yedigi Gol Ort. (Sezon)": home["season_conceded"],
                "Ev Oynanan Maç": home["season_count"],
                "Ev Sürpriz Yapma %": home["upset_made_rate"],
                "Ev Sürpriz Yeme %": home["upset_conceded_rate"],
                "Ev Sürpriz Yapma Sayısı": home["upset_made_count"],
                "Ev Sürpriz Yeme Sayısı": home["upset_conceded_count"],
                "Ev İlk Yarı Gol Ort.": home["first_half_average"],
                "Ev İlk Yarı Veri Sayısı": home["first_half_count"],
                "Ev İkinci Yarı Gol Ort.": home["second_half_average"],
                "Ev İkinci Yarı Veri Sayısı": home["second_half_count"],
                "Deplasman": match.Deplasman,
                "Dep Form (G/B/M)": away["form"],
                "Dep Form Mac": away["recent_count"],
                "Dep Son 5 Puan": away["recent_points"],
                "Dep Attigi Gol Ort. (Son 5)": away["recent_scored"],
                "Dep Yedigi Gol Ort. (Son 5)": away["recent_conceded"],
                "Dep Attigi Gol Ort. (Sezon)": away["season_scored"],
                "Dep Yedigi Gol Ort. (Sezon)": away["season_conceded"],
                "Dep Oynanan Maç": away["season_count"],
                "Dep Sürpriz Yapma %": away["upset_made_rate"],
                "Dep Sürpriz Yeme %": away["upset_conceded_rate"],
                "Dep Sürpriz Yapma Sayısı": away["upset_made_count"],
                "Dep Sürpriz Yeme Sayısı": away["upset_conceded_count"],
                "Dep İlk Yarı Gol Ort.": away["first_half_average"],
                "Dep İlk Yarı Veri Sayısı": away["first_half_count"],
                "Dep İkinci Yarı Gol Ort.": away["second_half_average"],
                "Dep İkinci Yarı Veri Sayısı": away["second_half_count"],
                "Sezon Takim Gol Ort. Toplami": season_goal_sum,
                "Son 5 Takim Gol Ort. Toplami": recent_goal_sum,
                "Form Avantajli Takim": form_advantage_team,
                "Istatistiksel_Etiketler": ", ".join(labels),
                **poisson,
                "H2H_Gecmis_JSON": json.dumps(
                    h2h_matches, ensure_ascii=False, default=str
                ),
                "Korner_Tahmini": "X",
                "Kart_Tahmini": "X",
                "Oyuncu_Tahmini": "X",
                "H2H_API_Durumu": h2h_api_status,
            }
        )
    return pd.DataFrame(rows, columns=UPCOMING_COLUMNS)


def save_report(report_path, sheets):
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes = "A2"
            worksheet.sheet_view.showGridLines = False

            if len(frame.columns):
                last_column = worksheet.cell(row=1, column=len(frame.columns)).column_letter
                if len(frame.index):
                    table = Table(
                        displayName=f"Table_{sheet_name}",
                        ref=f"A1:{last_column}{len(frame.index) + 1}",
                    )
                    table.tableStyleInfo = TableStyleInfo(
                        name="TableStyleMedium2",
                        showFirstColumn=False,
                        showLastColumn=False,
                        showRowStripes=True,
                        showColumnStripes=False,
                    )
                    worksheet.add_table(table)
                else:
                    worksheet.auto_filter.ref = f"A1:{last_column}1"

            for cell in worksheet[1]:
                cell.fill = PatternFill("solid", fgColor="17365D")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")

            for column_cells in worksheet.columns:
                width = min(
                    max(
                        len(str(cell.value)) if cell.value is not None else 0
                        for cell in column_cells
                    )
                    + 2,
                    42,
                )
                worksheet.column_dimensions[
                    column_cells[0].column_letter
                ].width = max(width, 12)

            for row in worksheet.iter_rows(min_row=2):
                for cell in row:
                    if isinstance(cell.value, float):
                        cell.number_format = "0.00"

            for column_name in PERCENTAGE_COLUMNS.intersection(frame.columns):
                column_index = frame.columns.get_loc(column_name) + 1
                for row_index in range(2, len(frame) + 2):
                    worksheet.cell(row=row_index, column=column_index).number_format = (
                        "0.00%"
                    )


def save_comparison_chart(summary, chart_path):
    if summary.empty:
        raise ValueError("Grafik icin lig ozeti verisi bulunamadi.")

    leagues = summary["Lig"].astype(str).tolist()
    goal_averages = summary["Mac Basi Gol Ort."].tolist()
    over_percentages = (summary["2.5 Ust %"] * 100).tolist()
    positions = list(range(len(leagues)))

    plt.rcParams.update(
        {
            "font.family": "DejaVu Sans",
            "font.size": 10,
            "axes.titleweight": "bold",
            "axes.edgecolor": "#D9E2F2",
            "axes.labelcolor": "#243247",
            "xtick.color": "#52627A",
            "ytick.color": "#52627A",
        }
    )
    figure, (goals_axis, over_axis) = plt.subplots(
        1, 2, figsize=(13, max(5.5, len(leagues) * 0.62)), sharey=True
    )
    figure.patch.set_facecolor("#F4F7FB")

    for axis in (goals_axis, over_axis):
        axis.set_facecolor("#FFFFFF")
        axis.spines[["top", "right", "left"]].set_visible(False)
        axis.spines["bottom"].set_color("#D9E2F2")
        axis.grid(axis="x", color="#E8EDF5", linewidth=0.8)
        axis.set_axisbelow(True)
        axis.tick_params(axis="y", length=0, pad=10)

    goal_bars = goals_axis.barh(
        positions,
        goal_averages,
        height=0.58,
        color="#3867D6",
        edgecolor="none",
        zorder=3,
    )
    goals_axis.set_yticks(positions, labels=leagues)
    goals_axis.invert_yaxis()
    goals_axis.set_xlim(0, max(goal_averages) * 1.24)
    goals_axis.set_xlabel("Gol / mac")
    goals_axis.set_title("Mac basi gol ortalamasi", loc="left", pad=14, color="#17365D")
    goals_axis.bar_label(goal_bars, fmt="%.2f", padding=5, color="#243247", fontsize=9)

    over_bars = over_axis.barh(
        positions,
        over_percentages,
        height=0.58,
        color="#18A999",
        edgecolor="none",
        zorder=3,
    )
    over_axis.set_xlim(0, max(over_percentages) * 1.24)
    over_axis.set_xlabel("Maclarin yuzdesi")
    over_axis.set_title("2.5 ust biten maclar", loc="left", pad=14, color="#17365D")
    over_axis.bar_label(
        over_bars,
        labels=[f"{value:.1f}%" for value in over_percentages],
        padding=5,
        color="#243247",
        fontsize=9,
    )

    figure.suptitle(
        "Liglere gore gol karsilastirmasi",
        x=0.08,
        ha="left",
        fontsize=17,
        color="#17365D",
        fontweight="bold",
    )
    figure.text(
        0.08,
        0.94,
        "Bitmis mac verileri | Gol ortalamasi ve 2.5 ust orani",
        color="#66758A",
        fontsize=10,
    )
    figure.subplots_adjust(left=0.09, right=0.97, top=0.86, bottom=0.12, wspace=0.28)
    figure.savefig(chart_path, dpi=180, facecolor=figure.get_facecolor())
    plt.close(figure)


def main():
    base_dir = Path(__file__).resolve().parent
    input_path = base_dir / INPUT_NAME
    if not input_path.exists():
        raise FileNotFoundError(f"{INPUT_NAME} dosyasi bulunamadi.")

    data = pd.read_excel(input_path)
    missing_columns = REQUIRED_COLUMNS.difference(data.columns)
    if missing_columns:
        raise KeyError(
            "Excel dosyasinda gerekli sutunlar eksik: "
            + ", ".join(sorted(missing_columns))
        )

    data = data.copy()
    data["Durum"] = data["Durum"].astype("string").str.upper().str.strip()
    data["EvGol"] = pd.to_numeric(data["EvGol"], errors="coerce")
    data["DepGol"] = pd.to_numeric(data["DepGol"], errors="coerce")
    data["TarihUTC"] = pd.to_datetime(data["Tarih"], utc=True, errors="coerce")
    halftime_home_column = find_column(data, HALF_TIME_HOME_COLUMNS)
    halftime_away_column = find_column(data, HALF_TIME_AWAY_COLUMNS)
    data["IlkYariEvGol"] = (
        pd.to_numeric(data[halftime_home_column], errors="coerce")
        if halftime_home_column is not None
        else float("nan")
    )
    data["IlkYariDepGol"] = (
        pd.to_numeric(data[halftime_away_column], errors="coerce")
        if halftime_away_column is not None
        else float("nan")
    )
    finished = data.loc[
        (data["Durum"] == "FINISHED")
        & data["EvGol"].notna()
        & data["DepGol"].notna()
    ].drop_duplicates(
        subset=["Lig", "TarihUTC", "EvSahibi", "Deplasman"],
        keep="last",
    ).copy()
    if halftime_home_column is None or halftime_away_column is None:
        print(
            "Uyarı: Kaynak çalışma kitabında devre skoru sütunları yok; "
            "ilk/ikinci yarı gol ortalamaları boş bırakılacak."
        )

    summary = build_league_summary(finished)
    standings = build_standings(finished)
    now = pd.Timestamp.now(tz="UTC")
    h2h_cache = load_h2h_cache(base_dir / H2H_CACHE_NAME)
    api_key = get_api_football_key()
    upcoming = build_upcoming_matches(data, finished, now, h2h_cache, api_key)
    if api_key:
        print(
            f"API-Football H2H önbelleği: {len(h2h_cache['queries'])} "
            "eşleşme sorgusu kayıtlı."
        )
    featured = upcoming.loc[
        upcoming["Istatistiksel_Etiketler"].fillna("").str.len() > 0
    ].copy()

    summary_output = summary.copy()
    for column in ("Ev Sahibi Kazanma %", "2.5 Ust %", "KG Var %"):
        summary_output[column] = summary_output[column].astype(float)
    report_path = base_dir / REPORT_NAME
    save_report(
        report_path,
        {
            "Lig_Ozeti": summary_output,
            "Puan_Durumu": standings,
            "Yaklasan_Maclar": upcoming,
            "One_Cikan_Maclar": featured,
        },
    )

    chart_path = base_dir / CHART_NAME
    save_comparison_chart(summary, chart_path)

    print(summary.to_string(index=False, formatters={
        "Mac Basi Gol Ort.": "{:.2f}".format,
        "Ev Sahibi Kazanma %": "{:.2%}".format,
        "2.5 Ust %": "{:.2%}".format,
        "KG Var %": "{:.2%}".format,
    }))
    print(
        f"\nPuan durumu: {len(standings)} takim/saha satiri; "
        f"yaklasan 7 gunde {len(upcoming)} mac; "
        f"one cikan etiketli mac: {len(featured)}."
    )
    print(f"Excel raporu kaydedildi: {report_path}")
    print(f"Grafik kaydedildi: {chart_path}")


if __name__ == "__main__":
    main()
