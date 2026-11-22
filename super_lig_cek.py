import re
import tempfile
import time
from datetime import datetime, timedelta, timezone
from html import unescape
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import pandas as pd


BASE_URL = "https://www.tff.org/default.aspx?pageID={page_id}&hafta={week}"
WORKBOOK_NAME = "api_tum_ligler.xlsx"
LEAGUES = [
    ("TSL", 198, 34, "Trendyol Süper Lig 2026-2027 Sezonu Fikstürü"),
    ("T1L", 142, 38, "Trendyol 1. Lig 2026-2027 Sezonu Fikstürü"),
]
REQUEST_DELAY_SECONDS = 0.5
COLUMNS = [
    "Lig",
    "Tarih",
    "Durum",
    "EvSahibi",
    "Deplasman",
    "EvGol",
    "DepGol",
    "Kazanan",
]
LOCAL_TIMEZONE = timezone(timedelta(hours=3))


def fetch_round(league_code, page_id, season_text, week):
    request = Request(
        BASE_URL.format(page_id=page_id, week=week),
        headers={
            "User-Agent": "Mozilla/5.0 (compatible; SuperLigFixtureExporter/1.0)",
            "Accept": "text/html",
        },
    )
    try:
        with urlopen(request, timeout=30) as response:
            page = response.read().decode(
                response.headers.get_content_charset() or "windows-1254",
                errors="replace",
            )
    except (HTTPError, URLError, TimeoutError) as exc:
        raise RuntimeError(
            f"TFF {league_code} {week}. hafta sayfasi alinamadi: {exc}"
        ) from exc

    if season_text not in unescape(page):
        raise RuntimeError(
            f"TFF {league_code} {week}. hafta sayfasinda "
            "beklenen 2026-2027 sezonu dogrulanamadi."
        )
    return page


def span_text(row_html, suffix):
    pattern = (
        r"<span\b[^>]*\bid=[\"'][^\"']*_"
        + re.escape(suffix)
        + r"[\"'][^>]*>(.*?)</span>"
    )
    match = re.search(pattern, row_html, flags=re.IGNORECASE | re.DOTALL)
    if not match:
        return ""
    text = re.sub(r"<[^>]*>", "", match.group(1))
    return unescape(text).strip()


def parse_fixture_date(date_text, time_text):
    try:
        local_date = datetime.strptime(date_text, "%d.%m.%Y")
    except ValueError as exc:
        raise ValueError(f"TFF fikstur tarihi taninamadi: {date_text!r}") from exc

    if not time_text:
        return local_date.strftime("%Y-%m-%d")

    try:
        local_datetime = datetime.strptime(
            f"{date_text} {time_text}", "%d.%m.%Y %H:%M"
        ).replace(tzinfo=LOCAL_TIMEZONE)
    except ValueError as exc:
        raise ValueError(
            f"TFF fikstur tarihi/saati taninamadi: {date_text} {time_text}"
        ) from exc

    return (
        local_datetime.astimezone(timezone.utc)
        .isoformat(timespec="seconds")
        .replace("+00:00", "Z")
    )


def parse_week(page, week, league_code):
    match_rows = re.findall(
        r"<tr\s+class=[\"']haftaninMaclariTr[\"'][^>]*>(.*?)</tr>",
        page,
        flags=re.IGNORECASE | re.DOTALL,
    )
    fixtures = []
    for row_html in match_rows:
        date_text = span_text(row_html, "lblTarih")
        home_team = span_text(row_html, "Label4")
        away_team = span_text(row_html, "Label1")
        if not date_text or not home_team or not away_team:
            raise ValueError(
                f"TFF {league_code} {week}. haftasinda "
                "takim veya tarih bilgisi eksik."
            )

        home_score_text = span_text(row_html, "Label5")
        away_score_text = span_text(row_html, "Label6")
        if home_score_text or away_score_text:
            if not home_score_text.isdigit() or not away_score_text.isdigit():
                raise ValueError(
                    f"TFF {league_code} {week}. haftasinda "
                    "taninamayan mac skoru: "
                    f"{home_score_text!r}-{away_score_text!r}"
                )
            home_score = int(home_score_text)
            away_score = int(away_score_text)
            status = "FINISHED"
            winner = (
                "HOME_TEAM"
                if home_score > away_score
                else "AWAY_TEAM"
                if away_score > home_score
                else "DRAW"
            )
        else:
            home_score = None
            away_score = None
            status = "SCHEDULED"
            winner = None

        fixtures.append(
            {
                "Lig": league_code,
                "Tarih": parse_fixture_date(
                    date_text, span_text(row_html, "lblSaat")
                ),
                "Durum": status,
                "EvSahibi": home_team,
                "Deplasman": away_team,
                "EvGol": home_score,
                "DepGol": away_score,
                "Kazanan": winner,
            }
        )

    if not fixtures:
        raise ValueError(
            f"TFF {league_code} {week}. haftasinda mac kaydi bulunamadi."
        )
    return fixtures


def download_season():
    fixtures = []
    request_count = sum(round_count for _, _, round_count, _ in LEAGUES)
    completed_requests = 0
    for league_code, page_id, round_count, season_text in LEAGUES:
        for week in range(1, round_count + 1):
            page = fetch_round(league_code, page_id, season_text, week)
            week_fixtures = parse_week(page, week, league_code)
            fixtures.extend(week_fixtures)
            print(
                f"{league_code} {week}. hafta: "
                f"{len(week_fixtures)} mac alindi."
            )
            completed_requests += 1
            if completed_requests < request_count:
                time.sleep(REQUEST_DELAY_SECONDS)

    keys = [
        (row["Lig"], row["Tarih"], row["EvSahibi"], row["Deplasman"])
        for row in fixtures
    ]
    if len(keys) != len(set(keys)):
        raise ValueError("TFF fiksturunde yinelenen mac kayitlari bulundu.")
    return fixtures


def update_workbook(fixtures, workbook_path):
    if not workbook_path.exists():
        raise FileNotFoundError(f"{WORKBOOK_NAME} dosyasi bulunamadi.")

    existing = pd.read_excel(workbook_path)
    if list(existing.columns) != COLUMNS:
        raise ValueError(
            "Mevcut Excel sutunlari beklenen sirayla uyusmuyor: "
            + ", ".join(map(str, existing.columns))
        )

    league_values = existing["Lig"].astype("string").str.strip().str.upper()
    league_codes = {league_code for league_code, _, _, _ in LEAGUES}
    old_matches = league_values.isin(league_codes)
    old_counts = {
        league_code: int((league_values == league_code).sum())
        for league_code in sorted(league_codes)
    }
    kept = existing.loc[~old_matches, COLUMNS]
    updated = pd.concat(
        [kept, pd.DataFrame(fixtures, columns=COLUMNS)],
        ignore_index=True,
    )

    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            suffix=".xlsx",
            prefix="api_tum_ligler_",
            dir=workbook_path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        updated.to_excel(temporary_path, index=False)
        temporary_path.replace(workbook_path)
    except Exception:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()
        raise

    return old_counts, len(updated)


def main():
    base_dir = Path(__file__).resolve().parent
    fixtures = download_season()
    workbook_path = base_dir / WORKBOOK_NAME
    old_counts, total_count = update_workbook(fixtures, workbook_path)
    print("\nTFF'den alinan ve guncellenen maclar:")
    for league_code, _, _, _ in LEAGUES:
        league_fixtures = [
            row for row in fixtures if row["Lig"] == league_code
        ]
        finished_count = sum(
            row["Durum"] == "FINISHED" for row in league_fixtures
        )
        scheduled_count = len(league_fixtures) - finished_count
        print(
            f"{league_code}: {len(league_fixtures)} mac "
            f"({finished_count} bitmis, {scheduled_count} gelecek); "
            f"eski satir sayisi: {old_counts[league_code]}."
        )
    print(f"{WORKBOOK_NAME} guncellendi; toplam satir: {total_count}.")


if __name__ == "__main__":
    main()
