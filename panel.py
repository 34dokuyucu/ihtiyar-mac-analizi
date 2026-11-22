import base64
from difflib import SequenceMatcher
from datetime import datetime, timedelta, timezone
from html import escape
from io import BytesIO
import json
import math
import os
from pathlib import Path
import re
import subprocess
import sys
import unicodedata

import pandas as pd
import streamlit as st
from PIL import Image
from streamlit.errors import StreamlitSecretNotFoundError


BASE_DIR = Path(__file__).resolve().parent
COVER_PATH = BASE_DIR / "kapak.png"
with Image.open(COVER_PATH) as cover_image:
    background_image = BytesIO()
    cover_image.convert("RGB").save(
        background_image, format="JPEG", quality=88, optimize=True
    )
COVER_BASE64 = base64.b64encode(background_image.getvalue()).decode("ascii")
COVER_DATA_URI = f"data:image/jpeg;base64,{COVER_BASE64}"
REPORT_PATH = BASE_DIR / "haftalik_analiz_raporu.xlsx"
MATCH_DATA_PATH = BASE_DIR / "api_tum_ligler.xlsx"
CHART_PATH = BASE_DIR / "lig_karsilastirma.png"
ODDS_PATH = BASE_DIR / "piyasa_oranlari.xlsx"
ODDS_COLUMNS = [
    "EvSahibi",
    "Deplasman",
    "Ort_MS1",
    "Ort_MSX",
    "Ort_MS2",
    "Ort_Ust25",
]
REQUIRED_SHEETS = {
    "Lig_Ozeti",
    "Puan_Durumu",
    "Yaklasan_Maclar",
    "One_Cikan_Maclar",
}
MATCH_DATA_COLUMNS = [
    "Lig",
    "Tarih",
    "Durum",
    "EvSahibi",
    "Deplasman",
    "EvGol",
    "DepGol",
]
JOINT_OVER_25_BOTH_SCORE_COLUMN = "Ust_25_KG_Var_%"
POISSON_PERCENT_COLUMNS = [
    "MS1_%",
    "MSX_%",
    "MS2_%",
    "Ust_25_%",
    "KG_Var_%",
    JOINT_OVER_25_BOTH_SCORE_COLUMN,
]
VIEW_OPTIONS = ["Otomatik", "Mobil Görünüm", "Masaüstü Görünümü"]
DATE_FILTER_OPTIONS = [
    "Tümü",
    "Bugün",
    "Önümüzdeki 3 Gün",
    "Özel Tarih",
]
OUTCOME_COLUMNS = {
    "Ev Sahibi Kazanır (MS1)": ("MS1_%", "Ort_MS1"),
    "Beraberlik (MSX)": ("MSX_%", "Ort_MSX"),
    "Deplasman Kazanır (MS2)": ("MS2_%", "Ort_MS2"),
    "2.5 Üst": ("Ust_25_%", "Ort_Ust25"),
    "Karşılıklı Gol (KG Var)": ("KG_Var_%", None),
}
LEAGUE_NAMES = {
    "TSL": "Türkiye Süper Ligi",
    "T1L": "Türkiye 1. Ligi",
    "PL": "İngiltere Premier Ligi",
    "ELC": "İngiltere Championship",
    "PD": "İspanya La Liga",
    "SA": "İtalya Serie A",
    "BL1": "Almanya Bundesliga",
    "FL1": "Fransa Ligue 1",
    "DED": "Hollanda Eredivisie",
    "PPL": "Portekiz Ligi",
    "CL": "UEFA Şampiyonlar Ligi",
    "J1": "Japonya J1 League",
    "K1": "Güney Kore K League 1",
    "UAE": "BAE Pro League",
    "EGY": "Mısır Premier League",
    "RSA": "Güney Afrika Premier Division",
    "PL2": "İngiltere Premier League 2 U21",
    "UYL": "UEFA Youth League U19",
}
FEATURED_COLUMN_LABELS = {
    "Tarih (UTC)": "Tarih / Saat (UTC)",
    "Istatistiksel_Etiketler": "Analiz Etiketleri",
    "Ev_xG_Lambda": "EV Gol Ort. (xG)",
    "Dep_xG_Lambda": "DEP Gol Ort. (xG)",
    "MS1_%": "MS1 %",
    "MSX_%": "MSX %",
    "MS2_%": "MS2 %",
    "Ust_25_%": "2.5 Üst %",
    "KG_Var_%": "KG Var %",
    JOINT_OVER_25_BOTH_SCORE_COLUMN: "2.5 Üst & KG Var %",
    "En_Olasi_Skor": "En Olası Skor",
}
LEAGUE_SUMMARY_COLUMN_LABELS = {
    "Oynanan Mac": "Oynanan Maç",
    "Mac Basi Gol Ort.": "Maç Başı Gol Ort.",
    "Ev Sahibi Kazanma %": "Ev Sahibi Galibiyeti %",
    "2.5 Ust %": "2.5 Üst %",
    "KG Var %": "KG Var %",
}
STANDINGS_COLUMN_LABELS = {
    "Saha": "Saha",
    "Sira": "Sıra",
    "Takim": "Takım",
    "O": "Oynanan",
    "G": "Galibiyet",
    "B": "Beraberlik",
    "M": "Mağlubiyet",
    "AG": "Atılan Gol",
    "YG": "Yenilen Gol",
    "AV": "Gol Farkı",
    "Puan": "Puan",
}
UPDATE_STEPS = [
    ("api_veri_cek.py", "Lig verilerini çekme"),
    ("super_lig_cek.py", "TFF lig verilerini güncelleme"),
    ("oran_cek.py", "Piyasa oranlarını çekme"),
    ("detayli_analiz.py", "Analiz raporunu ve grafiği oluşturma"),
]
API_SECRET_NAMES = ["FOOTBALL_DATA_API_KEY", "ODDS_API_KEY"]


@st.cache_data(ttl=60, show_spinner="Rapor yükleniyor...")
def load_report():
    if not REPORT_PATH.is_file():
        raise FileNotFoundError(f"Rapor dosyası bulunamadı: {REPORT_PATH}")

    workbook = pd.ExcelFile(REPORT_PATH)
    missing = REQUIRED_SHEETS.difference(workbook.sheet_names)
    if missing:
        raise ValueError(
            "Rapor dosyasında gerekli sekmeler eksik: "
            + ", ".join(sorted(missing))
        )

    return {
        sheet_name: pd.read_excel(REPORT_PATH, sheet_name=sheet_name)
        for sheet_name in REQUIRED_SHEETS
    }


@st.cache_data(ttl=60, show_spinner="Piyasa oranları yükleniyor...")
def load_odds():
    if not ODDS_PATH.is_file():
        return pd.DataFrame(columns=ODDS_COLUMNS)

    odds = pd.read_excel(ODDS_PATH)
    missing = set(ODDS_COLUMNS).difference(odds.columns)
    if missing:
        raise ValueError(
            "Piyasa oranı dosyasında sütunlar eksik: "
            + ", ".join(sorted(missing))
        )
    return odds[ODDS_COLUMNS].copy()


@st.cache_data(ttl=60, show_spinner="Maç verileri yükleniyor...")
def load_match_data():
    if not MATCH_DATA_PATH.is_file():
        raise FileNotFoundError(f"Maç verisi bulunamadı: {MATCH_DATA_PATH}")

    matches = pd.read_excel(MATCH_DATA_PATH)
    missing = set(MATCH_DATA_COLUMNS).difference(matches.columns)
    if missing:
        raise ValueError(
            "Maç verisi dosyasında gerekli sütunlar eksik: "
            + ", ".join(sorted(missing))
        )
    return matches[MATCH_DATA_COLUMNS].copy()


def build_team_statistics(matches):
    columns = [
        "Lig",
        "Takim",
        "Oynanan",
        "Ust25Orani",
        "GolYemedenBitirmeOrani",
        "GolOrtalamasi",
    ]
    if matches.empty:
        return pd.DataFrame(columns=columns)

    finished = matches.loc[
        matches["Durum"].astype(str).str.upper().eq("FINISHED")
    ].copy()
    finished["EvGol"] = pd.to_numeric(finished["EvGol"], errors="coerce")
    finished["DepGol"] = pd.to_numeric(finished["DepGol"], errors="coerce")
    finished = finished.dropna(
        subset=["Lig", "Tarih", "EvSahibi", "Deplasman", "EvGol", "DepGol"]
    ).drop_duplicates(
        subset=["Lig", "Tarih", "EvSahibi", "Deplasman"],
        keep="last",
    )
    if finished.empty:
        return pd.DataFrame(columns=columns)

    home = pd.DataFrame(
        {
            "Lig": finished["Lig"],
            "Takim": finished["EvSahibi"],
            "GolAtilan": finished["EvGol"],
            "GolYenilen": finished["DepGol"],
        }
    )
    away = pd.DataFrame(
        {
            "Lig": finished["Lig"],
            "Takim": finished["Deplasman"],
            "GolAtilan": finished["DepGol"],
            "GolYenilen": finished["EvGol"],
        }
    )
    team_matches = pd.concat([home, away], ignore_index=True)
    team_matches["Ust25"] = (
        team_matches["GolAtilan"] + team_matches["GolYenilen"] > 2.5
    )
    team_matches["GolYemedenBitirme"] = team_matches["GolYenilen"].eq(0)

    statistics = (
        team_matches.groupby(["Lig", "Takim"], as_index=False)
        .agg(
            Oynanan=("GolAtilan", "size"),
            Ust25Orani=("Ust25", "mean"),
            GolYemedenBitirmeOrani=("GolYemedenBitirme", "mean"),
            GolOrtalamasi=("GolAtilan", "mean"),
        )
    )
    return statistics.loc[statistics["Oynanan"] >= 5].copy()


def percent_display(frame, columns):
    display_frame = frame.copy()
    for column in columns:
        if column in display_frame.columns:
            display_frame[column] = pd.to_numeric(
                display_frame[column], errors="coerce"
            ) * 100
    return display_frame


def probability_level(value, fraction=False):
    try:
        numeric_value = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(numeric_value):
        return None

    percentage = numeric_value * 100 if fraction else numeric_value
    if percentage >= 65:
        return "high"
    if percentage >= 45:
        return "medium"
    return "low"


def probability_cell_style(value):
    styles = {
        "high": "color: #10B981; font-weight: 800;",
        "medium": "color: #F59E0B;",
        "low": "color: #EF4444;",
    }
    return styles.get(probability_level(value), "")


def color_probability_cells(frame, columns):
    return frame.style.map(probability_cell_style, subset=columns)


def probability_column_config(columns):
    return {
        column: st.column_config.NumberColumn(
            column,
            format="%.1f%%",
            help="Poisson modelinden hesaplanan olasılık.",
        )
        for column in columns
    }


def match_label(row):
    return (
        f"{row['Tarih (UTC)']} | {league_label(row['Lig'])} | "
        f"{row['Ev Sahibi']} - {row['Deplasman']}"
    )


def filter_matches_by_date(frame, start_date=None, end_date=None):
    if start_date is None or end_date is None:
        return frame.copy()

    match_dates = pd.to_datetime(
        frame["Tarih (UTC)"], errors="coerce", utc=True
    ).dt.date
    mask = match_dates.ge(start_date) & match_dates.le(end_date)
    return frame.loc[mask].copy()


def poisson_probabilities(expected_goals, max_goals=6):
    expected_goals = float(expected_goals)
    if not math.isfinite(expected_goals) or expected_goals < 0:
        raise ValueError("Poisson beklenen gol değeri geçerli değil.")

    return [
        math.exp(-expected_goals)
        * expected_goals**goals
        / math.factorial(goals)
        for goals in range(max_goals + 1)
    ]


def over_25_and_both_score_probability(home_lambda, away_lambda):
    home_goals = poisson_probabilities(home_lambda)
    away_goals = poisson_probabilities(away_lambda)
    score_matrix = [
        [home_probability * away_probability for away_probability in away_goals]
        for home_probability in home_goals
    ]
    matrix_total = sum(sum(row) for row in score_matrix)
    if not math.isfinite(matrix_total) or matrix_total <= 0:
        raise ValueError("Poisson gol matrisi geçerli değil.")

    return sum(
        score_matrix[home_score][away_score]
        for home_score in range(1, 7)
        for away_score in range(1, 7)
        if home_score + away_score >= 3
    ) / matrix_total


def add_joint_probability_column(matches):
    result = matches.copy()
    result[JOINT_OVER_25_BOTH_SCORE_COLUMN] = [
        over_25_and_both_score_probability(home_lambda, away_lambda)
        for home_lambda, away_lambda in zip(
            result["Ev_xG_Lambda"], result["Dep_xG_Lambda"]
        )
    ]
    return result


def format_match_date(value):
    parsed = pd.to_datetime(value, errors="coerce", utc=True)
    if pd.isna(parsed):
        return "Tarih bilgisi yok"

    date_label = parsed.strftime("%d.%m.%Y")
    if parsed.hour == 0 and parsed.minute == 0 and parsed.second == 0:
        return f"{date_label} · Saat açıklanmadı"
    return f"{date_label} · {parsed.strftime('%H:%M')} UTC"


def format_probability(value):
    try:
        probability = float(value)
    except (TypeError, ValueError):
        return "—"
    if not math.isfinite(probability):
        return "—"
    return f"{probability * 100:.1f}%"


def format_decimal(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "—"
    if not math.isfinite(number):
        return "—"
    return f"{number:.2f}"


def format_detail_average(value, sample_count):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "Veri yok"
    if not math.isfinite(number) or not sample_count:
        return "Veri yok"
    return f"{number:.2f} (n={int(sample_count)})"


def format_upset_rate(rate, upset_count, played_count):
    try:
        percentage = float(rate) * 100
    except (TypeError, ValueError):
        return "Veri yok"
    if not math.isfinite(percentage) or not played_count:
        return "Veri yok"
    return f"{percentage:.1f}% ({int(upset_count or 0)}/{int(played_count)})"


def render_match_details(match):
    with st.expander(
        "Detaylı Analiz (Form, Sürpriz, Yarı Golleri & Eski Maçlar H2H)",
        expanded=False,
    ):
        team_columns = st.columns(2)
        teams = [
            (
                "Ev Sahibi",
                "Ev Sahibi",
                "Ev Form (G/B/M)",
                "Ev Oynanan Maç",
                "Ev Sürpriz Yapma %",
                "Ev Sürpriz Yapma Sayısı",
                "Ev Sürpriz Yeme %",
                "Ev Sürpriz Yeme Sayısı",
                "Ev İlk Yarı Gol Ort.",
                "Ev İlk Yarı Veri Sayısı",
                "Ev İkinci Yarı Gol Ort.",
                "Ev İkinci Yarı Veri Sayısı",
            ),
            (
                "Deplasman",
                "Deplasman",
                "Dep Form (G/B/M)",
                "Dep Oynanan Maç",
                "Dep Sürpriz Yapma %",
                "Dep Sürpriz Yapma Sayısı",
                "Dep Sürpriz Yeme %",
                "Dep Sürpriz Yeme Sayısı",
                "Dep İlk Yarı Gol Ort.",
                "Dep İlk Yarı Veri Sayısı",
                "Dep İkinci Yarı Gol Ort.",
                "Dep İkinci Yarı Veri Sayısı",
            ),
        ]
        for column, fields in zip(team_columns, teams):
            (
                heading,
                team_column,
                form_column,
                played_column,
                upset_made_column,
                upset_made_count_column,
                upset_conceded_column,
                upset_conceded_count_column,
                first_half_column,
                first_half_count_column,
                second_half_column,
                second_half_count_column,
            ) = fields
            with column:
                st.markdown(f"**{escape(str(match.get(team_column, heading)))}**")
                st.metric("Son 5 maç formu", str(match.get(form_column) or "Yok"))
                played = match.get(played_column, 0)
                st.metric(
                    "Sürpriz yapma",
                    format_upset_rate(
                        match.get(upset_made_column),
                        match.get(upset_made_count_column),
                        played,
                    ),
                )
                st.metric(
                    "Sürpriz yeme",
                    format_upset_rate(
                        match.get(upset_conceded_column),
                        match.get(upset_conceded_count_column),
                        played,
                    ),
                )
                st.metric(
                    "İlk yarı gol yoğunluğu",
                    format_detail_average(
                        match.get(first_half_column),
                        match.get(first_half_count_column),
                    ),
                )
                st.metric(
                    "İkinci yarı gol yoğunluğu",
                    format_detail_average(
                        match.get(second_half_column),
                        match.get(second_half_count_column),
                    ),
                )

        st.caption(
            "Sürpriz oranı, maç öncesi lig puanı/maç oranı daha düşük olan "
            "takımın galibiyetinin tüm bitmiş maçlara oranıdır."
        )
        h2h_value = match.get("H2H_Gecmis_JSON")
        try:
            h2h_matches = json.loads(str(h2h_value)) if pd.notna(h2h_value) else []
        except (TypeError, json.JSONDecodeError):
            h2h_matches = []
        st.markdown("**Eski Maçlar H2H**")
        if h2h_matches:
            rows = []
            for h2h_match in h2h_matches:
                rows.append(
                    '<article class="h2h-row">'
                    f'<span>{escape(str(h2h_match.get("TarihUTC", "Tarih yok")))} · '
                    f'{escape(league_label(h2h_match.get("Lig", "")))}</span>'
                    '<strong>'
                    f'{escape(str(h2h_match.get("EvSahibi", "—")))} '
                    f'{escape(str(h2h_match.get("EvGol", "—")))}–'
                    f'{escape(str(h2h_match.get("DepGol", "—")))} '
                    f'{escape(str(h2h_match.get("Deplasman", "—")))}'
                    '</strong>'
                    f'<small>{escape(str(h2h_match.get("Kaynak", "")))}</small>'
                    "</article>"
                )
            st.markdown(
                '<div class="h2h-list">' + "".join(rows) + "</div>",
                unsafe_allow_html=True,
            )
        else:
            st.info("Bu iki takımın maç öncesine ait geçmiş karşılaşması bulunamadı.")
        api_status = match.get("H2H_API_Durumu")
        if pd.notna(api_status) and str(api_status).strip():
            st.caption(str(api_status))
        st.caption(
            "İlk/ikinci yarı ortalamaları yalnızca kaynak dosyada devre skoru "
            "bulunan bitmiş maçlardan hesaplanır."
        )


def render_match_cards(matches):
    rows = list(matches.iterrows())
    for start in range(0, len(rows), 3):
        batch = rows[start : start + 3]
        columns = st.columns(3)
        for column, (_, match) in zip(columns, batch):
            label_value = match.get("Istatistiksel_Etiketler")
            labels = (
                [
                    label.strip()
                    for label in str(label_value).split(",")
                    if label.strip()
                ]
                if pd.notna(label_value)
                else []
            )
            label_markup = "".join(
                f'<span class="match-tag">{escape(label)}</span>'
                for label in labels
            ) or '<span class="match-tag match-tag-muted">Etiket yok</span>'

            probabilities = [
                ("MS1", match.get("MS1_%")),
                ("MSX", match.get("MSX_%")),
                ("MS2", match.get("MS2_%")),
                ("2.5 Üst", match.get("Ust_25_%")),
                ("KG Var", match.get("KG_Var_%")),
                (
                    "2.5 Üst &amp; KG Var",
                    match.get(JOINT_OVER_25_BOTH_SCORE_COLUMN),
                ),
            ]
            probability_markup = "".join(
                f'<div class="probability-box prob-{probability_level(value, fraction=True) or "na"}">'
                f"<span>{label}</span>"
                f"<strong>{format_probability(value)}</strong>"
                "</div>"
                for label, value in probabilities
            )
            home_xg = format_decimal(match.get("Ev_xG_Lambda"))
            away_xg = format_decimal(match.get("Dep_xG_Lambda"))
            card = (
                '<article class="match-card">'
                '<div class="match-card-top">'
                f'<span class="match-league">{escape(league_label(match.get("Lig", "")))}</span>'
                f'<span class="match-date">{escape(format_match_date(match.get("Tarih (UTC)")))}</span>'
                "</div>"
                '<div class="match-teams">'
                f'<span>{escape(str(match.get("Ev Sahibi", "—")))}</span>'
                '<span class="match-vs">VS</span>'
                f'<span>{escape(str(match.get("Deplasman", "—")))}</span>'
                "</div>"
                f'<div class="match-tags">{label_markup}</div>'
                '<div class="match-xg">'
                f"<span>EV Gol Ort. (xG) <strong>{home_xg}</strong></span>"
                f"<span>DEP Gol Ort. (xG) <strong>{away_xg}</strong></span>"
                f"<span>En olası skor <strong>{escape(str(match.get('En_Olasi_Skor', '—')))}</strong></span>"
                "</div>"
                f'<div class="probability-grid">{probability_markup}</div>'
                "</article>"
            )
            with column:
                st.markdown(card, unsafe_allow_html=True)
                render_match_details(match)


def render_market_difference_cards(difference_table):
    cards = []
    for _, row in difference_table.iterrows():
        model_level = probability_level(row["Model Olasılığı (%)"]) or "na"
        market_level = (
            probability_level(row["Örtük Piyasa Olasılığı (%)"]) or "na"
        )
        cards.append(
            '<article class="match-card market-difference-card">'
            f'<div class="market-match-label">{escape(str(row["Maç"]))}</div>'
            f'<div class="match-tags"><span class="match-tag">{escape(str(row["Piyasa"]))}</span></div>'
            '<div class="probability-grid market-probability-grid">'
            f'<div class="probability-box prob-{model_level}"><span>Model Olasılığı</span><strong>{float(row["Model Olasılığı (%)"]):.1f}%</strong></div>'
            f'<div class="probability-box prob-{market_level}"><span>Piyasa Olasılığı</span><strong>{float(row["Örtük Piyasa Olasılığı (%)"]):.1f}%</strong></div>'
            f'<div class="probability-box"><span>Fark</span><strong>{float(row["Fark (yüzde puan)"]):+.1f} puan</strong></div>'
            f'<div class="probability-box"><span>Ort. Oran</span><strong>{float(row["Ortalama Oran"]):.2f}</strong></div>'
            "</div>"
            "</article>"
        )
    st.markdown(
        '<div class="match-card-list">' + "".join(cards) + "</div>",
        unsafe_allow_html=True,
    )


def configured_secret(name):
    value = os.environ.get(name, "").strip()
    if value:
        return value

    try:
        return str(st.secrets.get(name, "") or "").strip()
    except StreamlitSecretNotFoundError:
        return ""


def run_update_pipeline():
    environment = os.environ.copy()
    for secret_name in API_SECRET_NAMES:
        secret_value = configured_secret(secret_name)
        if secret_value:
            environment[secret_name] = secret_value
        else:
            environment.pop(secret_name, None)

    with st.status("Veriler güncelleniyor...", expanded=True) as status:
        for index, (script_name, description) in enumerate(UPDATE_STEPS, start=1):
            status.update(
                label=f"[{index}/{len(UPDATE_STEPS)}] {description}...",
                state="running",
            )
            st.write(f"[{index}/{len(UPDATE_STEPS)}] Başladı: {description}")
            script_path = BASE_DIR / script_name
            if not script_path.is_file():
                message = f"Güncelleme betiği bulunamadı: {script_path}"
                st.error(message)
                status.update(label="Veri güncellemesi başarısız.", state="error")
                return False

            try:
                result = subprocess.run(
                    [sys.executable, str(script_path)],
                    cwd=BASE_DIR,
                    env=environment,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    check=False,
                )
            except OSError as exc:
                st.error(f"{script_name} başlatılamadı: {exc}")
                status.update(label="Veri güncellemesi başarısız.", state="error")
                return False

            if result.stdout.strip():
                st.code(result.stdout[-12000:], language="text")
            if result.returncode != 0:
                st.error(
                    f"{script_name} başarısız oldu (çıkış kodu "
                    f"{result.returncode})."
                )
                if result.stderr.strip():
                    st.code(result.stderr[-12000:], language="text")
                status.update(label="Veri güncellemesi başarısız.", state="error")
                return False

            st.write(f"[{index}/{len(UPDATE_STEPS)}] Tamamlandı.")

        status.update(label="Tüm veriler başarıyla güncellendi.", state="complete")
    return True


def league_label(code):
    return LEAGUE_NAMES.get(str(code), str(code))


def make_leaderboard_entries(frame, value_column, percentage=True):
    entries = []
    for _, row in frame.head(5).iterrows():
        if "Takim" in frame.columns:
            name = str(row["Takim"])
            detail = (
                f"{league_label(row['Lig'])} · {int(row['Oynanan']):,} maç"
            )
        else:
            name = league_label(row["Lig"])
            detail = f"{int(row['Oynanan Mac']):,} maç"

        value = float(row[value_column])
        badge = f"{value * 100:.1f}%" if percentage else f"{value:.2f} gol/maç"
        entries.append({"name": name, "detail": detail, "value": badge})
    return entries


def render_leaderboard_card(title, entries):
    if entries:
        rows = "".join(
            '<div class="leaderboard-row">'
            f'<span class="leaderboard-rank">{index:02d}</span>'
            '<div class="leaderboard-entity">'
            f'<strong>{escape(entry["name"])}</strong>'
            f'<small>{escape(entry["detail"])}</small>'
            "</div>"
            f'<span class="leaderboard-badge">{escape(entry["value"])}</span>'
            "</div>"
            for index, entry in enumerate(entries, start=1)
        )
    else:
        rows = (
            '<div class="leaderboard-empty">'
            "En az 5 maçlık yeterli veri bulunmuyor."
            "</div>"
        )

    st.markdown(
        '<section class="leaderboard-card">'
        '<div class="leaderboard-card-heading">'
        f"<h3>{escape(title)}</h3>"
        '<span class="leaderboard-top-label">İLK 5</span>'
        "</div>"
        f'<div class="leaderboard-list">{rows}</div>'
        "</section>",
        unsafe_allow_html=True,
    )


def render_leaderboards(league_summary, team_statistics, match_data_error=None):
    if match_data_error:
        st.warning(f"Takım liderleri için maç verisi kullanılamıyor: {match_data_error}")

    eligible_leagues = league_summary.copy()
    eligible_leagues["Oynanan Mac"] = pd.to_numeric(
        eligible_leagues["Oynanan Mac"], errors="coerce"
    )
    eligible_leagues = eligible_leagues.loc[
        eligible_leagues["Oynanan Mac"] >= 5
    ].copy()

    league_btts = eligible_leagues.dropna(subset=["KG Var %"]).sort_values(
        ["KG Var %", "Oynanan Mac"], ascending=[False, False]
    )
    home_advantage = eligible_leagues.dropna(
        subset=["Ev Sahibi Kazanma %"]
    ).sort_values(
        ["Ev Sahibi Kazanma %", "Oynanan Mac"], ascending=[False, False]
    )
    team_over25 = team_statistics.sort_values(
        ["Ust25Orani", "Oynanan"], ascending=[False, False]
    )
    team_clean_sheets = team_statistics.sort_values(
        ["GolYemedenBitirmeOrani", "Oynanan"], ascending=[False, False]
    )
    team_scoring = team_statistics.sort_values(
        ["GolOrtalamasi", "Oynanan"], ascending=[False, False]
    )

    cards = [
        (
            "En İyi KG Var Ligleri (%)",
            make_leaderboard_entries(league_btts, "KG Var %"),
        ),
        (
            "Takımlara Göre En Çok 2.5 Üst (%)",
            make_leaderboard_entries(team_over25, "Ust25Orani"),
        ),
        (
            "En Çok Maçta Kalesini Gole Kapatan Takımlar (%)",
            make_leaderboard_entries(
                team_clean_sheets, "GolYemedenBitirmeOrani"
            ),
        ),
        (
            "Maç Başı En Golcü Takımlar",
            make_leaderboard_entries(
                team_scoring, "GolOrtalamasi", percentage=False
            ),
        ),
        (
            "En Yüksek Ev Sahibi Avantajı Olan Ligler (%)",
            make_leaderboard_entries(home_advantage, "Ev Sahibi Kazanma %"),
        ),
    ]

    st.caption(
        "Lig ve takım sıralamaları bitmiş maçlara dayanır; takım listelerinde "
        "en az 5 maç şartı uygulanır."
    )
    for index in range(0, len(cards), 2):
        columns = st.columns(2, gap="medium")
        for column, (title, entries) in zip(columns, cards[index : index + 2]):
            with column:
                render_leaderboard_card(title, entries)


def normalized_team_name(value):
    text = str(value).casefold().translate(
        str.maketrans(
            {
                "ı": "i",
                "ş": "s",
                "ğ": "g",
                "ç": "c",
                "ö": "o",
                "ü": "u",
            }
        )
    )
    text = unicodedata.normalize("NFKD", text)
    text = "".join(character for character in text if not unicodedata.combining(character))
    tokens = re.findall(r"[a-z0-9]+", text)
    ignored = {
        "a",
        "ac",
        "as",
        "afc",
        "cf",
        "club",
        "futbol",
        "fc",
        "fk",
        "kulubu",
        "sc",
        "sk",
        "sport",
        "sportif",
        "sports",
        "spor",
        "the",
    }
    aliases = {"internazionale": "inter", "milano": "milan", "munchen": "munich"}
    normalized = []
    for token in tokens:
        if token in ignored:
            continue
        if token.endswith("spor") and len(token) > len("spor") + 2:
            token = token[: -len("spor")]
        token = aliases.get(token, token)
        if token:
            normalized.append(token)
    return " ".join(normalized)


def team_similarity(first, second):
    first_name = normalized_team_name(first)
    second_name = normalized_team_name(second)
    if not first_name or not second_name:
        return 0.0
    if first_name == second_name:
        return 1.0

    first_tokens = set(first_name.split())
    second_tokens = set(second_name.split())
    common = first_tokens & second_tokens
    containment = len(common) / min(len(first_tokens), len(second_tokens))
    jaccard = len(common) / len(first_tokens | second_tokens)
    sequence = max(
        SequenceMatcher(None, first_name, second_name, autojunk=False).ratio(),
        SequenceMatcher(
            None,
            " ".join(sorted(first_tokens)),
            " ".join(sorted(second_tokens)),
            autojunk=False,
        ).ratio(),
    )
    return max(sequence, containment * 0.92, jaccard)


def find_market_odds(match, odds):
    best_match = None
    best_score = 0.0
    for _, candidate in odds.iterrows():
        home_score = team_similarity(match["Ev Sahibi"], candidate["EvSahibi"])
        away_score = team_similarity(match["Deplasman"], candidate["Deplasman"])
        score = (home_score + away_score) / 2
        if (
            home_score >= 0.70
            and away_score >= 0.70
            and score >= 0.78
            and score > best_score
        ):
            best_match = candidate
            best_score = score
    return best_match, best_score


def build_market_difference_table(upcoming, odds):
    rows = []
    matched_events = 0
    for _, match in upcoming.iterrows():
        market_match, similarity = find_market_odds(match, odds)
        if market_match is None:
            continue
        matched_events += 1

        for outcome, (model_column, odds_column) in OUTCOME_COLUMNS.items():
            if odds_column is None:
                continue
            market_price = pd.to_numeric(
                pd.Series([market_match[odds_column]]), errors="coerce"
            ).iloc[0]
            if pd.isna(market_price) or market_price <= 0:
                continue

            model_probability = float(match[model_column]) * 100
            implied_probability = 100 / float(market_price)
            difference = model_probability - implied_probability
            if difference >= 5:
                rows.append(
                    {
                        "Maç": match_label(match),
                        "Piyasa": outcome,
                        "Ortalama Oran": float(market_price),
                        "Model Olasılığı (%)": model_probability,
                        "Örtük Piyasa Olasılığı (%)": implied_probability,
                        "Fark (yüzde puan)": difference,
                        "Takım Adı Eşleşmesi (%)": similarity * 100,
                    }
                )
    table = pd.DataFrame(rows)
    if not table.empty:
        table = table.sort_values(
            "Fark (yüzde puan)", ascending=False
        ).reset_index(drop=True)
    return table, matched_events


st.set_page_config(
    page_title="İ H T İ Y A R Maç Analiz Paneli",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="auto",
)

if "giris_ok" not in st.session_state:
    giris_parolasi = st.text_input("Giriş Parolası", type="password")
    if st.button("Giriş Yap"):
        if giris_parolasi.strip().lower() == "ihtiyarpeek":
            st.session_state["giris_ok"] = True
            st.rerun()
    st.stop()

st.iframe(
    """
    <script>
    const head = window.parent.document.head;
    const metadata = [
        ["viewport", "width=device-width, initial-scale=1, viewport-fit=cover"],
        ["apple-mobile-web-app-capable", "yes"],
        ["apple-mobile-web-app-status-bar-style", "default"],
        ["apple-mobile-web-app-title", "İ H T İ Y A R Maç Analiz Paneli"],
        ["mobile-web-app-capable", "yes"],
        ["theme-color", "#f5f7fb"]
    ];
    for (const [name, content] of metadata) {
        let tag = head.querySelector(`meta[name="${name}"]`);
        if (!tag) {
            tag = head.ownerDocument.createElement("meta");
            tag.name = name;
            head.appendChild(tag);
        }
        tag.content = content;
    }
    </script>
    """,
    width=1,
    height=1,
)

st.markdown(
    """
    <style>
    :root {
        --panel-bg: #0B0F19;
        --panel-bg-secondary: #131B2E;
        --panel-card: #1A2236;
        --panel-border: #2A3650;
        --panel-text: #E5EAF3;
        --panel-muted: #9AA8BF;
        --panel-green: #10B981;
        --panel-cyan: #38BDF8;
        --panel-cover-image: url("__COVER_DATA_URI__");
    }
    .stApp, [data-testid="stAppViewContainer"] {
        color: var(--panel-text);
    }
    .stApp {
        position: relative;
        isolation: isolate;
        background: linear-gradient(135deg, var(--panel-bg), var(--panel-bg-secondary));
    }
    [data-testid="stAppViewContainer"] {
        background: transparent;
    }
    .stApp::before {
        position: fixed;
        z-index: 0;
        inset: 0;
        background:
            linear-gradient(rgba(10, 14, 23, 0.68), rgba(10, 14, 23, 0.78)),
            var(--panel-cover-image);
        background-size: cover;
        background-position: center center;
        background-repeat: no-repeat;
        background-attachment: fixed;
        content: "";
        pointer-events: none;
    }
    .stApp > header, .stApp > div[data-testid="stAppViewContainer"] {
        position: relative;
        z-index: 1;
    }
    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stSidebar"] {
        background: rgba(11, 15, 25, 0.88);
        border-right: 1px solid rgba(16, 185, 129, 0.25);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
    }
    [data-testid="stMarkdownContainer"] { color: var(--panel-text); }
    [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3 {
        color: #F7FAFF;
        letter-spacing: -0.02em;
    }
    [data-testid="stCaptionContainer"], [data-testid="stWidgetLabel"] {
        color: var(--panel-muted);
    }
    [data-testid="stMetric"] {
        background: rgba(17, 24, 39, 0.82);
        border: 1px solid rgba(16, 185, 129, 0.34);
        padding: 1rem 1.1rem;
        border-radius: 14px;
        box-shadow: 0 10px 28px rgba(0, 0, 0, 0.18);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
    }
    [data-testid="stMetricLabel"] { color: var(--panel-muted); }
    [data-testid="stMetricValue"] { color: #F7FAFF; }
    .st-key-model-probability-high [data-testid="stMetricValue"],
    .st-key-market-probability-high [data-testid="stMetricValue"] {
        color: #10B981;
        font-weight: 800;
    }
    .st-key-model-probability-medium [data-testid="stMetricValue"],
    .st-key-market-probability-medium [data-testid="stMetricValue"] {
        color: #F59E0B;
    }
    .st-key-model-probability-low [data-testid="stMetricValue"],
    .st-key-market-probability-low [data-testid="stMetricValue"] {
        color: #EF4444;
    }
    div[data-testid="stTabs"] button {
        font-weight: 650;
        color: var(--panel-muted);
    }
    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: var(--panel-green);
        border-bottom-color: var(--panel-green);
    }
    div[data-testid="stDataFrame"] {
        border: 1px solid rgba(16, 185, 129, 0.28);
        border-radius: 12px;
        overflow: hidden;
        background: rgba(17, 24, 39, 0.82);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
    }
    [data-testid="stButton"] button {
        border: 1px solid #34415D;
        border-radius: 10px;
        background: #1A2236;
        color: var(--panel-text);
        transition: border-color 150ms ease, transform 150ms ease;
    }
    [data-testid="stButton"] button:hover {
        border-color: var(--panel-cyan);
        color: #F7FAFF;
        transform: translateY(-1px);
    }
    [data-baseweb="select"] > div, [data-baseweb="input"] > div,
    [data-testid="stDateInput"] input {
        border-color: #34415D;
        background: #1A2236;
        color: var(--panel-text);
    }
    [data-testid="stAlert"] {
        border: 1px solid var(--panel-border);
        border-radius: 12px;
    }
    .match-card-list {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
        gap: 1rem;
        margin-top: 0.75rem;
    }
    .match-card {
        min-width: 0;
        padding: 1rem;
        border: 1px solid rgba(16, 185, 129, 0.4);
        border-radius: 16px;
        background: rgba(17, 24, 39, 0.82);
        box-shadow: 0 12px 30px rgba(0, 0, 0, 0.2);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        transition: border-color 160ms ease, box-shadow 160ms ease, transform 160ms ease;
    }
    .match-card:hover {
        border-color: rgba(16, 185, 129, 0.72);
        box-shadow: 0 14px 34px rgba(0, 0, 0, 0.28), 0 0 18px rgba(16, 185, 129, 0.08);
        transform: translateY(-2px);
    }
    [data-testid="stExpander"] {
        margin-top: 0.5rem;
        border: 1px solid rgba(16, 185, 129, 0.36);
        border-radius: 14px;
        background: rgba(17, 24, 39, 0.82);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
    }
    [data-testid="stExpander"] summary {
        color: var(--panel-text);
    }
    [data-testid="stExpander"] summary:hover {
        color: #34D399;
    }
    .h2h-list {
        display: grid;
        gap: 0.5rem;
        margin: 0.5rem 0;
    }
    .h2h-row {
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        align-items: center;
        gap: 0.2rem 0.8rem;
        padding: 0.7rem 0.8rem;
        border: 1px solid rgba(16, 185, 129, 0.22);
        border-radius: 10px;
        background: rgba(17, 24, 39, 0.72);
    }
    .h2h-row > span {
        color: var(--panel-muted);
        font-size: 0.74rem;
    }
    .h2h-row > strong {
        grid-column: 1;
        color: var(--panel-text);
        font-size: 0.88rem;
    }
    .h2h-row > small {
        grid-column: 2;
        grid-row: 1 / span 2;
        color: var(--panel-green);
        font-size: 0.68rem;
        text-align: right;
    }
    .match-card-top, .match-teams, .match-xg {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 0.75rem;
    }
    .match-league {
        color: var(--panel-cyan);
        font-size: 0.86rem;
        font-weight: 750;
    }
    .match-date {
        color: var(--panel-muted);
        font-size: 0.78rem;
        text-align: right;
    }
    .match-teams {
        justify-content: center;
        margin: 1rem 0 0.85rem;
        color: #F7FAFF;
        font-size: 1.05rem;
        font-weight: 750;
        text-align: center;
    }
    .match-vs {
        color: var(--panel-green);
        font-size: 0.7rem;
        letter-spacing: 0.08em;
    }
    .match-tags {
        display: flex;
        flex-wrap: wrap;
        gap: 0.4rem;
        margin-bottom: 0.8rem;
    }
    .match-tag {
        padding: 0.25rem 0.55rem;
        border-radius: 999px;
        background: rgba(56, 189, 248, 0.12);
        color: #7DD3FC;
        font-size: 0.72rem;
        font-weight: 650;
    }
    .match-tag-muted {
        background: #222D43;
        color: var(--panel-muted);
    }
    .match-xg {
        flex-wrap: wrap;
        justify-content: flex-start;
        gap: 0.35rem 0.9rem;
        padding: 0.65rem 0;
        border-top: 1px solid var(--panel-border);
        border-bottom: 1px solid var(--panel-border);
        color: var(--panel-muted);
        font-size: 0.75rem;
    }
    .match-xg strong {
        color: var(--panel-text);
    }
    .probability-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.5rem;
        margin-top: 0.75rem;
    }
    .probability-box {
        min-width: 0;
        padding: 0.55rem 0.35rem;
        border: 1px solid rgba(16, 185, 129, 0.24);
        border-radius: 10px;
        background: rgba(17, 24, 39, 0.82);
        text-align: center;
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
    }
    .probability-box.prob-high {
        border-color: rgba(16, 185, 129, 0.45);
        background: rgba(16, 185, 129, 0.08);
    }
    .probability-box.prob-medium {
        border-color: rgba(245, 158, 11, 0.4);
        background: rgba(245, 158, 11, 0.07);
    }
    .probability-box.prob-low {
        border-color: rgba(239, 68, 68, 0.38);
        background: rgba(239, 68, 68, 0.07);
    }
    .probability-box span, .probability-box strong {
        display: block;
        overflow-wrap: anywhere;
    }
    .probability-box span {
        color: var(--panel-muted);
        font-size: 0.69rem;
    }
    .probability-box strong {
        margin-top: 0.15rem;
        color: #F7FAFF;
        font-size: 0.95rem;
    }
    .probability-box.prob-high strong {
        color: #10B981;
        font-weight: 800;
    }
    .probability-box.prob-medium strong {
        color: #F59E0B;
        font-weight: 650;
    }
    .probability-box.prob-low strong {
        color: #EF4444;
        font-weight: 650;
    }
    .probability-legend {
        display: flex;
        flex-wrap: wrap;
        gap: 0.45rem 0.9rem;
        align-items: center;
        margin: 0.15rem 0 0.9rem;
        color: #9AA8BF;
        font-size: 0.76rem;
    }
    .probability-legend-item {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        padding: 0.3rem 0.55rem;
        border: 1px solid #2A3650;
        border-radius: 999px;
        background: rgba(26, 34, 54, 0.72);
    }
    .probability-legend-dot {
        width: 0.5rem;
        height: 0.5rem;
        border-radius: 50%;
    }
    .probability-legend-high { background: #10B981; }
    .probability-legend-medium { background: #F59E0B; }
    .probability-legend-low { background: #EF4444; }
    .market-match-label {
        color: var(--panel-text);
        font-size: 0.9rem;
        font-weight: 700;
        line-height: 1.45;
    }
    .leaderboard-card {
        height: 100%;
        padding: 1.15rem;
        border: 1px solid rgba(16, 185, 129, 0.34);
        border-radius: 18px;
        background: rgba(17, 24, 39, 0.82);
        box-shadow: 0 12px 30px rgba(0, 0, 0, 0.2);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
    }
    .leaderboard-card-heading {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 0.75rem;
        margin-bottom: 0.75rem;
    }
    .leaderboard-card-heading h3 {
        margin: 0;
        color: #F7FAFF;
        font-size: 1rem;
        line-height: 1.4;
    }
    .leaderboard-top-label {
        flex: 0 0 auto;
        padding: 0.25rem 0.5rem;
        border: 1px solid rgba(56, 189, 248, 0.25);
        border-radius: 999px;
        background: rgba(56, 189, 248, 0.1);
        color: var(--panel-cyan);
        font-size: 0.65rem;
        font-weight: 750;
        letter-spacing: 0.08em;
    }
    .leaderboard-list {
        display: grid;
        gap: 0.15rem;
    }
    .leaderboard-row {
        display: grid;
        grid-template-columns: 1.8rem minmax(0, 1fr) auto;
        align-items: center;
        gap: 0.65rem;
        padding: 0.7rem 0;
        border-bottom: 1px solid rgba(154, 168, 191, 0.12);
    }
    .leaderboard-row:last-child { border-bottom: 0; }
    .leaderboard-rank {
        color: var(--panel-cyan);
        font-size: 0.75rem;
        font-weight: 750;
    }
    .leaderboard-entity { min-width: 0; }
    .leaderboard-entity strong, .leaderboard-entity small {
        display: block;
        overflow-wrap: anywhere;
    }
    .leaderboard-entity strong {
        color: #F1F5FB;
        font-size: 0.86rem;
        font-weight: 680;
    }
    .leaderboard-entity small {
        margin-top: 0.16rem;
        color: var(--panel-muted);
        font-size: 0.7rem;
    }
    .leaderboard-badge {
        padding: 0.35rem 0.6rem;
        border: 1px solid rgba(16, 185, 129, 0.28);
        border-radius: 999px;
        background: rgba(16, 185, 129, 0.12);
        color: #34D399;
        font-size: 0.78rem;
        font-weight: 750;
        white-space: nowrap;
    }
    .leaderboard-empty {
        padding: 1rem 0;
        color: var(--panel-muted);
        font-size: 0.85rem;
    }
    .st-key-featured-cards, .st-key-difference-cards,
    .st-key-comparison-selected-cards { display: none; }
    @media (max-width: 768px) {
        .block-container {
            padding: 1rem 0.75rem 2.5rem;
        }
        [data-testid="stHorizontalBlock"] {
            flex-direction: column !important;
            gap: 0.65rem !important;
        }
        [data-testid="stHorizontalBlock"] > [data-testid="column"] {
            flex: 1 1 100% !important;
            width: 100% !important;
            min-width: 100% !important;
        }
        [data-testid="stMetric"] {
            padding: 0.75rem 0.9rem;
        }
        [data-testid="stTabs"] [role="tablist"] {
            overflow-x: auto;
            flex-wrap: nowrap;
        }
        [data-testid="stTabs"] button {
            flex: 0 0 auto;
            white-space: nowrap;
        }
        [data-testid="stRadio"] [role="radiogroup"] {
            flex-wrap: wrap;
            gap: 0.4rem 0.8rem;
        }
        .st-key-featured-table, .st-key-difference-table {
            display: none !important;
        }
        .st-key-featured-cards, .st-key-difference-cards,
        .st-key-comparison-selected-cards {
            display: block !important;
        }
        .match-card-list {
            grid-template-columns: minmax(0, 1fr);
            gap: 0.75rem;
        }
        .match-card {
            padding: 0.85rem;
        }
        .leaderboard-card {
            padding: 0.95rem;
        }
        .probability-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }
    }
    </style>
    """.replace("__COVER_DATA_URI__", COVER_DATA_URI),
    unsafe_allow_html=True,
)

st.iframe(
    """
    <script>
    function setAppleTouchIcon(attempt = 0) {
        const doc = window.parent.document;
        const coverStyle = window.parent.getComputedStyle(doc.documentElement)
            .getPropertyValue("--panel-cover-image").trim();
        const coverUrl = coverStyle.replace(/^url\\(["']?/, "").replace(/["']?\\)$/, "");
        if (!coverUrl) {
            if (attempt < 20) window.setTimeout(() => setAppleTouchIcon(attempt + 1), 100);
            return;
        }

        let icon = doc.head.querySelector('link[rel="apple-touch-icon"]');
        if (!icon) {
            icon = doc.createElement("link");
            icon.rel = "apple-touch-icon";
            doc.head.appendChild(icon);
        }
        icon.href = coverUrl;
        icon.type = "image/jpeg";
        icon.sizes = "180x180";
    }
    setAppleTouchIcon();
    </script>
    """,
    width=1,
    height=1,
)

st.title("İ H T İ Y A R Maç Analiz Paneli")
st.caption(
    "Lig verileri, istatistiksel etiketler ve 0–6 gol Poisson model olasılıkları."
)
st.markdown(
    """
    <div class="probability-legend" aria-label="Olasılık renk açıklaması">
      <span class="probability-legend-item"><i class="probability-legend-dot probability-legend-high"></i>Yeşil: ≥%65 Yüksek</span>
      <span class="probability-legend-item"><i class="probability-legend-dot probability-legend-medium"></i>Sarı: %45-%64,9 Orta</span>
      <span class="probability-legend-item"><i class="probability-legend-dot probability-legend-low"></i>Kırmızı: &lt;%45 Düşük</span>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.divider()
    st.subheader("Veri kaynağı")
    st.caption("haftalik_analiz_raporu.xlsx")
    if st.button("Verileri yeniden yükle", width="stretch"):
        st.cache_data.clear()
        st.rerun()
    st.caption("Önbellek en fazla 60 saniyede bir yenilenir.")

try:
    report = load_report()
except (FileNotFoundError, ValueError) as exc:
    st.error(str(exc))
    st.stop()

try:
    market_odds = load_odds()
except ValueError as exc:
    st.error(str(exc))
    st.stop()

match_data_error = None
try:
    team_statistics = build_team_statistics(load_match_data())
except (FileNotFoundError, ValueError) as exc:
    team_statistics = build_team_statistics(pd.DataFrame())
    match_data_error = str(exc)

league_summary = report["Lig_Ozeti"]
standings = report["Puan_Durumu"]
upcoming = report["Yaklasan_Maclar"]
featured = report["One_Cikan_Maclar"]

view_mode = st.radio(
    "Görünüm",
    options=VIEW_OPTIONS,
    horizontal=True,
    key="match_view_mode",
)
if view_mode == "Mobil Görünüm":
    st.markdown(
        """
        <style>
        .st-key-featured-table, .st-key-difference-table { display: none !important; }
        .st-key-featured-cards, .st-key-difference-cards,
        .st-key-comparison-selected-cards { display: block !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )
elif view_mode == "Masaüstü Görünümü":
    st.markdown(
        """
        <style>
        .st-key-featured-table, .st-key-difference-table { display: block !important; }
        .st-key-featured-cards, .st-key-difference-cards,
        .st-key-comparison-selected-cards { display: none !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

utc_today = datetime.now(timezone.utc).date()
with st.container(border=True):
    date_filter = st.radio(
        "Maç tarih filtresi",
        options=DATE_FILTER_OPTIONS,
        horizontal=True,
        key="match_date_filter",
    )
    date_start = None
    date_end = None
    date_range_ready = True

    if date_filter == "Bugün":
        date_start = date_end = utc_today
    elif date_filter == "Önümüzdeki 3 Gün":
        date_start = utc_today
        date_end = utc_today + timedelta(days=2)
    elif date_filter == "Özel Tarih":
        selected_dates = st.date_input(
            "Başlangıç ve bitiş tarihi (UTC)",
            value=(utc_today, utc_today),
            key="match_date_range",
        )
        if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
            date_start, date_end = selected_dates
            if date_start > date_end:
                st.warning("Başlangıç tarihi bitiş tarihinden sonra olamaz.")
                date_range_ready = False
        else:
            st.info("Filtreyi uygulamak için başlangıç ve bitiş tarihini seçin.")
            date_range_ready = False

if date_range_ready:
    filtered_upcoming = filter_matches_by_date(upcoming, date_start, date_end)
    filtered_featured = filter_matches_by_date(featured, date_start, date_end)
else:
    filtered_upcoming = upcoming.iloc[0:0].copy()
    filtered_featured = featured.iloc[0:0].copy()

metric_columns = st.columns(3)
metric_columns[0].metric("Yaklaşan Maç", f"{len(filtered_upcoming):,}")
metric_columns[1].metric("Öne Çıkan Maç", f"{len(filtered_featured):,}")
metric_columns[2].metric(
    "Lig Sayısı",
    f"{filtered_upcoming['Lig'].nunique() if 'Lig' in filtered_upcoming else 0:,}",
)

featured_tab, comparison_tab, standings_tab, league_tab, leaders_tab = st.tabs(
    [
        "Öne Çıkan Maçlar",
        "Olasılık Karşılaştırma",
        "Puan Durumu",
        "Lig İstatistikleri",
        "İstatistik Liderleri",
    ]
)

with featured_tab:
    st.subheader("Öne Çıkan Maçlar")
    if featured.empty:
        st.info("Rapor dosyasında öne çıkan maç bulunmuyor.")
    elif filtered_featured.empty:
        st.info("Seçilen tarih aralığında öne çıkan maç bulunmuyor.")
    else:
        league_options = sorted(filtered_featured["Lig"].dropna().astype(str).unique())
        label_options = sorted(
            {
                label.strip()
                for value in filtered_featured["Istatistiksel_Etiketler"].dropna()
                for label in str(value).split(",")
                if label.strip()
            }
        )
        filter_columns = st.columns([1, 2])
        selected_leagues = filter_columns[0].multiselect(
            "Lig",
            options=league_options,
            default=league_options,
            format_func=league_label,
            key="featured_leagues",
        )
        selected_labels = filter_columns[1].multiselect(
            "İstatistiksel etiket",
            options=label_options,
            default=label_options,
            key="featured_labels",
        )

        filtered = filtered_featured.loc[
            filtered_featured["Lig"].isin(selected_leagues)
        ].copy()
        filtered = filtered.loc[
            filtered["Istatistiksel_Etiketler"].apply(
                lambda value: any(
                    label in [item.strip() for item in str(value).split(",")]
                    for label in selected_labels
                )
            )
        ]
        if filtered.empty:
            st.info("Seçilen filtrelerle eşleşen maç yok.")
        else:
            st.caption(f"Filtrelere uyan maç sayısı: {len(filtered)}")
            columns = [
                "Tarih (UTC)",
                "Lig",
                "Ev Sahibi",
                "Deplasman",
                "Istatistiksel_Etiketler",
                "Ev_xG_Lambda",
                "Dep_xG_Lambda",
                "MS1_%",
                "MSX_%",
                "MS2_%",
                "Ust_25_%",
                "KG_Var_%",
                JOINT_OVER_25_BOTH_SCORE_COLUMN,
                "En_Olasi_Skor",
            ]
            featured_with_joint = add_joint_probability_column(filtered)
            display = percent_display(
                featured_with_joint[columns], POISSON_PERCENT_COLUMNS
            )
            display["Lig"] = display["Lig"].map(league_label)
            display = display.rename(columns=FEATURED_COLUMN_LABELS)
            display = color_probability_cells(
                display,
                [
                    "MS1 %",
                    "MSX %",
                    "MS2 %",
                    "2.5 Üst %",
                    "KG Var %",
                    "2.5 Üst & KG Var %",
                ],
            )
            with st.container(key="featured-table"):
                st.dataframe(
                    display,
                    column_config={
                        **probability_column_config(
                            [
                                "MS1 %",
                                "MSX %",
                                "MS2 %",
                                "2.5 Üst %",
                                "KG Var %",
                                "2.5 Üst & KG Var %",
                            ]
                        ),
                        "EV Gol Ort. (xG)": st.column_config.NumberColumn(
                            "EV Gol Ort. (xG)", format="%.2f"
                        ),
                        "DEP Gol Ort. (xG)": st.column_config.NumberColumn(
                            "DEP Gol Ort. (xG)", format="%.2f"
                        ),
                    },
                    hide_index=True,
                    width="stretch",
                )
            with st.container(key="featured-cards"):
                render_match_cards(featured_with_joint)

with comparison_tab:
    st.subheader("Model ve Piyasa Oranı Karşılaştırması")
    st.caption(
        "Piyasa örtük olasılığı 100 / oran olarak hesaplanır. Fark yüzde puan "
        "cinsindendir; karşılaştırma eğitim amaçlıdır, oran marjı düzeltilmez "
        "ve bahis önerisi değildir."
    )
    if upcoming.empty:
        st.info("Karşılaştırma için yaklaşan maç bulunmuyor.")
    elif filtered_upcoming.empty:
        st.info("Seçilen tarih aralığında yaklaşan maç bulunmuyor.")
    else:
        match_options = list(range(len(filtered_upcoming)))
        selected_match_index = st.selectbox(
            "Maç",
            options=match_options,
            format_func=lambda index: match_label(filtered_upcoming.iloc[index]),
            key="comparison_match",
        )
        selected_match = filtered_upcoming.iloc[selected_match_index]
        selected_match_with_joint = add_joint_probability_column(
            filtered_upcoming.iloc[[selected_match_index]]
        )
        with st.container(key="comparison-selected-cards"):
            render_match_cards(selected_match_with_joint)
        market_match, match_similarity = find_market_odds(
            selected_match, market_odds
        )
        if market_match is None:
            st.info(
                "Bu maç için otomatik piyasa eşleşmesi bulunamadı. "
                "Oran kutularına değerleri elle girebilirsiniz."
            )
        else:
            st.success(
                "Takım adları eşleşti "
                f"(benzerlik: {match_similarity:.1%}); oranlar büro "
                "ortalamalarıyla dolduruldu. Kutuları değiştirebilirsiniz."
            )

        odds_fields = [
            ("Ev Sahibi Oranı (MS1)", "Ort_MS1"),
            ("Beraberlik Oranı (MSX)", "Ort_MSX"),
            ("Deplasman Oranı (MS2)", "Ort_MS2"),
            ("2.5 Üst Oranı", "Ort_Ust25"),
        ]
        odds_inputs = {}
        odds_columns = st.columns(4)
        for column, (label, odds_column) in zip(odds_columns, odds_fields):
            default_odds = 2.00
            if market_match is not None:
                market_value = pd.to_numeric(
                    pd.Series([market_match[odds_column]]), errors="coerce"
                ).iloc[0]
                if pd.notna(market_value) and market_value >= 1.01:
                    default_odds = float(market_value)
            odds_inputs[odds_column] = column.number_input(
                label,
                min_value=1.01,
                max_value=1000.0,
                value=default_odds,
                step=0.01,
                format="%.2f",
                key=f"market_{selected_match_index}_{odds_column}",
            )

        outcome_label = st.selectbox(
            "Karşılaştırılacak sonuç",
            options=list(OUTCOME_COLUMNS),
            key="comparison_outcome",
        )
        model_column, market_column = OUTCOME_COLUMNS[outcome_label]
        if market_column is None:
            odds = st.number_input(
                "KG Var manuel oranı",
                min_value=1.01,
                max_value=1000.0,
                value=2.00,
                step=0.01,
                format="%.2f",
                key=f"manual_kg_{selected_match_index}",
            )
        else:
            odds = odds_inputs[market_column]

        model_probability = float(selected_match[model_column]) * 100
        implied_probability = 100.0 / odds
        difference = model_probability - implied_probability
        model_probability_level = probability_level(model_probability) or "na"
        market_probability_level = probability_level(implied_probability) or "na"
        result_columns = st.columns(3)
        with result_columns[0].container(
            key=f"model-probability-{model_probability_level}"
        ):
            st.metric("Model Olasılığı", f"{model_probability:.2f}%")
        with result_columns[1].container(
            key=f"market-probability-{market_probability_level}"
        ):
            st.metric(
                "Piyasanın Örtük Olasılığı", f"{implied_probability:.2f}%"
            )
        result_columns[2].metric(
            "Model − Piyasa Farkı",
            f"{difference:+.2f} yüzde puan",
            delta_color="off",
        )
        st.info(
            f"Model xG/Lambda: ev sahibi {selected_match['Ev_xG_Lambda']:.2f} — "
            f"deplasman {selected_match['Dep_xG_Lambda']:.2f}; "
            f"en olası skor {selected_match['En_Olasi_Skor']}."
        )

        st.divider()
        st.subheader("Model - Piyasa Farkı Tablosu")
        st.caption(
            "Takım adı eşleşen maçlarda, Poisson olasılığı − 100 / ortalama oran "
            "farkı en az +5 yüzde puan olan sonuçlar listelenir."
        )
        if market_odds.empty:
            st.info("piyasa_oranlari.xlsx bulunamadı veya henüz oran içermiyor.")
        else:
            difference_table, matched_count = build_market_difference_table(
                filtered_upcoming, market_odds
            )
            st.metric("Piyasa oranıyla eşleşen maç", matched_count)
            if difference_table.empty:
                st.info("+5 yüzde puan eşiğini geçen eşleşme bulunamadı.")
            else:
                display = percent_display(
                    difference_table,
                    [
                        "Model Olasılığı (%)",
                        "Örtük Piyasa Olasılığı (%)",
                        "Takım Adı Eşleşmesi (%)",
                    ],
                )
                display = display.rename(
                    columns={
                        "Piyasa": "Sonuç",
                        "Ortalama Oran": "Ort. Oran",
                        "Model Olasılığı (%)": "Model Olasılığı",
                        "Örtük Piyasa Olasılığı (%)": "Piyasa Olasılığı",
                        "Takım Adı Eşleşmesi (%)": "Takım Adı Uyumu",
                    }
                )
                display = color_probability_cells(
                    display, ["Model Olasılığı", "Piyasa Olasılığı"]
                )
                with st.container(key="difference-table"):
                    st.dataframe(
                        display,
                        column_config={
                            "Ort. Oran": st.column_config.NumberColumn(
                                "Ort. Oran", format="%.2f"
                            ),
                            "Model Olasılığı": st.column_config.NumberColumn(
                                "Model Olasılığı", format="%.1f%%"
                            ),
                            "Piyasa Olasılığı": st.column_config.NumberColumn(
                                "Piyasa Olasılığı", format="%.1f%%"
                            ),
                            "Fark (yüzde puan)": st.column_config.NumberColumn(
                                "Fark (yüzde puan)", format="%.2f"
                            ),
                            "Takım Adı Uyumu": st.column_config.NumberColumn(
                                "Takım Adı Uyumu", format="%.1f%%"
                            ),
                        },
                        hide_index=True,
                        width="stretch",
                    )
                with st.container(key="difference-cards"):
                    render_market_difference_cards(difference_table)

with standings_tab:
    st.subheader("Puan Durumu")
    if standings.empty:
        st.info("Puan durumu verisi bulunmuyor.")
    else:
        league_options = sorted(standings["Lig"].dropna().astype(str).unique())
        selected_league = st.selectbox(
            "Lig Seçin",
            league_options,
            format_func=league_label,
            key="standings_league",
        )
        venue_label = st.radio(
            "Puan tablosu",
            options=["İç saha", "Deplasman"],
            horizontal=True,
            key="standings_venue",
        )
        venue_value = "Ic Saha" if venue_label == "İç saha" else "Deplasman"
        table = standings.loc[
            (standings["Lig"] == selected_league)
            & (standings["Saha"] == venue_value)
        ].sort_values(["Sira", "Puan"], ascending=[True, False]).copy()
        table["Lig"] = table["Lig"].map(league_label)
        table["Saha"] = table["Saha"].map(
            {"Ic Saha": "İç Saha", "Deplasman": "Deplasman"}
        )
        table = table.rename(columns=STANDINGS_COLUMN_LABELS)
        st.dataframe(table, hide_index=True, width="stretch")

with league_tab:
    st.subheader("Lig İstatistikleri")
    if league_summary.empty:
        st.info("Lig özeti bulunmuyor.")
    else:
        summary_display = percent_display(
            league_summary,
            ["Ev Sahibi Kazanma %", "2.5 Ust %", "KG Var %"],
        )
        summary_display["Lig"] = summary_display["Lig"].map(league_label)
        summary_display = summary_display.rename(
            columns=LEAGUE_SUMMARY_COLUMN_LABELS
        )
        summary_display = color_probability_cells(
            summary_display,
            ["Ev Sahibi Galibiyeti %", "2.5 Üst %", "KG Var %"],
        )
        st.dataframe(
            summary_display,
            column_config={
                "Ev Sahibi Galibiyeti %": st.column_config.NumberColumn(
                    "Ev Sahibi Galibiyeti %", format="%.1f%%"
                ),
                "2.5 Üst %": st.column_config.NumberColumn(
                    "2.5 Üst %", format="%.1f%%"
                ),
                "KG Var %": st.column_config.NumberColumn(
                    "KG Var %", format="%.1f%%"
                ),
                "Maç Başı Gol Ort.": st.column_config.NumberColumn(
                    "Maç Başı Gol Ort.", format="%.2f"
                ),
            },
            hide_index=True,
            width="stretch",
        )
        if CHART_PATH.is_file():
            st.image(
                str(CHART_PATH),
                caption="Liglerin maç başı gol ortalamaları ve 2.5 üst oranları",
                width="stretch",
            )
        else:
            st.warning(f"Lig karşılaştırma grafiği bulunamadı: {CHART_PATH}")

with leaders_tab:
    st.subheader("İstatistik Liderleri")
    render_leaderboards(league_summary, team_statistics, match_data_error)

st.divider()
admin_column, _ = st.columns([1, 3])
with admin_column:
    admin_password = st.text_input(
        "Yönetici parolası",
        type="password",
        key="admin_update_password",
    )
    if admin_password.strip().lower() == "do258025":
        if st.button(
            "Verileri Şimdi Güncelle",
            type="primary",
            width="stretch",
            key="run-data-update",
        ):
            if run_update_pipeline():
                st.cache_data.clear()
                st.session_state["data_update_success"] = True
                st.rerun()

if st.session_state.pop("data_update_success", False):
    st.success("Veriler, rapor ve grafik güncellendi.")
