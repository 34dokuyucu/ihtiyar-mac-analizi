from pathlib import Path
import subprocess
import sys

import pandas as pd


EXCEL_NAME = "tum_ligler_2627.xlsx"


def find_column(columns, candidates):
    normalized = {str(column).strip().lower(): column for column in columns}
    for candidate in candidates:
        if candidate.lower() in normalized:
            return normalized[candidate.lower()]
    return None


def read_excel(path):
    workbook = pd.ExcelFile(path)
    frames = []
    for sheet_name in workbook.sheet_names:
        frame = pd.read_excel(path, sheet_name=sheet_name)
        if not frame.empty:
            frames.append(frame)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def main():
    base_dir = Path(__file__).resolve().parent
    excel_path = base_dir / EXCEL_NAME

    if not excel_path.exists():
        generator = base_dir / "veri_cek.py"
        if not generator.exists():
            raise FileNotFoundError("veri_cek.py bulunamadi.")

        result = subprocess.run(
            [sys.executable, str(generator)],
            cwd=base_dir,
            text=True,
            capture_output=True,
        )
        if result.stdout:
            print(result.stdout, end="")
        if result.stderr:
            print(result.stderr, end="", file=sys.stderr)
        if result.returncode != 0:
            raise RuntimeError(
                f"veri_cek.py dosya olustururken hata verdi "
                f"(cikis kodu {result.returncode})."
            )

    if not excel_path.exists():
        raise FileNotFoundError(
            f"{EXCEL_NAME} veri_cek.py calistirilinca da olusturulmadi."
        )

    data = read_excel(excel_path)
    if data.empty:
        raise ValueError(f"{EXCEL_NAME} icinde veri bulunamadi.")

    league_column = find_column(
        data.columns,
        ["Lig", "League", "Div", "Division", "LeagueName", "LigAdi"],
    )
    home_goals_column = find_column(
        data.columns,
        ["EvSahibiGol", "FTHG", "HomeGoals", "HG", "HGols", "EvGol"],
    )
    away_goals_column = find_column(
        data.columns,
        ["DeplasmanGol", "FTAG", "AwayGoals", "AG", "AGols", "DepGol"],
    )
    result_column = find_column(
        data.columns,
        ["Sonuc", "FTR", "Result", "FullTimeResult", "MacSonucu"],
    )

    missing = [
        name
        for name, column in (
            ("lig", league_column),
            ("ev sahibi golu", home_goals_column),
            ("deplasman golu", away_goals_column),
            ("mac sonucu", result_column),
        )
        if column is None
    ]
    if missing:
        raise KeyError(
            "Gerekli Excel sutunlari bulunamadi: "
            + ", ".join(missing)
            + ". Mevcut sutunlar: "
            + ", ".join(map(str, data.columns))
        )

    home_goals = pd.to_numeric(data[home_goals_column], errors="coerce")
    away_goals = pd.to_numeric(data[away_goals_column], errors="coerce")
    league = data[league_column]
    result = data[result_column].astype("string").str.strip().str.upper()

    usable = league.notna() & home_goals.notna() & away_goals.notna() & result.notna()
    league = league[usable].astype(str).str.strip()
    home_goals = home_goals[usable]
    away_goals = away_goals[usable]
    result = result[usable]

    if league.empty:
        raise ValueError("Hesaplama icin kullanilabilir mac kaydi bulunamadi.")

    analysis = pd.DataFrame(
        {
            "lig": league,
            "toplam_gol": home_goals + away_goals,
            "ev_sahibi_kazandi": result.isin(["H", "1", "EV SAHIBI"]),
            "iki_bucuk_ustu": (home_goals + away_goals) > 2.5,
        }
    )

    summary = analysis.groupby("lig", sort=True).agg(
        toplam_mac_sayisi=("lig", "size"),
        mac_basi_ortalama_gol=("toplam_gol", "mean"),
        ev_sahibi_kazanma_yuzdesi=("ev_sahibi_kazandi", "mean"),
        iki_bucuk_ust_bitme_orani=("iki_bucuk_ustu", "mean"),
    )
    summary["ev_sahibi_kazanma_yuzdesi"] *= 100
    summary["iki_bucuk_ust_bitme_orani"] *= 100
    summary = summary.reset_index()
    summary = summary.rename(
        columns={
            "lig": "Lig",
            "toplam_mac_sayisi": "Toplam Mac Sayisi",
            "mac_basi_ortalama_gol": "Mac Basi Ortalama Gol",
            "ev_sahibi_kazanma_yuzdesi": "Ev Sahibi Kazanma %",
            "iki_bucuk_ust_bitme_orani": "2.5 Ustu %",
        }
    )

    summary["Mac Basi Ortalama Gol"] = summary["Mac Basi Ortalama Gol"].map(
        lambda value: f"{value:.2f}"
    )
    summary["Ev Sahibi Kazanma %"] = summary["Ev Sahibi Kazanma %"].map(
        lambda value: f"{value:.2f}%"
    )
    summary["2.5 Ustu %"] = summary["2.5 Ustu %"].map(
        lambda value: f"{value:.2f}%"
    )

    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
