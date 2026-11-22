from io import BytesIO
from pathlib import Path
from urllib.request import Request, urlopen

import pandas as pd


BASE_URL = "https://www.football-data.co.uk/mmz4281/2627"
OUTPUT_NAME = "tum_ligler_2627.xlsx"

LEAGUES = {
    "T1": "Turkiye Super Lig",
    "E0": "Premier Lig",
    "SP1": "La Liga",
    "I1": "Serie A",
    "D1": "Bundesliga",
    "F1": "Ligue 1",
    "N1": "Eredivisie",
    "P1": "Portekiz",
}


def download_csv(code):
    url = f"{BASE_URL}/{code}.csv"
    request = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "text/csv,*/*",
        },
    )
    with urlopen(request, timeout=30) as response:
        if response.status != 200:
            raise RuntimeError(f"{url} indirilemedi (HTTP {response.status}).")
        return pd.read_csv(BytesIO(response.read()), encoding="utf-8-sig")


def main():
    base_dir = Path(__file__).resolve().parent
    output_path = base_dir / OUTPUT_NAME
    frames = []

    for code, league_name in LEAGUES.items():
        try:
            frame = download_csv(code)
        except Exception as exc:
            raise RuntimeError(
                f"{code} lig CSV dosyasi indirilemedi: {exc}"
            ) from exc

        required_columns = {"FTHG", "FTAG", "FTR"}
        missing = required_columns.difference(frame.columns)
        if missing:
            raise KeyError(
                f"{code} CSV dosyasinda beklenen sutunlar yok: "
                + ", ".join(sorted(missing))
            )

        frame = frame.copy()
        frame["Lig"] = league_name
        frames.append(frame)

    combined = pd.concat(frames, ignore_index=True)
    combined.to_excel(output_path, index=False)
    print(
        f"{len(combined):,} mac kaydi, {len(frames)} lig "
        f"'{OUTPUT_NAME}' dosyasina kaydedildi."
    )


if __name__ == "__main__":
    main()