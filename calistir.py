import os
import subprocess
import sys
from pathlib import Path


API_KEY_ENV = "FOOTBALL_DATA_API_KEY"
STEPS = [
    ("super_lig_cek.py", "TFF Süper Lig verilerini güncelleme"),
    ("detayli_analiz.py", "Haftalık analiz raporu ve grafiği oluşturma"),
]


def run_step(base_dir, number, total, script_name, description, environment):
    script_path = base_dir / script_name
    print(f"\n[{number}/{total}] Başlıyor: {description}", flush=True)

    if not script_path.is_file():
        print(f"[{number}/{total}] Bitti: başarısız — {script_name} bulunamadı.", flush=True)
        return False

    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=base_dir,
            env=environment,
            check=False,
        )
    except OSError as exc:
        print(f"[{number}/{total}] Bitti: başarısız — {exc}", flush=True)
        return False

    if result.returncode != 0:
        print(
            f"[{number}/{total}] Bitti: başarısız "
            f"(çıkış kodu {result.returncode}).",
            flush=True,
        )
        return False

    print(f"[{number}/{total}] Bitti: başarılı.", flush=True)
    return True


def run_api_step(base_dir, environment):
    number = 1
    total = 3
    script_name = "api_veri_cek.py"
    description = "API lig verilerini indirme"
    script_path = base_dir / script_name
    print(f"\n[{number}/{total}] Başlıyor: {description}", flush=True)

    api_environment = environment.copy()
    api_key = api_environment.get(API_KEY_ENV, "").strip()
    if not api_key:
        api_key = input("football-data.org API anahtarını girin: ").strip()
    if not api_key:
        print(
            f"[{number}/{total}] Bitti: başarısız — API anahtarı girilmedi.",
            flush=True,
        )
        return False
    api_environment[API_KEY_ENV] = api_key

    if not script_path.is_file():
        print(f"[{number}/{total}] Bitti: başarısız — {script_name} bulunamadı.", flush=True)
        return False

    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=base_dir,
            env=api_environment,
            check=False,
        )
    except OSError as exc:
        print(f"[{number}/{total}] Bitti: başarısız — {exc}", flush=True)
        return False

    if result.returncode != 0:
        print(
            f"[{number}/{total}] Bitti: başarısız "
            f"(çıkış kodu {result.returncode}).",
            flush=True,
        )
        return False

    print(f"[{number}/{total}] Bitti: başarılı.", flush=True)
    return True


def main():
    base_dir = Path(__file__).resolve().parent
    environment = os.environ.copy()

    if not run_api_step(base_dir, environment):
        print("Otomasyon durduruldu; sonraki adımlar çalıştırılmadı.", flush=True)
        return 1

    for step_number, (script_name, description) in enumerate(STEPS, start=2):
        if not run_step(
            base_dir,
            step_number,
            3,
            script_name,
            description,
            environment,
        ):
            print("Otomasyon durduruldu; sonraki adımlar çalıştırılmadı.", flush=True)
            return 1

    print("\nTüm otomasyon adımları başarıyla tamamlandı.", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
