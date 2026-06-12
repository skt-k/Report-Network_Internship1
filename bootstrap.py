import os
import subprocess
import sys
import venv
from pathlib import Path


VENV_DIR = Path("venv")
REQUIREMENTS_FILE = Path("requirements.txt")


def get_venv_python():
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def create_virtualenv():
    if VENV_DIR.exists():
        print("[info] พบ virtual environment เดิมที่ venv/ แล้ว")
    else:
        print("[info] สร้าง virtual environment ที่ venv/")
        venv.create(VENV_DIR, with_pip=True)

    python_path = get_venv_python()
    if not python_path.exists():
        raise FileNotFoundError(f"ไม่พบ Python ใน virtualenv: {python_path}")
    return python_path


def install_requirements(python_path):
    if not REQUIREMENTS_FILE.exists():
        raise FileNotFoundError(f"ไม่พบไฟล์ requirements.txt ที่ {REQUIREMENTS_FILE}")

    print("[info] ติดตั้ง dependencies จาก requirements.txt...")
    command = [str(python_path), "-m", "pip", "install", "-r", str(REQUIREMENTS_FILE)]
    subprocess.check_call(command)


def main():
    try:
        python_path = create_virtualenv()
        install_requirements(python_path)
    except subprocess.CalledProcessError as exc:
        print(f"❌ ติดตั้ง dependencies ล้มเหลว: {exc}")
        return 1
    except Exception as exc:
        print(f"❌ เกิดข้อผิดพลาด: {exc}")
        return 1

    print("\n✅ ติดตั้ง dependencies สำเร็จแล้ว")
    print("วิธีใช้งานต่อ:")
    if os.name == "nt":
        print("  venv\\Scripts\\activate")
    else:
        print("  source venv/bin/activate")
    print("  python generate_report.py")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
