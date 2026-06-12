import os
import subprocess
import sys
import venv
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent
VENV_DIR = REPO_ROOT / 'venv'
REQUIREMENTS_FILE = REPO_ROOT / 'requirements.txt'
REQUIRED_MODULES = [
    'dotenv',
    'pandas',
    'numpy',
    'openpyxl',
    'docx',
    'supabase',
]


def get_venv_python():
    if os.name == 'nt':
        return VENV_DIR / 'Scripts' / 'python.exe'
    return VENV_DIR / 'bin' / 'python'


def in_venv():
    return hasattr(sys, 'real_prefix') or getattr(sys, 'base_prefix', sys.prefix) != sys.prefix


def dependencies_missing():
    for module_name in REQUIRED_MODULES:
        try:
            __import__(module_name)
        except ImportError:
            return True
    return False


def install_requirements(python_executable):
    if not REQUIREMENTS_FILE.exists():
        raise FileNotFoundError(f'ไม่พบไฟล์ requirements.txt ที่ {REQUIREMENTS_FILE}')

    print('[setup] Installing dependencies from requirements.txt...')
    subprocess.check_call(
        [str(python_executable), '-m', 'pip', 'install', '-r', str(REQUIREMENTS_FILE)]
    )


def create_virtualenv():
    if VENV_DIR.exists():
        print('[setup] Found existing virtual environment in venv/')
    else:
        print('[setup] Creating a virtual environment in venv/')
        venv.create(VENV_DIR, with_pip=True)

    python_path = get_venv_python()
    if not python_path.exists():
        raise FileNotFoundError(f'ไม่พบ Python ใน virtualenv: {python_path}')
    return python_path


def bootstrap():
    if in_venv():
        if dependencies_missing():
            install_requirements(sys.executable)
        return

    venv_python = create_virtualenv()
    install_requirements(venv_python)
    print('[setup] Re-running generate_report.py inside venv...')
    os.execv(str(venv_python), [str(venv_python)] + sys.argv)


bootstrap()

from data_loader import (
    create_supabase_client,
    fetch_buildings,
    fetch_survey_data,
    fetch_traceroute_hops,
)
from gui import select_building
from report_builder import ReportBuilder


def make_output_dir(target_building):
    output_dir = os.path.join('Reports', target_building)
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def main():
    try:
        client = create_supabase_client()
    except EnvironmentError as error:
        print(f'❌ {error}')
        return 1

    buildings = fetch_buildings(client)
    if not buildings:
        print('❌ ไม่พบข้อมูลตึกใน Supabase')
        return 1

    target_building = select_building(buildings)
    if not target_building:
        print('❌ ไม่ได้เลือกตึก หยุดการทำงาน')
        return 1

    df = fetch_survey_data(client, target_building)
    print(f'surveys: {len(df)} แถว')
    if df.empty:
        print('❌ ไม่พบข้อมูลการสำรวจสำหรับตึกนี้')
        return 1

    survey_ids = df['id'].tolist() if 'id' in df.columns else []
    df_hops = fetch_traceroute_hops(client, survey_ids)
    print(f'traceroute_hops: {len(df_hops)} แถว')

    output_dir = make_output_dir(target_building)
    builder = ReportBuilder(df, df_hops, target_building, output_dir)

    excel_path = builder.build_excel()
    print(f'✅ Excel: {excel_path}')

    word_path = builder.build_word(excel_path)
    print(f'✅ Word:  {word_path}')
    print('✅ เสร็จสมบูรณ์!')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
