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


def generate_report(target_building, update_status, client):
    """ฟังก์ชันสำหรับสร้างรายงาน พร้อม callback สำหรับอัพเดตสถานะ"""
    try:
        update_status(f'📊 กำลังโหลดข้อมูลการสำรวจ...')
        df = fetch_survey_data(client, target_building)
        update_status(f'✓ โหลดข้อมูลแล้ว: {len(df)} แถว')
        
        if df.empty:
            update_status('❌ ไม่พบข้อมูลการสำรวจสำหรับตึกนี้')
            return 1

        update_status(f'📡 กำลังโหลด traceroute hops...')
        survey_ids = df['id'].tolist() if 'id' in df.columns else []
        df_hops = fetch_traceroute_hops(client, survey_ids)
        update_status(f'✓ โหลด traceroute hops แล้ว: {len(df_hops)} แถว')

        output_dir = make_output_dir(target_building)
        builder = ReportBuilder(df, df_hops, target_building, output_dir)

        update_status('📝 กำลังสร้าง Excel...')
        excel_path = builder.build_excel()
        update_status(f'✅ Excel: {excel_path}')

        update_status('📄 กำลังสร้าง Word...')
        word_path = builder.build_word(excel_path)
        update_status(f'✅ Word:  {word_path}')
        
        update_status('✅ เสร็จสมบูรณ์!')
        return 0
    except Exception as e:
        update_status(f'❌ เกิดข้อผิดพลาด: {str(e)}')
        return 1


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

    def on_generate(target_building, update_status):
        """Callback สำหรับการ generate ที่เรียกจาก GUI"""
        return generate_report(target_building, update_status, client)

    target_building = select_building(buildings, on_generate)
    if not target_building:
        print('❌ ไม่ได้เลือกตึก หยุดการทำงาน')
        return 1

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
