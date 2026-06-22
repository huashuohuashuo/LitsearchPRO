from pathlib import Path
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
PYTHON = sys.executable


def run_spec(name):
    subprocess.run([PYTHON, "-m", "PyInstaller", "--noconfirm", str(ROOT / name)], cwd=ROOT, check=True)


def main():
    subprocess.run([PYTHON, str(ROOT / "generate_generic_logo.py")], cwd=ROOT, check=True)
    run_spec("LitSearchPro_Generic_Client.spec")
    run_spec("LitSearchPro_Generic_Server.spec")
    run_spec("LitSearchPro_Generic_Uninstall.spec")
    run_spec("LitSearchPro_Generic_Server_Uninstall.spec")
    payload = ROOT / "installer_payload"
    payload.mkdir(exist_ok=True)
    shutil.copy2(ROOT / "dist" / "LitSearchPro_Generic_v22.1.21_Uninstall.exe",
                 payload / "LitSearchPro_Generic_v22.1.21_Uninstall.exe")
    shutil.copy2(ROOT / "dist" / "LitSearchPro_Generic_Server_v22.1.21_Uninstall.exe",
                 payload / "LitSearchPro_Generic_Server_v22.1.21_Uninstall.exe")
    run_spec("LitSearchPro_Generic_Setup.spec")
    run_spec("LitSearchPro_Generic_Server_Setup.spec")
    print(ROOT / "dist")


if __name__ == "__main__":
    main()
