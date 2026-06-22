from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT.parent / "LitSearchPro_v22.1.21_Generic_Complete_Source.zip"

EXCLUDED_PARTS = {"build", "dist", "installer_payload", "__pycache__"}


def main():
    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(ROOT.rglob("*")):
            if not path.is_file():
                continue
            relative = path.relative_to(ROOT)
            if any(part in EXCLUDED_PARTS for part in relative.parts):
                continue
            archive.write(path, Path(ROOT.name) / relative)
    print(OUTPUT)


if __name__ == "__main__":
    main()
