from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEXT_SUFFIXES = {".py", ".spec", ".md"}
FORBIDDEN = [
    "重庆大学",
    "Chongqing University",
    "cqu_logo",
    "lib.cqu.edu.cn",
    "www.cqu.edu.cn",
    "i.cqu.edu.cn",
    "huxi.cqu.edu.cn",
]


def main():
    failures = []
    for path in ROOT.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in TEXT_SUFFIXES:
            continue
        if "__pycache__" in path.parts or path.name == "test_generic_branding.py":
            continue
        text = path.read_text(encoding="utf-8", errors="ignore")
        for marker in FORBIDDEN:
            if marker in text:
                failures.append(f"{path.relative_to(ROOT)}: {marker}")
    assert not failures, "\n".join(failures)
    print("generic branding scan OK")


if __name__ == "__main__":
    main()
