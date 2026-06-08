import sys
from pathlib import Path


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


ROOT = app_root()
REPO_ROOT = ROOT.parent

for p in [ROOT, REPO_ROOT]:
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))


def main() -> int:
    try:
        from Altomizer.desktop_qt import main as desktop_main
    except Exception as exc:
        import traceback
        print("Altomizer Desktop could not start.")
        print("Crash:")
        traceback.print_exc()
        input("Press Enter to close...")
        return 1

    return int(desktop_main())


if __name__ == "__main__":
    raise SystemExit(main())