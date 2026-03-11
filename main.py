import sys
from pathlib import Path


# MVP: проект без упаковки/инсталляции, поэтому добавляем `src` в sys.path.
ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from filldoc.app import run  # noqa: E402


if __name__ == "__main__":
    run()

