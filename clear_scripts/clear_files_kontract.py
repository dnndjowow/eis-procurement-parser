import sys
from pathlib import Path


BASE_PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_PROJECT_DIR))

from script_kontracts.clear_files import main  # noqa: E402


if __name__ == "__main__":
    main()
