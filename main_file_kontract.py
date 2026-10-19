import sys
from pathlib import Path


BASE_PROJECT_DIR = Path(__file__).resolve().parent
CONTRACTS_DIR = BASE_PROJECT_DIR / "script_kontracts"
sys.path.insert(0, str(BASE_PROJECT_DIR))
sys.path.insert(0, str(CONTRACTS_DIR))

from script_kontracts.main_file_kontract import main, parse_args  # noqa: E402


if __name__ == "__main__":
    args = parse_args()
    main(
        write_enabled=args.write,
        refresh_publication_from_main=True,
        apply_changed_filter=True,
    )
