import shutil
import sys
from pathlib import Path


BASE_PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_PROJECT_DIR))

from script_kontracts.settings import BASE_DIR, FAS_DIR


def clear_directory(directory):
    directory.mkdir(exist_ok=True)
    removed = 0

    for path in directory.iterdir():
        if path.is_dir():
            shutil.rmtree(path)
        else:
            path.unlink()
        removed += 1

    return removed


def main():
    print("Будет очищена папка документов ФАС:")
    print(f"- {FAS_DIR.relative_to(BASE_DIR)}")

    answer = input("Удалить все файлы из этой папки? Напиши YES для подтверждения: ")
    if answer != "YES":
        print("Очистка отменена")
        return

    removed = clear_directory(FAS_DIR)
    print(f"{FAS_DIR.relative_to(BASE_DIR)}: удалено {removed}")
    print("Готово")


if __name__ == "__main__":
    main()
