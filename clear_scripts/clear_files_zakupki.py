import shutil
import sys
from pathlib import Path


BASE_PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_PROJECT_DIR))

from script_zakupki.settings import BASE_DIR, ZAKUPKI_FILES_DIR


TARGET_DIRS = [
    ZAKUPKI_FILES_DIR,
]


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
    print("Будут очищены папки закупок:")
    for directory in TARGET_DIRS:
        print(f"- {directory.relative_to(BASE_DIR)}")

    answer = input("Удалить все файлы из этих папок? Напиши YES для подтверждения: ")
    if answer != "YES":
        print("Очистка отменена")
        return

    total_removed = 0
    for directory in TARGET_DIRS:
        removed = clear_directory(directory)
        total_removed += removed
        print(f"{directory.relative_to(BASE_DIR)}: удалено {removed}")

    print(f"Готово. Всего удалено: {total_removed}")


if __name__ == "__main__":
    main()
