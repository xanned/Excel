import os
import re
from pathlib import Path

import pandas as pd

new_dir = 'new data'
joined_tables = 'joined_data.xlsx'
all_tables = 'all_tables_data.xlsx'
SKIP_ROWS = 2


def find_files(path) -> list[Path]:
    files = [file for file in Path(path).iterdir() if file.is_file() and re.search(r'.\.xlsx$', file.name)]
    return files


def open_excel_files(file: Path):
    df = pd.read_excel(file, engine="openpyxl", skiprows=SKIP_ROWS)
    df = df.dropna(axis='rows', how='all')
    df = df.dropna(axis='columns', how='all')
    return df


def main():
    user_input = input("Введите путь до папки c файлами таблиц: \n(Enter для текущей папки)\n").strip()
    if user_input == "":
        directory = os.path.abspath(os.curdir)
    elif Path(user_input).is_dir():
        directory = user_input
    else:
        print("Неверно указан путь")
        return
    files = find_files(directory)

    if not files:
        print("Файлы не найдены")
        return

    new_directory = Path(directory).joinpath(new_dir)
    joined_tables_file = new_directory.joinpath(joined_tables)
    all_tables_file = new_directory.joinpath(all_tables)
    Path(new_directory).mkdir(exist_ok=True)

    tables = []
    writer = pd.ExcelWriter(all_tables_file, engine='openpyxl')

    for file in files:
        data = open_excel_files(file)
        tables.append(data)
        sheet_name = file.stem
        data.to_excel(writer, sheet_name=sheet_name, index=False)

    df = pd.concat(tables, join='inner', ignore_index=True)
    df.to_excel(joined_tables_file, index=False)
    writer.close()


if __name__ == '__main__':
    main()
