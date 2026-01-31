import pandas as pd
import os
import re
import time

MARKER_PATTERN = re.compile(r"[]")

def split_key_and_suffix(cell):
    if not isinstance(cell, str):
        return "", ""

    key = cell.split("=", 1)[0]
    match = MARKER_PATTERN.search(cell)
    suffix = cell[match.start():] if match else ""

    return key, suffix


def export_lang_files():
    start = time.time()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_dir, "item info auto.xlsx")

    # ----------------------
    # LOAD AUTOMATION
    # ----------------------
    df_auto = pd.read_excel(
        excel_path,
        sheet_name="automation",
        dtype=str,
        usecols=[0, 1, 2]
    )
    df_auto.columns = ["Key", "Lang", "Text"]

    # ----------------------
    # LOAD LANGUAGE RULES
    # ----------------------
    df_lang = pd.read_excel(
        excel_path,
        sheet_name="languages",
        dtype=str
    )
    df_lang.set_index(df_lang.columns[0], inplace=True)

    # ----------------------
    # LOAD TEMPLATE
    # ----------------------
    df_temp = pd.read_excel(
        excel_path,
        sheet_name="new generation",
        dtype=str
    )

    a_col = df_temp.columns[0]
    df_temp[["Key", "Suffix"]] = df_temp[a_col].apply(
        lambda x: pd.Series(split_key_and_suffix(x))
    )

    output_dir = os.path.join(base_dir, "Exported_Lang_Files")
    os.makedirs(output_dir, exist_ok=True)

    # ----------------------
    # EXPORT
    # ----------------------
    for lang in df_auto["Lang"].unique():
        print(f"Exporting {lang}")

        # Language rules
        decimal_comma = df_lang.loc[lang, "decimal"] == "yes"
        unit_symbol = df_lang.loc[lang, "unit"]

        lookup = dict(
            zip(
                df_auto[df_auto["Lang"] == lang]["Key"],
                df_auto[df_auto["Lang"] == lang]["Text"]
            )
        )

        base = df_temp["Key"].map(lookup).fillna("")

        suffix = df_temp["Suffix"]

        if decimal_comma:
            suffix = suffix.str.replace(".", ",", regex=False)

        suffix = suffix.str.replace("s", unit_symbol, regex=False)

        final = base + " " + suffix

        final.to_csv(
            os.path.join(output_dir, f"{lang}.lang"),
            index=False,
            header=False,
            encoding="utf-8"
        )

    print(f"DONE in {time.time() - start:.2f}s")


if __name__ == "__main__":
    export_lang_files()
