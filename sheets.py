from __future__ import annotations

import codecs
import io

import pandas as pd


def read_zwcad_dump(bio: io.BytesIO, n_columns: int | None = None) -> pd.DataFrame:
    with codecs.getreader("cp1250")(bio) as f:
        lines = [line.rstrip() for line in f]
        df = pd.DataFrame(
            [line.split(";")[:n_columns] for line in lines[1:]],
            columns=lines[0].split(";")[:n_columns],
        )
    df = pd.read_csv(io.StringIO(df.to_csv(index=False)), decimal=",")

    column_names = {
        "Systém": "system",
        "Číslo": "reference",
        "Název": "element",
        "Typ": "desc", 
        "Insulation": "insulation_thickness",
        "Plocha": "element_surface",
        "Průměr": "diameter",
        "Šířka": "width",
        "Výška": "height",
        "Součet": "quantity",
        "--": "uom",
    }

    df = df.drop(columns=["č."])
    df = df.rename(columns=column_names)
    df["desc"] = df.desc.fillna(df.element)
    return df


def to_excel(df: pd.DataFrame) -> bytes:
    """
    Converts a pandas DataFrame to an Excel file in memory.

    Args:
        df (pd.DataFrame): The DataFrame to convert.

    Returns:
        bytes: The Excel file as bytes.
    """
    output = io.BytesIO()
    # Use the 'xlsxwriter' engine for better compatibility and performance
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data