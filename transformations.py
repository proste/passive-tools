from __future__ import annotations

import numpy as np
import pandas as pd


def dummy_transformation(df: pd.DataFrame) -> pd.DataFrame:
    return df


def parse_spec(desc: str) -> dict[str, float]:
    def parse_attribute(attribute):
        k, v = attribute.split("=")
        for k_ in k.split(","):
            yield k_, float(v.rstrip("°"))

    return {
        k: v
        for attribute in desc.split(", ")
        for k, v in parse_attribute(attribute)
    }


def compute_insulation(row: pd.Series) -> pd.Series:
    if not(row.insulation_thickness > 0):
        return pd.Series()

    thickness = int(row.insulation_thickness)
    if row.element == "Roura":
        width = np.pi * (row.diameter + 2 * thickness)
        length = row.quantity * 1000
    elif row.element == "Koleno":
        spec = parse_spec(row.desc)
        if "D" in spec:
            width = np.pi * spec["D"]
            length = 2 * np.pi * (spec["R"] + spec["D"] / 2 + thickness) * spec["a"] / 360
        else:
            length = (
                spec["E"]
                + spec["F"]
                + (spec["R"] + spec["A"] + row.insulation_thickness) * 2 * np.pi * spec["a"] / 360
            )
            width = (
                2 * (spec["A"] + 2 * row.insulation_thickness)
                + 2 * (spec["B"] + 2 * row.insulation_thickness)
            )
    elif row.element == "Potrubí":
        width, height = map(int, row.desc.split(" x "))
        width = 2 * width + 2 * height
        height += 2 * row.insulation_thickness
        length = row.quantity * 1000
    else:
        return pd.Series()

    return pd.Series(
        {
            "insulation_width": width,
            "insulation_height": length,
            "insulation_area": width * length / 1000000,
        },
    )


def insulation_transformation(df: pd.DataFrame) -> pd.DataFrame:
    return df.join(df.apply(compute_insulation, axis=1))