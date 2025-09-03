import io
import re

import numpy as np
import pandas as pd


def natural_keys(text):
    return tuple((int(c) if c.isdigit() else c) for c in re.split(r'(\d+)', text))


def safe_float(f):
    if isinstance(f, float):
        return f
    try:
        return float(f.replace(" ", ""))
    except:
        return np.nan


def parse_spec(spec):
    def parse_attribute(attribute):
        k, v = attribute.split("=")
        for k_ in k.split(","):
            yield k_, float(v.rstrip("°"))

    return {
        k: v
        for attribute in spec.split(", ")
        for k, v in parse_attribute(attribute)
    }

def compute_insulation(row: pd.Series) -> float:
    if not(row.insulation_thickness > 0):
        return pd.Series()
    
    thickness = int(row.insulation_thickness)
    try:
        if row.element == "Roura":
            width = np.pi * (row.diameter + 2 * thickness)
            length = row.length
        elif row.element == "Koleno":
            spec = parse_spec(row.spec)
            if "D" in spec:
                width = np.pi * spec["D"]
                length = 2 * np.pi * (spec["R"] + spec["D"] / 2 + thickness) * spec["a"] / 360
            else:
                length = (
                    spec.get("E", 0)
                    + spec.get("F", 0)
                    + (spec["R"] + spec["A"] + row.insulation_thickness) * 2 * np.pi * spec["a"] / 360
                )
                width = (
                    2 * (spec["A"] + 2 * row.insulation_thickness)
                    + 2 * (spec["B"] + 2 * row.insulation_thickness)
                )
        elif row.element == "Potrubí":
            width, height, length = map(int, row.spec.split(" x "))
            width = 2 * width + 2 * height + 8 * row.insulation_thickness
        elif row.element == "Redukce":
            spec = parse_spec(row.spec)
            width = np.pi * max(spec["D"], spec["D2"])
            length = spec["L"]
        elif row.element == "tlumič hluku":
            D, L, AI = map(int, row.spec.split(" ")[-1].split("/"))
            width = np.pi * (D + 2 * AI + 2 * thickness)
            length = L
        else:
            return pd.Series()
    
        return pd.Series(
            {
                f"insulation_width": width,
                f"insulation_height": length,
                f"insulation_area": width * length / 1000000,
            },
        )
    except:
        return {"insulation_area": -1}


def summarize_element(df: pd.DataFrame) -> pd.Series:
    return pd.Series(
        {
            "quantity": df.quantity.sum(),
            "uom": df.uom.unique()[-1],
            # "unit_price": df.unit_price.unique()[-1],
            # "price": df.price.sum(),
            "PN": df.PN.unique()[-1],
        },
    )


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    column_names = {
        "Systém": "system",
        "Číslo": "position",
        "PN": "PN",
        "Název": "element",
        "Typ": "spec",
        "Insulation": "insulation_thickness",
        "Plocha": "element_surface",  # m2?
        "Průměr": "diameter",
        "Délka": "length",
        "Šířka": "width",
        "Výška": "height",
        "Součet": "quantity",
        "--": "uom",
    }
    
    df = df.copy().drop(columns=["č."])
    df = df.rename(columns=column_names)
    # normalize
    df = df[df.system.notna()]
    df["spec"] = df.spec.fillna(df.element)
    df["insulation_thickness"] = df.insulation_thickness.fillna(df.pop("izolace"))
    # drop empty columns
    allna = df.isna().all(axis=0)
    df = df.drop(columns=allna[allna].index.tolist())
    # ducts
    pieces = df.pop("Vzduchovody, kusů").fillna(1)
    df["quantity"] = df.quantity.where(pieces <= 1, df.length / 1000)
    ## rectangular ducts
    df["spec"] = df.apply(lambda row: f"{row.spec} x {int(row.length)}" if ("Potrubí" in row.element) else row.spec, axis=1)              
    df["quantity"] = df.quantity.where(~df.element.str.contains("Potrubí"), 1)
    df["uom"] = df.uom.where(~df.element.str.contains("Potrubí"), "ks")
    df = df.reindex(df.index.repeat(pieces)).reset_index(drop=True)
    return df


def insulate_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.join(df.apply(compute_insulation, axis=1).fillna(0))
    
    if "insulation_area" in df.columns:
        # estimate insulation when formula missing
        idf = df[df.insulation_area > 0][["insulation_thickness", "element_surface", "insulation_area"]]
        insulation_area_ratio = {
            thick: (group.insulation_area / group.element_surface).quantile(0.75)
            for thick, group in idf.groupby("insulation_thickness")
        }
        insulation_mask = df.insulation_area == -1
        df.insulation_area[insulation_mask] = (
            df.element_surface[insulation_mask] * df.insulation_thickness.map(insulation_area_ratio)
        )

        insulation_df = df[["system", "insulation_thickness", "insulation_area", "quantity", "uom"]].copy()
        insulation_df["insulation_area"] = insulation_df.quantity.where(insulation_df.uom == "ks", 1) * insulation_df.insulation_area
        insulation_df = insulation_df.groupby(["system", "insulation_thickness"]).insulation_area.sum().to_frame().reset_index()
        insulation_df = insulation_df.rename(columns={"insulation_thickness": "spec", "insulation_area": "quantity"})
        insulation_df = insulation_df[insulation_df.quantity != 0]
        insulation_df.spec = insulation_df.spec.apply(lambda s: f"tl={s:.0f}mm")
        insulation_df[["element", "uom", "position"]] = ["Izolace", "m2", "i"]
    else:
        insulation_df = pd.DataFrame()
    
    df = pd.concat([df, insulation_df])
    return df


def denormalize_potrubi(row: pd.Series) -> pd.Series:
    if ("Potrubí" in row.element) and ((row.width, row.height) in [(200, 50), (160, 40)]):
        return pd.Series({**row.to_dict(), "spec": " x ".join(row.spec.split(" x ")[:-1]), "uom": "m", "quantity": row.length / 1000})
    return row


def summarize_df(blueprints_df: pd.DataFrame, manual_df: pd.DataFrame, df: pd.DataFrame, header: dict[str, str] = {}) -> pd.DataFrame:
    shopping_list_df = df.apply(denormalize_potrubi, axis=1).groupby(["element", "spec"]).apply(summarize_element, include_groups=False)
    shopping_list_df["quantity"] = shopping_list_df.quantity.round(decimals=1)

    inventory_df = df.copy()
    inventory_df["quantity"] = inventory_df.quantity.round(decimals=2)

    rename = {
        "position": "Č. pozice",
        "element": "Název",
        "spec": "Typ",
        "quantity": "Množství",
        "uom": "Jednotka",
        "unit_price": "Cena za jednotku",
        "price": "Cena celkem",
        "insulation_thickness": "Izolace (mm)",
    }

    inventory_order = ["position", "element", "spec", "quantity", "uom", "PN"]
    shopping_order = ["element", "spec", "quantity", "uom", "PN"]
    
    bio = io.BytesIO()
    writer = pd.ExcelWriter(bio, engine='xlsxwriter')
    blueprints_df.to_excel(writer, "Data z výkresu", index=False)
    manual_df.to_excel(writer, "Data doplněná", index=False)
    (
        df
        .loc[df.position.sort_values().index]
        .to_excel(writer, sheet_name='Data', index=False)
    )
    (
        inventory_df
        .groupby(["system", "position", "element", "spec"], dropna=False)
        .apply(lambda df: pd.Series({**df.iloc[0].to_dict(), "quantity": df.quantity.sum()}))
        .reset_index(drop=True)
        .sort_values("position", key=lambda col: col.apply(natural_keys))
        .groupby("system")
        .apply(
            lambda df: pd.concat([pd.DataFrame([{"element": df.system.unique().item()}]), df]),
            include_groups=True,
        )
        [inventory_order]
        .rename(columns=rename)
        .to_excel(writer, sheet_name='Souhrn za systém', startrow=len(header) + 2, index=False)
    )

    shopping_list_summary = (
        shopping_list_df
        .sort_index(key=lambda x: x.str.normalize("NFKD").str.lower())
        .reset_index()
        [shopping_order]
    )
    (
        shopping_list_summary
        .rename(columns=rename, index=rename)
        .to_excel(writer, sheet_name='Souhrn celkový', startrow=len(header) + 2, index=False)
    )

    workbook = writer.book
    bold_border_style = 2
    format_left = workbook.add_format({'left': bold_border_style})
    format_right = workbook.add_format({'right': bold_border_style})
    format_top_left = workbook.add_format({'top': bold_border_style, 'left': bold_border_style})
    format_top_right = workbook.add_format({'top': bold_border_style, 'right': bold_border_style})
    format_bottom_left = workbook.add_format({'bottom': bold_border_style, 'left': bold_border_style})
    format_bottom_right = workbook.add_format({'bottom': bold_border_style, 'right': bold_border_style})

    worksheet = writer.sheets["Souhrn za systém"]
    worksheet.set_column(0, 0, 8, None)
    worksheet.set_column(1, 1, 40, None)
    worksheet.set_column(2, 2, 35, None)
    worksheet.set_column(3, 3, 8, None)
    worksheet.set_column(4, 4, 8, None)
    worksheet.set_column(5, 5, 10, None)
    for row_i, (k, v) in enumerate(header.items()):
        worksheet.write(row_i + 1, 1, k, {0: format_top_left, len(header) - 1: format_bottom_left}.get(row_i, format_left))
        worksheet.write(row_i + 1, 2, v, {0: format_top_right, len(header) - 1: format_bottom_right}.get(row_i, format_right))
    worksheet.set_header("&CTECHNICKÁ ZPRÁVA - VÝPIS MATERIÁLU - Souhrn za systém\nEPD Rychnov –  Martin Jindrák, Březová 803, 468 02 Rychnov u Jablonce nad Nisou, E: martin.jindrak@pasivprojekt.cz, T: 778044062")
    worksheet.set_footer("&CStrana &P z &N")
    worksheet.set_print_scale(73)

    worksheet = writer.sheets["Souhrn celkový"]
    worksheet.set_column(0, 0, 40, None)
    worksheet.set_column(1, 1, 35, None)
    worksheet.set_column(2, 2, 8, None)
    worksheet.set_column(3, 3, 8, None)
    worksheet.set_column(4, 4, 10, None)
    worksheet.set_column(5, 5, 8, None)
    worksheet.set_column(6, 6, 8, None)
    for row_i, (k, v) in enumerate(header.items()):
        worksheet.write(row_i + 1, 0, k, {0: format_top_left, len(header) - 1: format_bottom_left}.get(row_i, format_left))
        worksheet.write(row_i + 1, 1, v, {0: format_top_right, len(header) - 1: format_bottom_right}.get(row_i, format_right))

    format_top = workbook.add_format({"valign": "top"})
    begins = np.arange(len(shopping_list_summary))[~shopping_list_summary.element.duplicated()]
    sizes = np.r_[begins[1:], len(shopping_list_summary)] - begins
    for begin, size in zip(begins, sizes):
        worksheet.merge_range(
            len(header) + 3 + begin, 0,
            len(header) + 3 + begin + size - 1, 0,
            shopping_list_summary.element.iloc[begin],
            format_top
        )
    worksheet.set_header("&CTECHNICKÁ ZPRÁVA - VÝPIS MATERIÁLU - Souhrn celkový\nEPD Rychnov –  Martin Jindrák, Březová 803, 468 02 Rychnov u Jablonce nad Nisou, E: martin.jindrak@pasivprojekt.cz, T: 778044062")
    worksheet.set_footer("&CStrana &P z &N")
    worksheet.set_print_scale(75)

    writer.close()
    return bio.getvalue()