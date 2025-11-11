import pandas as pd


def load_project(uploaded_file):
    if uploaded_file.name.endswith("csv"):
        blueprints_df = pd.read_csv(uploaded_file, delimiter=";", encoding="cp1250", decimal=",")
        manual_df = pd.DataFrame([], columns=["Systém", "Číslo", "Název", "Typ", "Součet", "--", "PN"])
        header_df = pd.DataFrame([{
            "zakázka:": "",
            "stavba:": "",
            "č. zakázky:": "",
        }])
    else:  # uploaded_file.name.endswith("xlsx")
        blueprints_df = pd.read_excel(uploaded_file, sheet_name="Data z výkresu")
        manual_df = pd.read_excel(uploaded_file, sheet_name="Data doplněná")
        try:
            header_df = pd.read_excel(uploaded_file, sheet_name="Hlavička")
        except:
            header_df = pd.DataFrame([{
                "zakázka:": "",
                "stavba:": "",
                "č. zakázky:": "",
            }])
    return blueprints_df, manual_df, header_df


def normalize_df(blueprints_df, manual_df):
    df = pd.concat([blueprints_df, manual_df])
    column_names = {
        'Systém': 'system',
        'Číslo': 'position',
        'PN': 'pn',
        'Název': 'name',
        'Typ': 'spec',
        'Insulation': 'insulation_mm',
        'izolace': 'insulation_manual_mm',
        'Vzduchovody, kusů': 'duct_count',
        'Průměr': 'diameter_mm',
        'Délka': 'length_mm',
        'Šířka': 'width_mm',
        'Výška': 'height_mm',
        'Plocha': 'surface_m2',
        'Součet': 'quantity',
        '--': 'unit',
    }
    drop_columns = [c for c in df.columns if c not in column_names]
    df = df.drop(columns=drop_columns).rename(columns=column_names)
    df['insulation_mm'] = df['insulation_mm'].fillna(df.pop('insulation_manual_mm')).fillna(0)
    return df