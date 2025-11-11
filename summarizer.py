import base64
import datetime as dt
import io
import re

import pandas as pd


def natural_keys(text):
    return tuple((int(c) if c.isdigit() else c) for c in re.split(r'(\d+)', text))


class Summarizer:
    HEADER_B64 = b'JkMmMTBURUNITklDS8OBIFpQUsOBVkEgLSBWw51QSVMgTUFURVJJw4FMVSAtIFNvdWhybiB6YSBzeXN0w6ltCkVQRCBSeWNobm92IOKAkyAgTWFydGluIEppbmRyw6FrLCBCxZllem92w6EgODAzLCA0NjggMDIgUnljaG5vdiB1IEphYmxvbmNlIG5hZCBOaXNvdSwgRTogbWFydGluLmppbmRyYWtAcGFzaXZwcm9qZWt0LmN6LCBUOiA3NzgwNDQwNjI='
    BOLD_LINE_WIDTH = 2
    FINE_LINE_WIDTH = 1
    COLUMN_NAMES = {
        "position": "Č. pozice",
        "name": "Název",
        "spec": "Typ",
        "quantity": "Množství",
        "unit": "Jednotka",
        "pn": "PN",
        "price": "Cena celkem",
        "unit_price": "Cena za jednotku",
    }
    
    def __init__(self, header):
        self.bio = io.BytesIO()
        self.writer = pd.ExcelWriter(self.bio, engine='xlsxwriter')
        self.workbook = self.writer.book
        self.header = base64.urlsafe_b64decode(self.HEADER_B64).decode("utf8")
        self.header_table = {
            **header,
            "vypracoval:": "M. Jindráková",
            "dne:": dt.date.today().strftime("%-d/%-m/%Y"),
        }

        self.format_bold = self.workbook.add_format({'bold': True, 'bottom': self.BOLD_LINE_WIDTH})
        self.format_fine = self.workbook.add_format({'bottom': self.FINE_LINE_WIDTH})
        self.format_left = self.workbook.add_format({'left': self.BOLD_LINE_WIDTH})
        self.format_right = self.workbook.add_format({'right': self.BOLD_LINE_WIDTH})
        self.format_top_left = self.workbook.add_format({'top': self.BOLD_LINE_WIDTH, 'left': self.BOLD_LINE_WIDTH})
        self.format_top_right = self.workbook.add_format({'top': self.BOLD_LINE_WIDTH, 'right': self.BOLD_LINE_WIDTH})
        self.format_bottom_left = self.workbook.add_format({'bottom': self.BOLD_LINE_WIDTH, 'left': self.BOLD_LINE_WIDTH})
        self.format_bottom_right = self.workbook.add_format({'bottom': self.BOLD_LINE_WIDTH, 'right': self.BOLD_LINE_WIDTH})

        self.write_inputs(pd.DataFrame([header]), "Hlavička")

    def write_inputs(self, df, name):
        df.to_excel(self.writer, sheet_name=name, index=False)

    def write_inventory(self, elements_df):
        # TODO rewrite to elements, implement addition of same elements and sorting of elements
        # apply groupby_sorted to sum elements, evade whole pandas magic
        worksheet, row_i = self._get_worksheet("Souhrn za systém", 1)
        inventory_order = ["position", "name", "spec", "quantity", "unit", "pn", "unit_price", "price"]

        # write column names
        for column_i, column_name in enumerate(inventory_order):
            worksheet.write(row_i, column_i, self.COLUMN_NAMES[column_name], self.format_bold)
        row_i += 1

        for system, system_df in (
            elements_df
            .groupby(["system", "position", "name", "spec", "unit", "pn"], dropna=False)
            .agg(dict(quantity=sum, price=sum))
            .reset_index()
            .sort_values("position", key=lambda col: col.apply(natural_keys))
            .groupby("system")
        ):
            row_i += 1
            for c_i, _ in enumerate(inventory_order):
                worksheet.write(row_i, c_i, system if c_i == 1 else "", self.format_bold)
            row_i += 1
            system_df = system_df.copy()
            system_df['unit_price'] = system_df.price / system_df.quantity
            system_df["quantity"] = system_df.quantity.round(decimals=2)
            system_df["price"] = system_df.price.round(decimals=2)
            system_df["unit_price"] = system_df.unit_price.round(decimals=2)
            for _, row in system_df[inventory_order].rename(columns=self.COLUMN_NAMES).iterrows():
                for col_i, v in enumerate(row.fillna("").values):
                    worksheet.write(row_i, col_i, v, self.format_fine)
                row_i += 1

        worksheet.set_column(0, 0, 8, None)
        worksheet.set_column(1, 1, 35, None)
        worksheet.set_column(2, 2, 40, None)
        worksheet.set_column(3, 3, 8, None)
        worksheet.set_column(4, 4, 8, None)
        worksheet.set_column(5, 5, 12, None)
        worksheet.print_area(0, 0, row_i, 5)
        worksheet.fit_to_pages(1, 0)

    def write_shopping_list(self, elements_df):
        worksheet, row_i = self._get_worksheet("Souhrn celkový", 0)
        shopping_order = ["name", "spec", "quantity", "unit", "pn"]

        # write column names
        for column_i, column_name in enumerate(shopping_order):
            worksheet.write(row_i, column_i, self.COLUMN_NAMES[column_name], self.format_bold)
        row_i += 1
        
        for name, name_df in (
            elements_df
            .groupby(["name", "spec", "pn", "unit"], dropna=False)
            .quantity.sum().to_frame()
            .sort_index(key=lambda x: x.str.normalize("NFKD").str.lower())
            .reset_index()
            .groupby("name")
        ):
            name_df = name_df.copy()
            name_df["quantity"] = name_df.quantity.round(decimals=1)
            for row_i_, (_, row) in enumerate(name_df[shopping_order].rename(columns=self.COLUMN_NAMES).iterrows()):
                for col_i, v in enumerate(row.fillna("").values):
                    if col_i == 0:
                        worksheet.write(row_i, col_i, name if row_i_ == 0 else "", self.format_fine if row_i_ == len(name_df) - 1 else None)
                    else:
                        worksheet.write(row_i, col_i, v, self.format_fine)
                row_i += 1

        worksheet.set_column(0, 0, 35, None)
        worksheet.set_column(1, 1, 40, None)
        worksheet.set_column(2, 2, 8, None)
        worksheet.set_column(3, 3, 8, None)
        worksheet.set_column(4, 4, 12, None)
        worksheet.print_area(0, 0, row_i, 4)
        worksheet.fit_to_pages(1, 0)

    def close(self):
        self.writer.close()
        return self.bio.getvalue()

    def _get_worksheet(self, name, header_offset):
        worksheet = self.workbook.add_worksheet(name)
        worksheet.set_header(self.header)
        worksheet.set_footer("&CStrana &P z &N")
        worksheet.hide_gridlines(2)
    
        row_i = 1
        header_len = len(self.header_table)
        for h_i, (k, v) in enumerate(self.header_table.items()):
            worksheet.write(
                row_i, header_offset, k,
                {0: self.format_top_left, header_len - 1: self.format_bottom_left}.get(h_i, self.format_left),
            )
            worksheet.write(
                row_i, header_offset + 1, v,
                {0: self.format_top_right, header_len - 1: self.format_bottom_right}.get(h_i, self.format_right),
            )
            row_i += 1
        row_i += 1
        
        return worksheet, row_i