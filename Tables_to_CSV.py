from docx import Document
import pandas as pd
import re
from rapidfuzz import fuzz
from pathlib import Path
import Lot_extraction as dc

# use the program after extracting relevant tables with document.py to get the data in the relevant tables and combine to a csv file
# synonym dictionary, exclude terms dictionary
DESCRIPTIONS_SYN = [
    "Ամբողջական անվանումը",
    "Անվանում",
    "Ապրանքի անվանում",
    "Չափաբաժնի անվանումը"

]

LONG_DESCRIPTIONS_SYN = [
    "Տեխնիկական բնութագիրը",
    "Տեխնիկական բնութագիր",
    "Տեխնիկական բնութագիրը",
    "Տեխնիկական բնութագիր",
    "Տեխնիկական բնութագրեր",
    "բնութագիր",
    "բնութագր"
]

UNITS_SYN = [
    "Չափման միավորը",
    "Չ/Մ",
    "Չափի միավորը",
    "չափման միավոր"

]

PRICE_SYN = [
    "Գնման գին",
    "Ընդհանուր գինը/ՀՀ դրամ",
    "Ընդհանուր գինը",
    "ընդհանուր քանակը"
]

QUANTITY_SYN = [
    "Ընդհանուր քանակը",
    "ընդհանուր քանակը",
    "ընդհանուր քանակը"
    
    
]

EXCLUDE_TERMS = [
    "Կից ներկայացվում է",
    "Կից ներկայացված է",
    "Բնութագիրը կցվում է",
    "տես կից ֆայլը"
    
]

FUZZ_THRESHOLD = 85

def convert_excel(path):
    xls = pd.ExcelFile(path)
    dfs = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
        df = df.fillna("")
        dfs.append(df)

    return dfs

from pathlib import Path

def load_tables(path):
    ext = Path(path).suffix.lower()

    if ext == ".csv":
        return [pd.read_csv(path, encoding="utf-8-sig")]

    if ext in [".xlsx", ".xls"]:
        return convert_excel(path)

    raise ValueError("Unsupported file")

from pathlib import Path
import pandas as pd

def load_mixed_tables(paths):
    dfs = []

    for path in paths:
        ext = Path(path).suffix.lower()

        if ext == ".csv":
            df = pd.read_csv(path, encoding="utf-8-sig", dtype=str).fillna("")
            dfs.append(df)

        elif ext in [".xlsx", ".xls"]:
            xls = pd.ExcelFile(path)
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, dtype=str).fillna("")
                dfs.append(df)

        else:
            print(f"Skipping unsupported file: {path}")

    return dfs

def normalize(text):
    if not text:
        return ""

    text = text.lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^\w\s]", "", text)
    return text.strip()

def fuzzy_matches(text, synonyms, threshold=FUZZ_THRESHOLD):
    text = normalize(text)
    return any(fuzz.partial_ratio(text, normalize(s)) >= threshold for s in synonyms)

def fuzzy_remove_excluded_series(col, exclude_terms, threshold):
    if not exclude_terms:
        return col

    norm_excludes = [normalize(t) for t in exclude_terms]

    def clean_cell(x):
        t = normalize(x)

        for e in norm_excludes:
            if fuzz.partial_ratio(t, e) >= threshold:
                return ""   # remove whole cell (recommended)
        return x

    return col.apply(clean_cell)

def convert_df(csv_files):
    dfs = []
    for csv in csv_files:
        df = pd.read_csv(csv, encoding="utf-8-sig")
        dfs.append(df)
    
    return dfs
    
def find_header_row(df, synonyms, search_rows=6):
    for i in range(min(search_rows, len(df))):
        row = df.loc[i].astype(str)
        mask = row.map(lambda x: fuzzy_matches(x, synonyms))

        if mask.any():
            return i

    return None

def get_CPV(dfs):
    cpv_series = []

    for df in dfs:
        if df.empty:
            continue

        # --- case 1: CPV is real column name ---
        cpv_cols = df.columns[df.columns.astype(str).str.contains("CPV", case=False, na=False)]

        if len(cpv_cols) > 0:
            col = df[cpv_cols[0]].reset_index(drop=True).astype(str)
            col.name = "ԳՄԱ (CPV)"
            return col

        # --- case 2: header row is inside data ---
        first_row = df.iloc[0].astype(str)

        mask = first_row.str.contains("CPV", case=False, na=False)

        if mask.any():
            cpv_pos = mask.values.argmax() 

            cpv_col = df.iloc[1:, cpv_pos].reset_index(drop=True).astype(str)
          # skip header row
            cpv_col.name = "ԳՄԱ (CPV)"
            return cpv_col

    return pd.Series(name="ԳՄԱ (CPV)", dtype=str)

def get_descriptions(dfs, synonyms=None):
    if synonyms is None:
        synonyms = DESCRIPTIONS_SYN
    for df in dfs:
        if df.empty:
            continue
        mask = df.columns.astype(str).map(lambda x: fuzzy_matches(x, synonyms))
        if mask.any():
            col = df.loc[:, mask].iloc[:, 0]
            col.name = "Ամբողջական անվանումը"
            return col
        first_row = df.iloc[0].astype(str)
        mask = first_row.map(lambda x: fuzzy_matches(x, synonyms))

        if mask.any():
            idx = mask.values.argmax()  
            col = df.iloc[1:, idx].reset_index(drop=True).astype(str)
            col.name = "Ամբողջական անվանումը"
            return col

    return pd.Series(name="Ամբողջական անվանումը")

def get_long_descriptions(dfs, synonyms=None, exclude_terms=None):
    if synonyms is None:
        synonyms = LONG_DESCRIPTIONS_SYN
    
    if exclude_terms is None:
        exclude_terms = EXCLUDE_TERMS

    for df in dfs:
        if df.empty:
            continue

        mask = df.columns.astype(str).map(lambda x: fuzzy_matches(x, synonyms, FUZZ_THRESHOLD))
        
        if mask.any():
            col = df.loc[:, mask].iloc[:, 0]

        else:
            header_row = find_header_row(df, synonyms)

            if header_row is None:
                continue

            row = df.loc[header_row].astype(str)
            mask = row.map(lambda x: fuzzy_matches(x, synonyms, FUZZ_THRESHOLD))

            if not mask.any():
                continue

            col_label = mask.values.argmax()  
            col = df.iloc[header_row + 1:, col_label].reset_index(drop=True).astype(str)
            col = col.reset_index(drop=True)
    
        col = fuzzy_remove_excluded_series(col.astype(str), exclude_terms, threshold=FUZZ_THRESHOLD).str.strip()

        if col.replace("", pd.NA).dropna().empty:
            continue

        col.name = "Տեխնիկական բնութագիրը"
        return col


    return pd.Series(name="Տեխնիկական բնութագիրը")


def get_units(dfs, synonyms=None):
    if synonyms is None:
        synonyms = UNITS_SYN

    for df in dfs:
        if df.empty:
            continue
        mask = df.columns.astype(str).map(lambda x: fuzzy_matches(x, synonyms))
        if mask.any():
            col = df.loc[:, mask].iloc[:, 0]
            col.name = "Չափման միավորը"
            return col
        first_row = df.iloc[0].astype(str)
        mask = first_row.map(lambda x: fuzzy_matches(x, synonyms))

        if mask.any():
            idx = mask.values.argmax()  
            col = df.loc[1:, idx].reset_index(drop=True).astype(str)
            col.name = "Չափման միավորը"
            return col

    return pd.Series(name="Չափման միավորը")


def get_price(dfs, synonyms=None):
    if synonyms is None:
        synonyms = PRICE_SYN

    for df in dfs:
        if df.empty:
            continue
        mask = df.columns.astype(str).map(lambda x: fuzzy_matches(x, synonyms))
        if mask.any():
            col = df.loc[:, mask].iloc[:, 0]
            col.name = "Գնման գին"
            return col
        first_row = df.iloc[0].astype(str)
        mask = first_row.map(lambda x: fuzzy_matches(x, synonyms))

        if mask.any():
            idx = mask.values.argmax() 
            col = df.iloc[1:, idx].reset_index(drop=True).astype(str)
            col.name = "Գնման գին"
            return col

    return pd.Series(name="Գնման գին")


def get_quantity(dfs, synonyms=None):
    if synonyms is None:
        synonyms = QUANTITY_SYN

    for df in dfs:
        if df.empty:
            continue
        mask = df.columns.astype(str).map(lambda x: fuzzy_matches(x, synonyms))
        if mask.any():
            col = df.loc[:, mask].iloc[:, 0]
            col.name = "Ընդհանուր քանակը"
            return col
        first_row = df.iloc[0].astype(str)
        mask = first_row.map(lambda x: fuzzy_matches(x, synonyms))

        if mask.any():
            idx = mask.values.argmax() 
            col = df.iloc[1:, idx].reset_index(drop=True).astype(str)
            col.name = "Ընդհանուր քանակը"
            return col

    return pd.Series(name="Ընդհանուր քանակը")

# def get_unit_price(dfs):
#     df_price = get_price(dfs)
#     df_quantity = get_quantity(dfs)

#     # If df_price is a Series
#     unit_price = pd.Series(df_price / df_quantity.replace(0, pd.NA), name="միավոր գինը (հաշվարկված)")
#     return unit_price
def get_unit_price(dfs):
    price = get_price(dfs)
    qty = get_quantity(dfs)

    price_num = pd.to_numeric(price, errors="coerce")
    qty_num = pd.to_numeric(qty, errors="coerce")

    unit_price = price_num / qty_num.replace(0, pd.NA)
    unit_price.name = "Միավոր գինը (հաշվարկված)"

    return unit_price


def extract_data(dfs, name="combined_data"):
     df_CPV = get_CPV(dfs)
     df_desc = get_descriptions(dfs)
     df_long_desc = get_long_descriptions(dfs)
     df_units = get_units(dfs)
     df_quantity = get_quantity(dfs)
     df_price = get_price(dfs)
     df_unit_price = get_unit_price(dfs)
     df_lot_nums = pd.Series(range(1, len(df_CPV) + 1), name="Չափաբաժինների համարներով")

     df = pd.concat([df_lot_nums, df_CPV, df_desc, df_long_desc, df_units, df_quantity,
                     df_price, df_unit_price], axis=1)
     df.to_csv(f"{name}.csv", index=False, encoding="utf-8-sig")
     return df

# word_doc_path = r"C:\Users\Jinny\Downloads\Metropoliten 26-23 hraver.docx"
# word_doc = Document(word_doc_path)
# name = dc.safe_filename(dc.get_name(word_doc))
# files = [
#   # r"C:\Users\Jinny\Desktop\vibe-coder\lot_1_i_texnikakan_bnutagir_1.csv",
#     r"C:\Users\Jinny\Desktop\vibe-coder\ԿԱՐԵՆ ԴԵՄԻՐՃՅԱՆԻ ԱՆՎԱՆ ԵՐԵՎԱՆԻ ՄԵՏՐՈՊՈԼԻՏԵՆ ՓԲԸ_0.csv",
#     r"C:\Users\Jinny\Desktop\vibe-coder\ԿԱՐԵՆ ԴԵՄԻՐՃՅԱՆԻ ԱՆՎԱՆ ԵՐԵՎԱՆԻ ՄԵՏՐՈՊՈԼԻՏԵՆ ՓԲԸ_1.csv",
#     r"C:\Users\Jinny\Desktop\vibe-coder\ԿԱՐԵՆ ԴԵՄԻՐՃՅԱՆԻ ԱՆՎԱՆ ԵՐԵՎԱՆԻ ՄԵՏՐՈՊՈԼԻՏԵՆ ՓԲԸ_2.csv"
# ]

# df = load_mixed_tables(files)
# print(get_CPV(df))
# print(get_descriptions(df))
#print(get_long_descriptions(df))
# print(get_units(df))
# # print(get_price(df))
# print(get_quantity(df))
# print(get_unit_price(df))
# extract_data(df, name)