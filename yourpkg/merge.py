import os
import pandas as pd

def _read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path)
    if ext in (".csv", ".txt"):
        for enc in ("utf-8-sig", "utf-8", "cp1251", "latin1"):
            try:
                return pd.read_csv(path, encoding=enc)
            except Exception:
                continue
        return pd.read_csv(path)  # last attempt
    try:
        return pd.read_excel(path)
    except Exception:
        return pd.read_csv(path)

def merge_two_files(file1_path, file2_path, mode="join", how="inner",
                    left_key=None, right_key=None,
                    out_path="/content/merged_result.xlsx"):
    df1 = _read_table(file1_path)
    df2 = _read_table(file2_path)
    if mode == "hconcat":
        merged = pd.concat([df1.reset_index(drop=True), df2.reset_index(drop=True)], axis=1)
    else:
        if not left_key or not right_key:
            raise ValueError("Укажите left_key и right_key при mode='join'")
        if str(left_key) not in map(str, df1.columns):
            raise KeyError(f"'{left_key}' нет в первом файле")
        if str(right_key) not in map(str, df2.columns):
            raise KeyError(f"'{right_key}' нет во втором файле")
        merged = pd.merge(df1, df2, how=how, left_on=str(left_key), right_on=str(right_key))
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    merged.to_excel(out_path, index=False)
    return out_path
