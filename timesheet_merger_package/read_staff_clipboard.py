import io
import csv
import pandas as pd
import pyperclip

# ── 配置区 ────────────────────────────────────────────────────
# 按列位置抓取，0-indexed（A=0, I=8, J=9, K=10, L=11）
COL_POSITIONS  = [0, 8, 9, 10, 11]
COL_NAMES      = ["SID", "Function", "Workstream", "Subteam", "Sub-Subteam"]
ID_COL         = "SID"
# ──────────────────────────────────────────────────────────────


def load_staff_from_clipboard():
    raw = pyperclip.paste()

    df = pd.read_csv(
        io.StringIO(raw),
        sep="\t",
        engine="python",
        quoting=csv.QUOTE_NONE,
        on_bad_lines="skip",
        header=None,        # 不把第一行当标题
        dtype=str,
    )

    print(f"读到 {len(df)} 行 × {len(df.columns)} 列")

    # 按列位置取出需要的列，重命名
    df = df.iloc[:, COL_POSITIONS]
    df.columns = COL_NAMES

    # 去掉整行全空的行
    df.dropna(how="all", inplace=True)

    # SID 去除首尾空格，过滤掉标题行本身（防止复制时带进来）
    df[ID_COL] = df[ID_COL].str.strip()
    df = df[df[ID_COL] != "SID"]   # 去掉可能混入的标题行
    df = df[df[ID_COL].notna() & (df[ID_COL] != "")]

    return df.to_dict(orient="records")


if __name__ == "__main__":
    print("请确认已在 Excel 中复制好数据（含或不含标题行均可），然后按回车继续...")
    input()

    staff_list = load_staff_from_clipboard()

    print(f"\n共读取 {len(staff_list)} 条记录，前 5 条预览：\n")
    for row in staff_list[:5]:
        print(row)

    staff_dict = {row["SID"]: row for row in staff_list}
    print(f"\nstaff_dict 已构建，共 {len(staff_dict)} 个 key")
