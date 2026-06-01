import io
import csv
import pandas as pd
import pyperclip

# ── 配置区 ────────────────────────────────────────────────────
ID_COL         = "SID"
COLUMNS_NEEDED = ["SID", "Function", "Workstream", "Subteam", "Sub-Subteam"]
# ──────────────────────────────────────────────────────────────


def load_staff_from_clipboard():
    # 读取 Windows 剪贴板原始文本
    raw = pyperclip.paste()

    # 用 StringIO 包装后交给 pandas 解析，QUOTE_NONE 避免引号干扰
    df = pd.read_csv(
        io.StringIO(raw),
        sep="\t",
        engine="python",
        quoting=csv.QUOTE_NONE,
        on_bad_lines="skip",
        dtype={ID_COL: str},
    )

    # 去掉整行全空的行
    df.dropna(how="all", inplace=True)

    # 只保留需要的列
    df = df[COLUMNS_NEEDED]

    # SID 去除首尾空格
    df[ID_COL] = df[ID_COL].str.strip()

    return df.to_dict(orient="records")


if __name__ == "__main__":
    print("请确认已在 Excel 中复制好数据（含标题行），然后按回车继续...")
    input()

    staff_list = load_staff_from_clipboard()

    print(f"\n共读取 {len(staff_list)} 条记录，前 5 条预览：\n")
    for row in staff_list[:5]:
        print(row)

    # 以 SID 为 key 的快速查找字典
    staff_dict = {row["SID"]: row for row in staff_list}
    print(f"\nstaff_dict 已构建，共 {len(staff_dict)} 个 key")
