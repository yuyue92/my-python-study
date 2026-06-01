import pandas as pd

# ── 配置区（按实际情况修改这里）──────────────────────────────
EXCEL_FILE  = "staff.xlsx"          # Excel 文件路径
SHEET_NAME  = 0                     # 第一个 Sheet；若有多个可改成 sheet 名字符串
HEADER_ROW  = 3                     # 标题行在 Excel 的第几行（0-indexed，3 = Excel第4行）
DATA_START  = 4                     # 真实数据从第几行开始（0-indexed，4 = Excel第5行）

# 只抓取这几列；列名以 Excel 第4行标题为准，有偏差直接改这里
ID_COL         = "SID"
COLUMNS_NEEDED = ["SID", "Function", "Workstream", "Subteam", "Sub-Subteam"]
# ──────────────────────────────────────────────────────────────


def load_staff(filepath=EXCEL_FILE):
    # 读取整张表，强制 SID 为字符串（防止数字 ID 被转成 float）
    df = pd.read_excel(
        filepath,
        sheet_name=SHEET_NAME,
        header=HEADER_ROW,
        skiprows=range(HEADER_ROW + 1, DATA_START),  # 跳过标题行和数据行之间的空行（本文件无，保留兼容）
        dtype={ID_COL: str},
    )

    # 去掉整行全空的行（Excel 中间偶尔会有空行）
    df.dropna(how="all", inplace=True)

    # 只保留需要的列（列名对不上时会报 KeyError，改 COLUMNS_NEEDED 即可）
    df = df[COLUMNS_NEEDED]

    # SID 去除首尾空格
    df[ID_COL] = df[ID_COL].str.strip()

    # 转成字典列表
    records = df.to_dict(orient="records")
    return records


if __name__ == "__main__":
    staff_list = load_staff()

    # 预览前 5 条
    for row in staff_list[:5]:
        print(row)

    print(f"\n共读取 {len(staff_list)} 条记录")

    # 以 SID 为 key 建立快速查找字典
    staff_dict = {row["SID"]: row for row in staff_list}
    # 使用示例：
    # info = staff_dict.get("XCS00061")
    # print(info)
    # → {'SID': 'XCS00061', 'Function': 'xxx', 'Workstream': 'xxx', 'Subteam': 'xxx', 'Sub-Subteam': 'xxx'}
