import pandas as pd

# ── 配置区（按实际情况修改这里）──────────────────────────────
EXCEL_FILE  = "staff.xlsx"          # Excel 文件路径
SHEET_NAME  = 0                     # 第一个 Sheet；若有多个可改成 sheet 名字符串
HEADER_ROW  = 0                     # 标题行在 Excel 的第几行（0 = 第1行）
DATA_START  = 3                     # 真实数据从第几行开始（0-indexed，3 = Excel第4行）
                                    # 如果数据从第5行开始，改成 4

# 只抓取这几列；列名以 Excel 标题行为准，有偏差直接改这里
COLUMNS_NEEDED = ["staffID", "team-leader", "department"]
# ──────────────────────────────────────────────────────────────


def load_staff(filepath=EXCEL_FILE):
    # 读取整张表，强制 staffID 为字符串（防止数字 ID 被转成 float）
    df = pd.read_excel(
        filepath,
        sheet_name=SHEET_NAME,
        header=HEADER_ROW,
        skiprows=range(1, DATA_START),   # 跳过标题行之后、数据行之前的空行
        dtype={"staffID": str},
    )

    # 去掉整行全空的行（Excel 中间偶尔会有空行）
    df.dropna(how="all", inplace=True)

    # 只保留需要的列（列名对不上时会报 KeyError，改 COLUMNS_NEEDED 即可）
    df = df[COLUMNS_NEEDED]

    # 转成字典列表
    records = df.to_dict(orient="records")
    return records


if __name__ == "__main__":
    staff_list = load_staff()

    # 预览前 5 条
    for row in staff_list[:5]:
        print(row)

    print(f"\n共读取 {len(staff_list)} 条记录")

    # ── 可选：以 staffID 为 key 建立快速查找字典 ──
    staff_dict = {row["staffID"]: row for row in staff_list}
    # 使用示例：
    # info = staff_dict.get("XCS00061")
    # print(info)  # {'staffID': 'XCS00061', 'team-leader': 'Jim', 'department': 'market1'}
