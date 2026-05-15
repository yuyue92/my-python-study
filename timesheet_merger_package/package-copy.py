"""
Timesheet Merger
================
从多个源文件（员工姓名.xlsx）提取考勤数据，
按格式类型分别汇总输出到 output_old.xlsx / output_new.xlsx。

使用方法：
  python timesheet_merger.py
  或指定目录：
  python timesheet_merger.py --src ./源文件目录 --out ./输出目录

格式自动识别：
  - new-format：源文件含名为 'Timesheet' 的 sheet，且表头有 'Employee' 列
  - old-format：其余所有源文件
"""

import argparse
import shutil
import warnings
from datetime import datetime, date
from pathlib import Path

import openpyxl
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)  # 忽略 openpyxl 扩展警告

# ─────────────────────────────────────────────
# 配置区（按需修改）
# ─────────────────────────────────────────────
DEFAULT_SRC_DIR = "./source_files"          # 源文件目录
DEFAULT_OUT_DIR = "./output"                # 输出目录
TEMPLATE_OLD    = "./templates/template-old-format.xlsx"
TEMPLATE_NEW    = "./templates/template-new-format.xlsx"

# 跳过这些 sheet（不是员工数据）
SKIP_SHEETS = {"readme", "lists"}

# old-format 模板：数据表头行号、数据起始行号、目标 sheet
OLD_HEADER_ROW  = 3
OLD_DATA_START  = 4
OLD_SHEET       = "FS"

# new-format 模板：数据表头行号、数据起始行号、目标 sheet
NEW_HEADER_ROW  = 10
NEW_DATA_START  = 11
NEW_SHEET       = "Timesheet"

# ─────────────────────────────────────────────
# Step 1：扫描源文件目录
# ─────────────────────────────────────────────
def discover_files(src_dir: Path) -> list[Path]:
    """返回目录下所有 .xlsx 文件路径（排除临时文件）"""
    files = [
        f for f in src_dir.glob("*.xlsx")
        if not f.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"在 {src_dir} 未找到任何 .xlsx 文件")
    print(f"[发现] 共找到 {len(files)} 个源文件")
    return sorted(files)


# ─────────────────────────────────────────────
# Step 2：识别文件格式
# ─────────────────────────────────────────────
def detect_format(wb: openpyxl.Workbook) -> str:
    """
    判断源文件格式：
    - 有名为 'Timesheet' 的 sheet，且第3行表头含 'Employee' → 'new'
    - 否则 → 'old'
    """
    sheet_names_lower = {s.lower(): s for s in wb.sheetnames}
    if "timesheet" in sheet_names_lower:
        ws = wb[sheet_names_lower["timesheet"]]
        # 第3行表头检查
        row3 = [str(c.value or "").strip() for c in ws[3]]
        if any("employee" in v.lower() for v in row3):
            return "new"
    return "old"


# ─────────────────────────────────────────────
# Step 3：从单个 sheet 提取数据
# ─────────────────────────────────────────────
def _excel_serial_to_date(val) -> str | None:
    """Excel 序列号或 datetime 转 yyyy-mm-dd 字符串"""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)):
        # Excel epoch: 1900-01-01 = 1（注意 Excel 的 1900-02-29 bug）
        try:
            from datetime import timedelta
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=int(val))).strftime("%Y-%m-%d")
        except Exception:
            return str(val)
    return str(val).strip() or None


def _find_header_row(ws) -> tuple[int, dict]:
    """
    自动在前6行里找含 'Date' 和 'Hours' 的表头行。
    返回 (行号, {列名小写: 列索引(1-based)})
    """
    for row_idx in range(1, 7):
        row_vals = [str(c.value or "").strip().lower() for c in ws[row_idx]]
        if "date" in row_vals and "hours" in row_vals:
            col_map = {v: i + 1 for i, v in enumerate(row_vals) if v}
            return row_idx, col_map
    raise ValueError(f"在 sheet '{ws.title}' 的前6行未找到有效表头")


def _get_col(col_map: dict, *candidates) -> int | None:
    """从多个候选列名里找第一个存在的列索引"""
    for name in candidates:
        if name in col_map:
            return col_map[name]
    return None


def extract_sheet(ws, staff_id: str, name: str) -> list[dict]:
    """
    从一个员工 sheet 提取所有有效行。
    有效行：Date 和 Hours 均不为空。
    """
    _, col_map = _find_header_row(ws)
    header_row, _ = _find_header_row(ws)
    data_start = header_row + 1

    # 列索引映射（兼容两种格式的不同列名）
    c_date    = _get_col(col_map, "date")
    c_role    = _get_col(col_map, "role / type", "role/type", "role / type ")
    c_module  = _get_col(col_map, "module")
    c_task    = _get_col(col_map, "task nature")
    c_desc    = _get_col(col_map, "description / notes", "description / notes (optional)",
                          "description/notes", "description / notes ")
    c_refid   = _get_col(col_map, "ref id (no need)", "ref id", "ref/ticket #", "ref id (no need) ")
    c_related = _get_col(col_map, "related number (auto) (no need)", "related number (auto)",
                          "related number (auto) (no need) ")
    c_hours   = _get_col(col_map, "hours")
    c_status  = _get_col(col_map, "status", "status (optional)")

    rows = []
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        def g(col_idx):
            if col_idx is None:
                return None
            v = row[col_idx - 1] if col_idx <= len(row) else None
            if isinstance(v, str):
                v = v.strip()
                return None if v in ("", " ") else v
            return v

        raw_date  = g(c_date)
        raw_hours = g(c_hours)

        # 过滤：Date 和 Hours 都必须有值
        if raw_date is None or raw_hours is None:
            continue
        try:
            hours_val = float(raw_hours)
        except (TypeError, ValueError):
            continue

        rows.append({
            "Date":           _excel_serial_to_date(raw_date),
            "Role / Type":    g(c_role),
            "Module":         g(c_module),
            "Task Nature":    g(c_task),
            "Description":    g(c_desc),
            "Ref ID":         g(c_refid),
            "Related Number": g(c_related),
            "Hours":          hours_val,
            "Status":         g(c_status),
            "Staff ID":       staff_id,
            "Name":           name,
        })
    return rows


# ─────────────────────────────────────────────
# Step 4：解析一个源文件 → 标准化记录列表
# ─────────────────────────────────────────────
def parse_source(path: Path) -> tuple[str, list[dict]]:
    """
    读取一个源文件，返回 (格式类型, 记录列表)。
    格式类型: 'old' | 'new'
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    fmt = detect_format(wb)

    all_rows = []
    for sname in wb.sheetnames:
        if sname.lower() in SKIP_SHEETS:
            continue

        ws = wb[sname]
        if ws.sheet_state != "visible":
            print(f"  [跳过] {path.name} → sheet '{sname}'：隐藏 sheet，忽略")
            continue

        # 在前5行内扫描 Staff ID 和 Name（兼容第1行空白的情况）
        staff_id, emp_name = "", ""
        for scan_row in ws.iter_rows(min_row=1, max_row=5, values_only=True):
            row_vals = list(scan_row)
            for i, val in enumerate(row_vals):
                label = str(val or "").strip().lower()
                if label == "staff id" and i + 1 < len(row_vals):
                    staff_id = str(row_vals[i + 1] or "").strip()
                if label == "name" and i + 1 < len(row_vals):
                    emp_name = str(row_vals[i + 1] or "").strip()
            if staff_id and emp_name:
                break  # 两个都找到了，不再继续扫描

        if not emp_name and not staff_id:
            print(f"  [跳过] {path.name} → sheet '{sname}'：未找到 Staff ID / Name")
            continue

        try:
            sheet_rows = extract_sheet(ws, staff_id, emp_name)
            print(f"  [读取] {path.name} → '{sname}' ({emp_name})：{len(sheet_rows)} 行有效数据")
            all_rows.extend(sheet_rows)
        except ValueError as e:
            print(f"  [警告] {path.name} → '{sname}'：{e}，已跳过")

    return fmt, all_rows


# ─────────────────────────────────────────────
# Step 5：写入 old-format 模板
# ─────────────────────────────────────────────
def write_old_format(records: list[dict], template_path: Path, out_path: Path):
    """将记录写入 old-format 模板的 FS sheet"""
    shutil.copy(template_path, out_path)
    wb = openpyxl.load_workbook(out_path)
    ws = wb[OLD_SHEET]

    # 清除模板里的示例数据行（第4行起）
    for row in ws.iter_rows(min_row=OLD_DATA_START, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    # 写入数据
    # old-format 列顺序（对应 row3 表头）：
    # A=Date, B=Role/Type, C=Module, D=Task Nature, E=Description,
    # F=Ref ID, G=Related Number, H=Hours, I=Status, J=Staff ID, K=Name
    for i, rec in enumerate(records):
        r = OLD_DATA_START + i
        ws.cell(r, 1).value  = rec["Date"]
        ws.cell(r, 2).value  = rec["Role / Type"]
        ws.cell(r, 3).value  = rec["Module"]
        ws.cell(r, 4).value  = rec["Task Nature"]
        ws.cell(r, 5).value  = rec["Description"]
        ws.cell(r, 6).value  = rec["Ref ID"]
        ws.cell(r, 7).value  = rec["Related Number"]
        ws.cell(r, 8).value  = rec["Hours"]
        ws.cell(r, 9).value  = rec["Status"]
        ws.cell(r, 10).value = rec["Staff ID"]
        ws.cell(r, 11).value = rec["Name"]

    wb.save(out_path)
    print(f"[输出] old-format → {out_path}（{len(records)} 行）")


# ─────────────────────────────────────────────
# Step 6：写入 new-format 模板
# ─────────────────────────────────────────────
def write_new_format(records: list[dict], template_path: Path, out_path: Path):
    """
    将记录写入 new-format 模板的 Timesheet sheet。
    Week# 区域（行1-9）保持原样不动，只填数据区（行10起）。
    new-format 列顺序（row10 表头）：
    A=Date, B=Role/Type, C=Task Nature, D=Module, E=Ref ID,
    F=Related Number(公式保留), G=Description, H=Hours,
    I=Status, J=Remarks, K=Staff ID, L=Name
    """
    shutil.copy(template_path, out_path)
    wb = openpyxl.load_workbook(out_path)
    ws = wb[NEW_SHEET]

    # 清除模板数据区（第11行起），保留第10行表头和第1-9行 Week 区
    for row in ws.iter_rows(min_row=NEW_DATA_START, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    for i, rec in enumerate(records):
        r = NEW_DATA_START + i
        ws.cell(r, 1).value  = rec["Date"]
        ws.cell(r, 2).value  = rec["Role / Type"]
        ws.cell(r, 3).value  = rec["Task Nature"]
        ws.cell(r, 4).value  = rec["Module"]
        ws.cell(r, 5).value  = rec["Ref ID"]
        # 列F：Related Number，写入公式（与模板行11-500一致）
        ws.cell(r, 6).value  = (
            f'=IFERROR(IF(C{r}="Change Enhancement","CE-"&E{r},'
            f'IF(C{r}="Change Request","CR-"&E{r},'
            f'IF(C{r}="EC Ticket","EC -"&E{r},'
            f'IF(C{r}="Production Incident","PROD -"&E{r},"")))),"")'
        )
        ws.cell(r, 7).value  = rec["Description"]
        ws.cell(r, 8).value  = rec["Hours"]
        ws.cell(r, 9).value  = rec["Status"]
        ws.cell(r, 10).value = None           # Remarks：留空
        ws.cell(r, 11).value = rec["Staff ID"]
        ws.cell(r, 12).value = rec["Name"]

    wb.save(out_path)
    print(f"[输出] new-format → {out_path}（{len(records)} 行）")


# ─────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Timesheet Merger")
    parser.add_argument("--src", default=DEFAULT_SRC_DIR, help="源文件目录")
    parser.add_argument("--out", default=DEFAULT_OUT_DIR, help="输出目录")
    parser.add_argument("--tpl-old", default=TEMPLATE_OLD, help="old-format 模板路径")
    parser.add_argument("--tpl-new", default=TEMPLATE_NEW, help="new-format 模板路径")
    args = parser.parse_args()

    src_dir      = Path(args.src)
    out_dir      = Path(args.out)
    tpl_old_path = Path(args.tpl_old)
    tpl_new_path = Path(args.tpl_new)

    # 检查路径
    if not src_dir.exists():
        raise FileNotFoundError(f"源文件目录不存在：{src_dir}")
    if not tpl_old_path.exists():
        raise FileNotFoundError(f"old-format 模板不存在：{tpl_old_path}")
    if not tpl_new_path.exists():
        raise FileNotFoundError(f"new-format 模板不存在：{tpl_new_path}")
    out_dir.mkdir(parents=True, exist_ok=True)

    # Step 1：扫描文件
    files = discover_files(src_dir)

    # Step 2-4：逐文件解析，按格式分类
    old_records: list[dict] = []
    new_records: list[dict] = []

    for path in files:
        print(f"\n处理：{path.name}")
        try:
            fmt, rows = parse_source(path)
            if fmt == "new":
                new_records.extend(rows)
            else:
                old_records.extend(rows)
        except Exception as e:
            print(f"  [错误] {path.name}：{e}，已跳过")

    # Step 5-6：写入输出
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    print()

    if old_records:
        out_old = out_dir / f"output_old_{timestamp}.xlsx"
        write_old_format(old_records, tpl_old_path, out_old)
    else:
        print("[跳过] 无 old-format 数据，不生成 output_old.xlsx")

    if new_records:
        out_new = out_dir / f"output_new_{timestamp}.xlsx"
        write_new_format(new_records, tpl_new_path, out_new)
    else:
        print("[跳过] 无 new-format 数据，不生成 output_new.xlsx")

    print(f"\n完成！old-format 共 {len(old_records)} 行，new-format 共 {len(new_records)} 行。")


if __name__ == "__main__":
    main()
