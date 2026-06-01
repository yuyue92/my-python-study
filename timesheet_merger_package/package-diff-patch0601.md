# Diff 说明：追加 Function / Workstream / Subteam / Sub-Subteam 四列（剪贴板方案）

共 5 处改动。

---

## 改动 1 — import 区，补加三个库

**位置：** `import openpyxl` 上方

```python
# 原来
import openpyxl
```

```python
# 改后
import csv
import io
import pyperclip
import pandas as pd
import openpyxl
```

---

## 改动 2 — 配置区，补加 Staff 列位置配置 + OUTPUT_COLS 扩展

```python
# 原来
DEFAULT_SRC_DIR = "./source_files"
DEFAULT_OUT_DIR = "./output"

# 输出固定列定义（顺序即为输出顺序）
OUTPUT_COLS = ["Date", "Role/Type", "Task Nature", "Module", "Hours", "Ref ID", "Staff ID", "Name"]
```

```python
# 改后
DEFAULT_SRC_DIR = "./source_files"
DEFAULT_OUT_DIR = "./output"

# Staff 花名册列位置（0-indexed，A=0）
STAFF_COL_POSITIONS = [0, 8, 9, 10, 11]          # A, I, J, K, L
STAFF_COL_NAMES     = ["SID", "Function", "Workstream", "Subteam", "Sub-Subteam"]
STAFF_ID_COL        = "SID"

# 输出固定列定义（顺序即为输出顺序）
OUTPUT_COLS = [
    "Date", "Role/Type", "Task Nature", "Module", "Hours",
    "Ref ID", "Staff ID", "Name", "Function", "Workstream", "Subteam", "Sub-Subteam"
]
```

---

## 改动 3 — 新增 load_staff_from_clipboard() 函数

**位置：** `setup_logger()` 函数结束后，`discover_files()` 函数之前，整段插入：

```python
# ─────────────────────────────────────────────
# Staff 花名册：从剪贴板读取
# ─────────────────────────────────────────────
def load_staff_from_clipboard() -> dict:
    print("\n" + "=" * 50)
    print("【花名册加载】请先在 Excel 花名册中复制 A~L 列所有数据（含标题行），")
    print("然后按回车继续...")
    print("=" * 50)
    input()

    try:
        raw = pyperclip.paste()
        df = pd.read_csv(
            io.StringIO(raw),
            sep="\t",
            engine="python",
            quoting=csv.QUOTE_NONE,
            on_bad_lines="skip",
            header=None,
            dtype=str,
        )

        df = df.iloc[:, STAFF_COL_POSITIONS]
        df.columns = STAFF_COL_NAMES
        df.dropna(how="all", inplace=True)
        df[STAFF_ID_COL] = df[STAFF_ID_COL].str.strip()
        df = df[df[STAFF_ID_COL].notna() & (df[STAFF_ID_COL] != "") & (df[STAFF_ID_COL] != "SID")]

        result = {}
        for _, row in df.iterrows():
            sid = row[STAFF_ID_COL]
            result[sid] = {
                "Function":    str(row.get("Function",    "") or "").strip() or "-",
                "Workstream":  str(row.get("Workstream",  "") or "").strip() or "-",
                "Subteam":     str(row.get("Subteam",     "") or "").strip() or "-",
                "Sub-Subteam": str(row.get("Sub-Subteam", "") or "").strip() or "-",
            }

        logger.info(f"[花名册] 已从剪贴板加载 {len(result)} 条记录")
        return result

    except Exception as e:
        logger.warning(f"[花名册] 剪贴板读取失败：{e}，相关列将全部填 '-'")
        return {}
```

---

## 改动 4 — write_output 函数：签名加参数 + 写数据 + 列宽

### 4a 函数签名
```python
# 原来
def write_output(records: list[dict], out_path: Path):
```
```python
# 改后
def write_output(records: list[dict], out_path: Path, staff_dict: dict):
```

### 4b 写数据段
```python
# 原来
    for row_idx, rec in enumerate(records, 2):
        ws.cell(row_idx, 1).value = rec["Date"]
        ws.cell(row_idx, 2).value = rec["Role/Type"]
        ws.cell(row_idx, 3).value = rec["Task Nature"]
        ws.cell(row_idx, 4).value = rec["Module"]
        ws.cell(row_idx, 5).value = rec["Hours"]
        ws.cell(row_idx, 6).value = rec["Ref ID"]
        ws.cell(row_idx, 7).value = rec["Staff ID"]
        ws.cell(row_idx, 8).value = rec["Name"]

    # 列宽
    for col_idx, width in zip(range(1, 9), [14, 16, 18, 18, 8, 14, 14, 20]):
```
```python
# 改后
    for row_idx, rec in enumerate(records, 2):
        sid        = rec.get("Staff ID", "")
        staff_info = staff_dict.get(sid, {})
        ws.cell(row_idx,  1).value = rec["Date"]
        ws.cell(row_idx,  2).value = rec["Role/Type"]
        ws.cell(row_idx,  3).value = rec["Task Nature"]
        ws.cell(row_idx,  4).value = rec["Module"]
        ws.cell(row_idx,  5).value = rec["Hours"]
        ws.cell(row_idx,  6).value = rec["Ref ID"]
        ws.cell(row_idx,  7).value = rec["Staff ID"]
        ws.cell(row_idx,  8).value = rec["Name"]
        ws.cell(row_idx,  9).value = staff_info.get("Function",    "-")
        ws.cell(row_idx, 10).value = staff_info.get("Workstream",  "-")
        ws.cell(row_idx, 11).value = staff_info.get("Subteam",     "-")
        ws.cell(row_idx, 12).value = staff_info.get("Sub-Subteam", "-")

    # 列宽
    for col_idx, width in zip(range(1, 13), [14, 16, 18, 18, 8, 14, 14, 20, 16, 16, 16, 16]):
```

---

## 改动 5 — main()：加载花名册 + 传入 write_output

### 5a 在 `logger.info(f"{'=' * 50}")` 最后一行之后，`files = discover_files(src_dir)` 之前，插入：
```python
    # 加载花名册（从剪贴板）
    staff_dict = load_staff_from_clipboard()
```

### 5b write_output 调用加参数
```python
# 原来
        write_output(all_records, out_path)
```
```python
# 改后
        write_output(all_records, out_path, staff_dict)
```

---

## 改动总结

| # | 位置 | 内容 |
|---|---|---|
| 1 | import 区 | 加 `csv / io / pyperclip / pandas` |
| 2 | 配置区 | 加列位置配置，OUTPUT_COLS 扩展到 12 列 |
| 3 | 新函数 | `load_staff_from_clipboard()` |
| 4 | `write_output` | 加 staff_dict 参数，写第 9~12 列，列宽扩展 |
| 5 | `main()` | 调用加载花名册，传入 write_output |
