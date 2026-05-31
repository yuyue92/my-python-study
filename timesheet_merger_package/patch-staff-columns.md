# Diff 说明：追加 Team Leader / Department 两列

共 4 处改动，按顺序操作即可。

---

## 改动 1 — 顶部 import 区，补加 pandas

**位置：** 第 29 行附近，`import openpyxl` 上方

```python
# 原来
import openpyxl
```

```python
# 改后
import pandas as pd
import openpyxl
```

---

## 改动 2 — 配置区，补加 staff Excel 路径 + 列名配置

**位置：** 配置区，`DEFAULT_OUT_DIR` 那几行后面紧接着加

```python
# 原来
DEFAULT_SRC_DIR = "./source_files"
DEFAULT_OUT_DIR = "./output"

# 输出固定列定义（顺序即为输出顺序）
OUTPUT_COLS = ["Date", "Role/Type", "Task Nature", "Module", "Hours", "Ref ID", "Staff ID", "Name"]
```

```python
# 改后
DEFAULT_SRC_DIR  = "./source_files"
DEFAULT_OUT_DIR  = "./output"
STAFF_EXCEL_PATH = "./staff.xlsx"   # ← 花名册路径，按实际修改
STAFF_HEADER_ROW = 0                # 标题行（0 = 第1行）
STAFF_DATA_START = 3                # 数据起始行，跳过前面空行（3 = 第4行）
STAFF_ID_COL     = "staffID"        # Staff ID 列名，按实际修改
STAFF_TL_COL     = "team-leader"    # Team Leader 列名，按实际修改
STAFF_DEPT_COL   = "department"     # Department 列名，按实际修改

# 输出固定列定义（顺序即为输出顺序）
OUTPUT_COLS = ["Date", "Role/Type", "Task Nature", "Module", "Hours", "Ref ID", "Staff ID", "Name", "Team Leader", "Department"]
```

---

## 改动 3 — 新增 load_staff_dict() 函数

**位置：** `setup_logger()` 函数定义结束后，`discover_files()` 函数之前，插入以下新函数

```python
# 新增（整段插入）
# ─────────────────────────────────────────────
# 花名册加载：返回 {staffID: {"team-leader": ..., "department": ...}}
# ─────────────────────────────────────────────
def load_staff_dict(path: str = STAFF_EXCEL_PATH) -> dict:
    try:
        df = pd.read_excel(
            path,
            header=STAFF_HEADER_ROW,
            skiprows=range(1, STAFF_DATA_START),
            dtype={STAFF_ID_COL: str},
        )
        df.dropna(how="all", inplace=True)
        result = {}
        for _, row in df.iterrows():
            sid = str(row.get(STAFF_ID_COL, "") or "").strip()
            if sid:
                result[sid] = {
                    "team-leader": str(row.get(STAFF_TL_COL, "") or "").strip() or "-",
                    "department":  str(row.get(STAFF_DEPT_COL, "") or "").strip() or "-",
                }
        logger.info(f"[花名册] 已加载 {len(result)} 条记录：{path}")
        return result
    except FileNotFoundError:
        logger.warning(f"[花名册] 文件未找到：{path}，team-leader / department 将全部填 '-'")
        return {}
    except Exception as e:
        logger.warning(f"[花名册] 加载失败：{e}，team-leader / department 将全部填 '-'")
        return {}
```

---

## 改动 4 — write_output() 追加两列写入，main() 里加载花名册并传入

### 4a — write_output 函数签名 + 写数据部分

```python
# 原来
def write_output(records: list[dict], out_path: Path):
```

```python
# 改后
def write_output(records: list[dict], out_path: Path, staff_dict: dict = {}):
```

---

```python
# 原来（write_output 里写数据那段）
    # 写数据
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
    # 写数据
    for row_idx, rec in enumerate(records, 2):
        sid = rec.get("Staff ID", "")
        staff_info = staff_dict.get(sid, {})
        ws.cell(row_idx, 1).value = rec["Date"]
        ws.cell(row_idx, 2).value = rec["Role/Type"]
        ws.cell(row_idx, 3).value = rec["Task Nature"]
        ws.cell(row_idx, 4).value = rec["Module"]
        ws.cell(row_idx, 5).value = rec["Hours"]
        ws.cell(row_idx, 6).value = rec["Ref ID"]
        ws.cell(row_idx, 7).value = rec["Staff ID"]
        ws.cell(row_idx, 8).value = rec["Name"]
        ws.cell(row_idx, 9).value  = staff_info.get("team-leader", "-")
        ws.cell(row_idx, 10).value = staff_info.get("department",  "-")

    # 列宽
    for col_idx, width in zip(range(1, 11), [14, 16, 18, 18, 8, 14, 14, 20, 16, 16]):
```

### 4b — main() 里加载花名册并传给 write_output

```python
# 原来（main() 里 Step 5 写入那段）
    if all_records:
        out_path = out_dir / f"output_{timestamp}.xlsx"
        write_output(all_records, out_path)
```

```python
# 改后
    if all_records:
        out_path  = out_dir / f"output_{timestamp}.xlsx"
        staff_dict = load_staff_dict()
        write_output(all_records, out_path, staff_dict)
```

---

## 改动总结

| # | 位置 | 内容 |
|---|---|---|
| 1 | import 区 | 加 `import pandas as pd` |
| 2 | 配置区 | 加 staff Excel 路径 + 列名配置，OUTPUT_COLS 扩展到 10 列 |
| 3 | 新函数 | `load_staff_dict()` 读花名册，返回字典 |
| 4 | write_output + main | 追加第 9、10 列写入；main 里调用加载并传入 |
