# 新增 _remove_external_links() 函数，在 write_new_format 里调用

+def _remove_external_links(xlsx_path: Path):
+    import zipfile, shutil, re
+    tmp_path = xlsx_path.with_suffix('.tmp.xlsx')
+    with zipfile.ZipFile(xlsx_path, 'r') as zin, \
+         zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
+        for item in zin.infolist():
+            if 'externalLinks' in item.filename:
+                continue                          # 整个 externalLinks 目录跳过
+            data = zin.read(item.filename)
+            if item.filename == 'xl/_rels/workbook.xml.rels':
+                data = re.sub(rb'<Relationship[^>]+externalLink[^>]+/>', b'', data)
+            if item.filename == 'xl/workbook.xml':
+                data = re.sub(rb'<externalReference[^>]+/>', b'', data)
+                data = re.sub(rb'<externalReferences>.*?</externalReferences>', b'', data, flags=re.DOTALL)
+            if item.filename == '[Content_Types].xml':
+                data = re.sub(rb'<Override[^>]+externalLink[^>]+/>', b'', data)
+            zout.writestr(item, data)
+    tmp_path.replace(xlsx_path)

 def write_new_format(...):
     shutil.copy(template_path, out_path)
+    _remove_external_links(out_path)   # ← 新增这一行
     wb = openpyxl.load_workbook(out_path)

-----------

if item.filename == 'xl/workbook.xml':
      data = re.sub(rb'<externalReference[^>]+/>', b'', data)
      data = re.sub(rb'<externalReferences>.*?</externalReferences>', b'', data, flags=re.DOTALL)
+     # 清除 definedNames 里引用外部工作簿的条目（形如 [1]Sheet!$A$1）
+     data = re.sub(rb'<definedName[^>]*>\[[0-9]+\][^<]*</definedName>', b'', data)

=======20260518========
+# 兼容文本日期格式列表
+_TEXT_DATE_FMTS = [
+    "%d-%b-%y",   # 6-Apr-26
+    "%d-%b-%Y",   # 6-Apr-2026
+    "%d/%m/%Y",   # 06/04/2026
+    "%d/%m/%y",   # 06/04/26
+    "%Y-%m-%d",   # 标准格式兜底
+]

 def _excel_serial_to_date(val) -> str | None:
     ...
     if isinstance(val, (int, float)):
         try:
             ...
         except Exception:
-            return str(val)      # 原来返回字符串，会污染数据
+            return None          # 改为 None，让调用方过滤

-    return str(val).strip() or None   # 原来不做任何文本识别
+    text = str(val).strip()
+    if not text:
+        return None
+    # 改动1：Total 行 → None → 调用方自动过滤
+    if "total" in text.lower():
+        return None
+    # 改动2：逐一尝试文本日期格式
+    for fmt in _TEXT_DATE_FMTS:
+        try:
+            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
+        except ValueError:
+            continue
+    return None   # 无法识别也返回 None

=======20260520新增改动=======

+import logging

+# 全局 logger 占位
+logger: logging.Logger = logging.getLogger("timesheet")

+def setup_logger(log_path: Path) -> logging.Logger:
+    """同时写控制台 + 文件的 logger"""
+    logger = logging.getLogger("timesheet")
+    fmt = logging.Formatter("%(asctime)s  %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
+    # 控制台 handler
+    ch = logging.StreamHandler(); ch.setFormatter(fmt); logger.addHandler(ch)
+    # 文件 handler
+    fh = logging.FileHandler(log_path, encoding="utf-8"); fh.setFormatter(fmt); logger.addHandler(fh)
+    return logger

 # parse_source / discover_files / write_output 所有 print → logger.info / logger.warning

 def main():
+    # 初始化 logger，日志文件 run_YYYYMMDD_HHMM.log 与输出同目录
+    global logger
+    logger = setup_logger(out_dir / f"run_{timestamp}.log")

+    skipped: list[str] = []   # 0行或跳过的文件
+    errored: list[str] = []   # 抛出异常的文件

     for idx, path in enumerate(files, 1):
         rows = parse_source(path)
-        if rows: all_records.extend(rows)
+        if rows:
+            all_records.extend(rows)
+        else:
+            skipped.append(path.name)   # 0行单独归入跳过列表

     # 末尾统计新增两个区块：
+    logger.info("【无数据/跳过】共 N 个文件：...")
+    logger.info("【异常错误】共 N 个文件：...")
+    logger.info("合计：X行 / Y人 | 跳过N个 | 异常N个")
