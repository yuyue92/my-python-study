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
