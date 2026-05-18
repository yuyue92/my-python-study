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

# ── 改动1：sheet 选取逻辑 ──
-    # 只处理 sheetname == 'Timesheet' 的 sheet，其余一律跳过
-    for sname in wb.sheetnames:
-        if sname.lower() != "timesheet":
-            print(f"  [跳过] sheet '{sname}'：非 Timesheet")
-            continue
-        ws = wb[sname]
+    # 优先取 active sheet；若不可用则 fallback 到名为 Timesheet 的 sheet
+    active_ws = wb.active
+    if active_ws is not None:
+        candidate = (active_ws.title, active_ws)
+        print(f"  [候选] 使用 active sheet：'{active_ws.title}'")
+    else:
+        for sname in wb.sheetnames:
+            if sname.lower() == "timesheet":
+                candidate = (sname, wb[sname])
+                print(f"  [候选] active 不可用，fallback 到 sheet：'{sname}'")
+                break

# ── 改动2：扫描范围 5→10 行，Staff ID/Name 独立查找 ──
-        for scan_row in ws.iter_rows(min_row=1, max_row=5, values_only=True):
+        for scan_row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
             for i, val in enumerate(row_vals):
                 label = str(val or "").strip().lower()
-                if label == "staff id" and i + 1 < len(row_vals):
+                if not staff_id and label == "staff id" and i + 1 < len(row_vals):
                     staff_id = ...
-                if label == "name" and i + 1 < len(row_vals):
+                if not emp_name and label == "name" and i + 1 < len(row_vals):
                     emp_name = ...
