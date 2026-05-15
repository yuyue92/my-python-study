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
