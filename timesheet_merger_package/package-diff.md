-        # 从第1行读取 Staff ID 和 Name
-        row1 = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
-        staff_id, emp_name = "", ""
-        for i, val in enumerate(row1):
-            if str(val or "").strip().lower() == "staff id" and i + 1 < len(row1):
-                staff_id = str(row1[i + 1] or "").strip()
-            if str(val or "").strip().lower() == "name" and i + 1 < len(row1):
-                emp_name = str(row1[i + 1] or "").strip()
+        # 在前5行内扫描 Staff ID 和 Name（兼容第1行空白的情况）
+        staff_id, emp_name = "", ""
+        for scan_row in ws.iter_rows(min_row=1, max_row=5, values_only=True):
+            row_vals = list(scan_row)
+            for i, val in enumerate(row_vals):
+                label = str(val or "").strip().lower()
+                if label == "staff id" and i + 1 < len(row_vals):
+                    staff_id = str(row_vals[i + 1] or "").strip()
+                if label == "name" and i + 1 < len(row_vals):
+                    emp_name = str(row_vals[i + 1] or "").strip()
+            if staff_id and emp_name:
+                break  # 两个都找到了，不再继续扫描
