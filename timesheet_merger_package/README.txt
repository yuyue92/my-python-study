====================================================
  Timesheet Merger — 使用说明
====================================================

【功能】
  从多个员工 .xlsx 源文件提取考勤数据，
  自动识别格式，分别汇总为 output_old.xlsx / output_new.xlsx。

【目录结构】
  timesheet_merger.py     ← 主脚本
  templates/
    template-old-format.xlsx   ← old 格式空白模板
    template-new-format.xlsx   ← new 格式空白模板
  source_files/           ← 将所有员工源文件放这里
  output/                 ← 输出结果自动保存到此处

【依赖安装（首次运行）】
  pip install openpyxl pandas

【运行方式】

  ① 默认方式（使用以上目录结构）：
     python timesheet_merger.py

  ② 自定义路径：
     python timesheet_merger.py \
       --src  ./我的源文件目录 \
       --out  ./我的输出目录 \
       --tpl-old ./模板/old.xlsx \
       --tpl-new ./模板/new.xlsx

【格式自动识别规则】
  - new-format：源文件含名为 'Timesheet' 的 sheet，
                且第3行表头含 'Employee' 列
  - old-format：其余所有源文件

【注意事项】
  1. 源文件命名按员工姓名（如 WilliamChan.xlsx）
  2. README / Lists sheet 自动跳过
  3. 一个源文件可含多个员工 sheet（如 source-data1 含两人）
  4. Date 和 Hours 同时为空的行会被过滤
  5. 日期统一输出为 yyyy-mm-dd 格式
  6. 每次运行使用干净空白模板，不追加旧数据
  7. 输出文件名带时间戳，不会覆盖上次结果

【输出说明】
  output_old_YYYYMMDD_HHMM.xlsx  ← 汇总 old-format 源文件
  output_new_YYYYMMDD_HHMM.xlsx  ← 汇总 new-format 源文件
  （若无对应格式数据，该文件不会生成）
