# -*- coding: utf-8 -*-
"""
Excel Folder Extractor

功能：
1. 递归扫描源文件夹及其所有子文件夹；
2. 将所有 Excel / 表格文件提取并移动到一个新的目标文件夹；
3. 如果目标文件夹中已经存在同名文件，直接覆盖替换；
4. 不删除、不修改源文件夹结构，只移动其中的 Excel 文件；
5. 不生成日志文件，只在控制台打印处理过程和最终统计。

使用方式一：直接运行
    python extract_excel_files.py

使用方式二：命令行指定路径
    python extract_excel_files.py --source "C:\\Users\\ASUS\\Desktop\\原始文件夹" --output "C:\\Users\\ASUS\\Desktop\\提取后的Excel"

预演模式：只打印将要执行的操作，不真正移动文件
    python extract_excel_files.py --source "源文件夹" --output "目标文件夹" --dry-run
"""

import argparse
import shutil
import sys
from pathlib import Path


# =========================
# 可手动修改的配置区
# =========================
SOURCE_DIR = r""          # 例如：r"C:\Users\ASUS\Desktop\原始文件夹"
OUTPUT_DIR = r""          # 例如：r"C:\Users\ASUS\Desktop\提取后的Excel"
DRY_RUN = False           # True = 只预演，不真正移动；False = 正式执行


# 需要提取的文件类型
EXCEL_EXTENSIONS = {
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".csv"
}


def setup_utf8_console():
    """尽量避免 Windows 控制台中文乱码。"""
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def is_excel_file(path: Path) -> bool:
    """判断是否为 Excel / 表格文件。"""
    return path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS


def is_relative_to_path(child: Path, parent: Path) -> bool:
    """兼容低版本 Python 的路径包含判断。"""
    try:
        child.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False
    except Exception:
        return False


def collect_excel_files(source_dir: Path, output_dir: Path):
    """
    递归收集源文件夹中的 Excel 文件。

    如果目标文件夹位于源文件夹内部，会自动跳过目标文件夹，
    防止把已经提取出来的文件重复扫描、重复移动。
    """
    excel_files = []

    for path in source_dir.rglob("*"):
        if is_relative_to_path(path, output_dir):
            continue

        if is_excel_file(path):
            excel_files.append(path)

    return excel_files


def extract_excel_files(source_dir: Path, output_dir: Path, dry_run: bool = False):
    """
    将所有 Excel 文件移动到目标文件夹。
    同名文件直接覆盖替换。
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    excel_files = collect_excel_files(source_dir, output_dir)

    print("\n" + "=" * 80)
    print("Excel 文件提取程序")
    print("=" * 80)
    print(f"源文件夹      : {source_dir}")
    print(f"目标文件夹    : {output_dir}")
    print(f"预演模式      : {dry_run}")
    print(f"发现表格文件数 : {len(excel_files)}")
    print("=" * 80)

    moved_count = 0
    replaced_count = 0
    skipped_count = 0
    failed_count = 0

    for index, src in enumerate(excel_files, start=1):
        try:
            dst = output_dir / src.name

            # 如果文件本身已经在目标文件夹中，则跳过
            if src.resolve() == dst.resolve():
                print(f"[{index}/{len(excel_files)}] 跳过：文件已在目标文件夹中 -> {src}")
                skipped_count += 1
                continue

            will_replace = dst.exists()

            if dry_run:
                if will_replace:
                    print(f"[{index}/{len(excel_files)}] 预演覆盖：{src}  ->  {dst}")
                else:
                    print(f"[{index}/{len(excel_files)}] 预演移动：{src}  ->  {dst}")
                moved_count += 1
                if will_replace:
                    replaced_count += 1
                continue

            # 同名文件直接覆盖：先删除目标旧文件，再移动新文件
            if will_replace:
                dst.unlink()
                replaced_count += 1
                print(f"[{index}/{len(excel_files)}] 覆盖旧文件：{dst}")

            shutil.move(str(src), str(dst))
            moved_count += 1
            print(f"[{index}/{len(excel_files)}] 已移动：{src}  ->  {dst}")

        except Exception as e:
            failed_count += 1
            print(f"[{index}/{len(excel_files)}] 失败：{src}")
            print(f"    错误原因：{e}")

    print("\n" + "=" * 80)
    print("处理完成，最终统计如下：")
    print("=" * 80)
    print(f"发现表格文件数 : {len(excel_files)}")
    print(f"成功移动数量   : {moved_count}")
    print(f"覆盖替换数量   : {replaced_count}")
    print(f"跳过数量       : {skipped_count}")
    print(f"失败数量       : {failed_count}")
    print(f"目标文件夹     : {output_dir}")
    print("=" * 80)

    return {
        "found": len(excel_files),
        "moved": moved_count,
        "replaced": replaced_count,
        "skipped": skipped_count,
        "failed": failed_count,
    }


def prompt_path(prompt_text: str) -> Path:
    """交互式输入路径。"""
    while True:
        value = input(prompt_text).strip().strip('"').strip("'")
        if value:
            return Path(value)
        print("路径不能为空，请重新输入。")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Recursively extract Excel files from nested folders into one output folder."
    )
    parser.add_argument(
        "--source",
        type=str,
        default="",
        help="Source root folder path."
    )
    parser.add_argument(
        "--output",
        type=str,
        default="",
        help="Output folder path."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview only. Do not move files."
    )
    return parser.parse_args()


def main():
    setup_utf8_console()
    args = parse_args()

    source_text = args.source or SOURCE_DIR
    output_text = args.output or OUTPUT_DIR
    dry_run = bool(args.dry_run or DRY_RUN)

    print("\nExcel 文件提取程序")
    print("-" * 80)
    print("说明：本程序只会把 Excel / 表格文件移动到目标文件夹。")
    print("说明：不会删除源文件夹，也不会清理空文件夹。")
    print("说明：目标文件夹中如有同名文件，会直接覆盖替换。")
    print("-" * 80)

    if source_text:
        source_dir = Path(source_text.strip().strip('"').strip("'")).expanduser()
    else:
        source_dir = prompt_path("请输入源文件夹路径：").expanduser()

    if output_text:
        output_dir = Path(output_text.strip().strip('"').strip("'")).expanduser()
    else:
        output_dir = prompt_path("请输入目标文件夹路径：").expanduser()

    if not source_dir.exists():
        print(f"\n[错误] 源文件夹不存在：{source_dir}")
        return

    if not source_dir.is_dir():
        print(f"\n[错误] 源路径不是文件夹：{source_dir}")
        return

    try:
        extract_excel_files(source_dir, output_dir, dry_run=dry_run)
    except KeyboardInterrupt:
        print("\n用户中断运行。")
    except Exception as e:
        print(f"\n[错误] 程序运行异常：{e}")


if __name__ == "__main__":
    main()
