# -*- coding: utf-8 -*-
"""
Excel Folder Merger

功能：
1. 递归扫描一个根目录下的所有 Excel 文件；
2. 将所有 Excel 文件移动到一个目标文件夹；
3. 自动处理重名文件，避免覆盖；
4. 删除移动后留下的空文件夹；
5. 输出基础打印信息，并保存运行日志。

使用方式：
方式一：直接双击/运行本脚本，然后按提示输入源文件夹和目标文件夹。
方式二：先修改下方 CONFIG 区域的 SOURCE_DIR 和 OUTPUT_DIR，再运行。
方式三：命令行运行：
    python merge_excel_files.py --source "C:\\Users\\ASUS\\Desktop\\原始文件夹" --output "C:\\Users\\ASUS\\Desktop\\合并后的Excel"

注意：
- 默认是“移动”文件，不是复制文件。
- 如果想先预演，不真正移动和删除，请把 DRY_RUN 改成 True，或命令行加 --dry-run。
"""

import argparse
import logging
import shutil
import sys
from datetime import datetime
from pathlib import Path


# =========================
# CONFIG：可按需修改
# =========================
SOURCE_DIR = r""          # 例如：r"C:\Users\ASUS\Desktop\原始文件夹"
OUTPUT_DIR = r""          # 例如：r"C:\Users\ASUS\Desktop\合并后的Excel"
DRY_RUN = False           # True = 只预演，不真正移动/删除；False = 正式执行

# Excel 文件后缀
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


def setup_logger(log_dir: Path) -> Path:
    """创建日志文件。"""
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"merge_excel_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return log_file


def is_excel_file(path: Path) -> bool:
    """判断是否为 Excel/表格文件。"""
    return path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS


def collect_excel_files(source_dir: Path, output_dir: Path):
    """
    递归收集 Excel 文件。
    如果目标文件夹在源文件夹内部，则跳过目标文件夹，避免重复扫描。
    """
    excel_files = []

    source_resolved = source_dir.resolve()
    output_resolved = output_dir.resolve()

    for path in source_dir.rglob("*"):
        try:
            resolved_path = path.resolve()
        except Exception:
            resolved_path = path

        # 跳过输出文件夹自身及其内部文件，避免重复移动
        try:
            if output_resolved == resolved_path or output_resolved in resolved_path.parents:
                continue
        except Exception:
            pass

        if is_excel_file(path):
            excel_files.append(path)

    return excel_files


def move_excel_files(source_dir: Path, output_dir: Path, dry_run: bool = False):
    """移动所有 Excel 文件到目标文件夹。"""
    output_dir.mkdir(parents=True, exist_ok=True)

    excel_files = collect_excel_files(source_dir, output_dir)

    logging.info("=" * 80)
    logging.info("Start scanning and moving Excel files")
    logging.info(f"Source folder : {source_dir}")
    logging.info(f"Output folder : {output_dir}")
    logging.info(f"Dry run       : {dry_run}")
    logging.info(f"Excel files found: {len(excel_files)}")
    logging.info("=" * 80)

    moved_count = 0
    skipped_count = 0
    failed_count = 0

    for src in excel_files:
        try:
            dst = output_dir / src.name

            if src.resolve() == dst.resolve():
                logging.info(f"[SKIP] Already in output folder: {src}")
                skipped_count += 1
                continue

            if dry_run:
                if dst.exists():
                    logging.info(f"[DRY-RUN REPLACE] {src}  ->  {dst}")
                else:
                    if dst.exists():
                        dst.unlink()
                        logging.info(f"[REPLACE] Existing file deleted: {dst}")
            else:
                shutil.move(str(src), str(dst))
                logging.info(f"[MOVED] {src}  ->  {dst}")

            moved_count += 1

        except Exception as e:
            logging.exception(f"[FAILED] Move failed: {src} | Error: {e}")
            failed_count += 1

    logging.info("-" * 80)
    logging.info(f"Move summary | moved={moved_count}, skipped={skipped_count}, failed={failed_count}")
    logging.info("-" * 80)

    return {
        "found": len(excel_files),
        "moved": moved_count,
        "skipped": skipped_count,
        "failed": failed_count,
    }


def delete_empty_dirs(source_dir: Path, output_dir: Path, dry_run: bool = False):
    """
    删除空文件夹。
    从最深层开始删除，避免父文件夹因子文件夹未删而无法删除。
    不删除源文件夹本身，不删除目标文件夹。
    """
    logging.info("=" * 80)
    logging.info("Start deleting empty folders")
    logging.info("=" * 80)

    deleted_count = 0
    failed_count = 0

    source_resolved = source_dir.resolve()
    output_resolved = output_dir.resolve()

    # 按路径深度从深到浅排序
    all_dirs = [p for p in source_dir.rglob("*") if p.is_dir()]
    all_dirs.sort(key=lambda p: len(p.parts), reverse=True)

    for folder in all_dirs:
        try:
            folder_resolved = folder.resolve()

            # 不删除源目录本身
            if folder_resolved == source_resolved:
                continue

            # 不删除输出目录及其父子关联目录
            if folder_resolved == output_resolved:
                continue
            if output_resolved in folder_resolved.parents:
                continue

            # 如果文件夹为空，则删除
            if not any(folder.iterdir()):
                if dry_run:
                    logging.info(f"[DRY-RUN DELETE EMPTY DIR] {folder}")
                else:
                    folder.rmdir()
                    logging.info(f"[DELETED EMPTY DIR] {folder}")
                deleted_count += 1

        except Exception as e:
            logging.exception(f"[FAILED] Delete empty folder failed: {folder} | Error: {e}")
            failed_count += 1

    logging.info("-" * 80)
    logging.info(f"Delete summary | deleted_empty_dirs={deleted_count}, failed={failed_count}")
    logging.info("-" * 80)

    return {
        "deleted_empty_dirs": deleted_count,
        "failed_delete_dirs": failed_count,
    }


def prompt_path(prompt_text: str) -> Path:
    """交互式输入路径。"""
    while True:
        value = input(prompt_text).strip().strip('"').strip("'")
        if value:
            return Path(value)
        print("Path cannot be empty. Please input again.")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Recursively move all Excel files from nested folders into one output folder."
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
        help="Preview only. Do not move files or delete folders."
    )
    return parser.parse_args()


def main():
    setup_utf8_console()
    args = parse_args()

    source_text = args.source or SOURCE_DIR
    output_text = args.output or OUTPUT_DIR
    dry_run = bool(args.dry_run or DRY_RUN)

    print("\nExcel Folder Merger")
    print("=" * 80)
    print("This script will move all Excel files from nested folders into one output folder.")
    print("It will also delete empty folders after moving files.")
    print("=" * 80)

    if source_text:
        source_dir = Path(source_text.strip().strip('"').strip("'"))
    else:
        source_dir = prompt_path("Please input source folder path: ")

    if output_text:
        output_dir = Path(output_text.strip().strip('"').strip("'"))
    else:
        output_dir = prompt_path("Please input output folder path: ")

    source_dir = source_dir.expanduser()
    output_dir = output_dir.expanduser()

    if not source_dir.exists():
        print(f"\n[ERROR] Source folder does not exist: {source_dir}")
        return

    if not source_dir.is_dir():
        print(f"\n[ERROR] Source path is not a folder: {source_dir}")
        return

    # 如果输出目录在源目录外，日志保存在输出目录；
    # 如果输出目录尚不存在，也可以先创建。
    output_dir.mkdir(parents=True, exist_ok=True)
    log_file = setup_logger(output_dir)

    logging.info(f"Log file: {log_file}")

    try:
        move_result = move_excel_files(source_dir, output_dir, dry_run=dry_run)
        delete_result = delete_empty_dirs(source_dir, output_dir, dry_run=dry_run)

        logging.info("=" * 80)
        logging.info("All tasks finished")
        logging.info(f"Excel files found       : {move_result['found']}")
        logging.info(f"Excel files moved       : {move_result['moved']}")
        logging.info(f"Excel files skipped     : {move_result['skipped']}")
        logging.info(f"Move failed             : {move_result['failed']}")
        logging.info(f"Empty folders deleted   : {delete_result['deleted_empty_dirs']}")
        logging.info(f"Delete folder failed    : {delete_result['failed_delete_dirs']}")
        logging.info(f"Log saved to            : {log_file}")
        logging.info("=" * 80)

        print("\nDone.")
        print(f"Output folder: {output_dir}")
        print(f"Log file     : {log_file}")

    except KeyboardInterrupt:
        logging.warning("User interrupted by keyboard.")
        print("\nInterrupted by user.")
    except Exception as e:
        logging.exception(f"Unexpected error: {e}")
        print(f"\n[ERROR] Unexpected error: {e}")


if __name__ == "__main__":
    main()
