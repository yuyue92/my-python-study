"""
flatten_files.py
把指定文件夹下所有文件（多层级）flat 复制到目标文件夹，去除目录层级。
同名文件自动重命名：filename.ext → filename_2.ext → filename_3.ext ...

用法：
  python flatten_files.py <源目录> <目标目录>

示例：
  python flatten_files.py ./src ./output_flat
"""

import os
import shutil
import argparse


def flatten(src: str, dst: str) -> None:
    src = os.path.abspath(src)
    dst = os.path.abspath(dst)

    if not os.path.isdir(src):
        print(f"[错误] 源目录不存在：{src}")
        return

    os.makedirs(dst, exist_ok=True)

    copied = 0
    skipped = 0
    renamed = 0

    for root, _, files in os.walk(src):
        # 跳过目标目录本身（防止源和目标重叠时死循环）
        if os.path.abspath(root) == dst:
            continue

        for filename in files:
            src_path = os.path.join(root, filename)
            dst_path = resolve_conflict(dst, filename)

            if dst_path != os.path.join(dst, filename):
                renamed += 1
                print(f"[重命名] {filename} → {os.path.basename(dst_path)}")

            shutil.copy2(src_path, dst_path)
            rel = os.path.relpath(src_path, src)
            print(f"[复制]   {rel}")
            copied += 1

    print(f"\n完成：共复制 {copied} 个文件，其中 {renamed} 个因同名被重命名，{skipped} 个跳过。")
    print(f"输出目录：{dst}")


def resolve_conflict(dst_dir: str, filename: str) -> str:
    """同名文件加数字后缀，避免覆盖。"""
    candidate = os.path.join(dst_dir, filename)
    if not os.path.exists(candidate):
        return candidate

    name, ext = os.path.splitext(filename)
    counter = 2
    while True:
        new_name = f"{name}_{counter}{ext}"
        candidate = os.path.join(dst_dir, new_name)
        if not os.path.exists(candidate):
            return candidate
        counter += 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="递归展平目录，所有文件复制到同一层级")
    parser.add_argument("src", help="源目录路径")
    parser.add_argument("dst", help="目标（输出）目录路径")
    args = parser.parse_args()

    flatten(args.src, args.dst)
