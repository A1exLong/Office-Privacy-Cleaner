import argparse
from dataclasses import asdict
import json
from pathlib import Path
import sys

from .engine import clean_file
from .models import CleanError


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="本地清理 Office 元数据与协作身份信息，保留数据并生成 _clean 副本。")
    parser.add_argument("files", nargs="*", help="支持多个文件，也可将文件拖到 exe 上。")
    parser.add_argument("--report", action="store_true", help="在输出旁保存 JSON 处理报告（包含本机文件路径）。")
    parser.add_argument("--no-pause", action="store_true", help="可执行程序运行结束后不等待回车。")
    args = parser.parse_args(argv)
    paths = args.files
    if not paths:
        print("将文件拖到 OfficePrivacyCleaner.exe 上，或输入完整文件路径。")
        try:
            text = input("文件路径（直接回车退出）：").strip().strip('"')
        except (EOFError, KeyboardInterrupt):
            return 0
        if not text:
            return 0
        paths = [text]
    failed = 0
    seen = set()
    for value in paths:
        key = str(Path(value).resolve()).casefold()
        if key in seen:
            continue
        seen.add(key)
        print(f"\n处理：{Path(value).name}", flush=True)
        try:
            report = clean_file(value)
            print(f"已生成：{report.output}")
            for name, count in report.changes.items():
                print(f"  清理：{name} × {count}")
            for warning in report.warnings:
                print(f"  范围说明：{warning}")
            if args.report:
                # Report failure is reported separately; the cleaned document is already valid.
                try:
                    Path(report.output + ".clean-report.json").write_text(json.dumps(asdict(report), ensure_ascii=False, indent=2), encoding="utf-8")
                except OSError as exc:
                    print(f"  无法保存处理报告：{exc}", file=sys.stderr)
                    failed += 1
        except (CleanError, OSError, ValueError) as exc:
            failed += 1
            print(f"失败：{exc}", file=sys.stderr)
    print(f"\n完成。失败任务：{failed}")
    if getattr(sys, "frozen", False) and not args.no_pause:
        try:
            input("按回车退出……")
        except (EOFError, KeyboardInterrupt):
            pass
    return 1 if failed else 0
