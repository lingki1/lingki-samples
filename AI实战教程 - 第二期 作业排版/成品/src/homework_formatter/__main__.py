"""命令行入口：python -m homework_formatter"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from . import pipeline
from .decisions import DecisionsError, load_decisions
from .inventory import Outcome
from .roster import RosterError, load_roster
from .rules import RulesError, load_rules

DEFAULT_INPUT = "学生作业"
DEFAULT_ROSTER = "名单.xlsx"
DEFAULT_OUTPUT = "排版结果"
DEFAULT_DECISIONS = "老师的决定.xlsx"
DEFAULT_RULES = "规则.toml"


class _Stop(Exception):
    """没法继续，把原因告诉老师后退出。"""


def _load_optional(given: str | None, default: str, label: str, loader, error_type):
    """名单和决定表：默认位置上有就用，没有就算了；明确指定了却找不到才报错。"""
    path = Path(given or default)
    if not path.is_file():
        if given:
            raise _Stop(f"找不到{label}：{path}")
        return None
    try:
        return loader(path)
    except error_type as error:
        raise _Stop(f"{label}有问题：{error}") from error
    except Exception as error:  # 不是真正的 .xlsx、加了密码、正被 Excel 开着……
        raise _Stop(
            f"{label}「{path.name}」打不开。请确认它是 .xlsx 格式、没有加密码、"
            f"也没有正被别的程序打开。（{type(error).__name__}）"
        ) from error


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="把学生交的 Word 作业批量改成学校要求的格式")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="学生作业所在的文件夹")
    parser.add_argument("--roster", default=None, help="班级名单（Excel，两列：学号、姓名）")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="结果放在哪个文件夹")
    parser.add_argument("--decisions", default=None, help="老师对个别文件的决定（Excel）")
    parser.add_argument("--rules", default=DEFAULT_RULES, help="格式规则表")
    parser.add_argument("--open", action="store_true", help="做完后打开结果文件夹")
    args = parser.parse_args(argv)

    try:
        input_dir = Path(args.input)
        if not input_dir.is_dir():
            raise _Stop(f"找不到作业文件夹：{input_dir.resolve()}")
        if not Path(args.rules).is_file():
            raise _Stop(f"找不到规则表：{Path(args.rules).resolve()}")
        try:
            rules = load_rules(Path(args.rules))
        except RulesError as error:
            raise _Stop(f"规则表有问题：{error}") from error
        roster = _load_optional(args.roster, DEFAULT_ROSTER, "名单", load_roster, RosterError)
        decisions = _load_optional(
            args.decisions, DEFAULT_DECISIONS, "决定表", load_decisions, DecisionsError
        )
        try:
            result = pipeline.run(input_dir, Path(args.output), rules, roster, decisions)
        except pipeline.PipelineError as error:
            raise _Stop(str(error)) from error
    except _Stop as stop:
        print(stop)
        return 2

    inventory = result.inventory
    print(f"共 {len(inventory.records)} 个文件：")
    for outcome in Outcome:
        if inventory.count(outcome):
            print(f"  {outcome.value}　{inventory.count(outcome)}")
    if roster is not None and inventory.missing_students:
        names = "、".join(name for _, name in inventory.missing_students)
        print(f"名单上有、但没找到作业的学生：{names}")
    if inventory.unmatched_decisions:
        names = "、".join(inventory.unmatched_decisions)
        print(f"注意：决定表里写了、但作业文件夹里没有这些文件（是不是打错了？）：{names}")
    if (result.run_dir / pipeline.REVIEW_FOLDER).is_dir():
        print("有需要您确认的作业：请看「需人工确认」文件夹，原因写在报告的「说明」一列。")
    if inventory.count(Outcome.UNPROCESSED):
        print("有没法处理的文件（没有生成排好的文件）：原因写在报告的「说明」一列。")
    print(f"结果在：{result.run_dir}")
    print(f"报告：{result.report_path}")
    if args.open and hasattr(os, "startfile"):
        os.startfile(result.run_dir)
    if result.incomplete:
        print(f"警告：这次运行中途出了意外，没有做完（{result.incomplete}）。结果不全，请重新运行一次。")
        return 1
    if not result.originals_untouched:
        print("警告：运行期间「学生作业」里的文件发生了变化，请先不要使用这次的结果。")
        return 1
    print(f"原件核对：{result.file_count} 个文件，运行前后完全一致。")
    return 0


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        # 输出被重定向时，遇到当前编码写不出的字也不要中断
        stream.reconfigure(errors="replace")
    sys.exit(main())
