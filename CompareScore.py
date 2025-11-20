#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
对两次考试成绩进行对比分析，找出总分及各科年级排名
进步/退步超过某个百分比阈值的学生。

支持两种模式：
1）单个 Excel 文件 + 指定两个 sheet；
2）两个 Excel 文件：上次考试文件 & 本次考试文件。

默认适配你当前的表格结构：
- 月考总分排名列：    '校排名.9'
- 期中总分排名列：    '折算后排名'
- 各科排名列：        '校排名', '校排名.1', …

可以根据实际情况在脚本中修改映射。
"""

import argparse
from pathlib import Path
from typing import Optional, List

import pandas as pd
import numpy as np


def prepare_exam_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    统一两张表的关键列名，确保都至少有：
    - '班级'
    - '姓名'

    适配你目前期中表中用 Unnamed:0 / Unnamed:1 的情况。
    """
    # 班级列
    if "班级" not in df.columns:
        if "Unnamed: 0" in df.columns:
            df = df.rename(columns={"Unnamed: 0": "班级"})
    # 姓名列
    if "姓名" not in df.columns:
        if "Unnamed: 1" in df.columns:
            df = df.rename(columns={"Unnamed: 1": "姓名"})
    return df


def analyze_scores(
    exam_prev: pd.DataFrame,
    exam_curr: pd.DataFrame,
    output_path: str,
    class_name: Optional[str] = None,
    threshold: float = 0.2,
) -> None:
    """
    核心分析函数：
    - exam_prev: 上次考试（如“月考”）的 DataFrame
    - exam_curr: 本次考试（如“期中”）的 DataFrame
    - class_name: 若为 None，则分析全部班级；否则只分析指定班
    - threshold: 进步/退步百分比阈值（默认 0.2 = 20%）
    """

    exam_prev = prepare_exam_df(exam_prev)
    exam_curr = prepare_exam_df(exam_curr)

    if "班级" not in exam_prev.columns:
        raise ValueError("上次考试数据中缺少 '班级' 列")
    if "班级" not in exam_curr.columns:
        raise ValueError("本次考试数据中缺少 '班级' 列")
    if "姓名" not in exam_prev.columns:
        raise ValueError("上次考试数据中缺少 '姓名' 列")
    if "姓名" not in exam_curr.columns:
        raise ValueError("本次考试数据中缺少 '姓名' 列")

    # ===================== 1. 科目与列映射（根据你现在的表结构） =====================
    subjects: List[str] = [
        "语文",
        "数学",
        "英语",
        "物理",
        "化学",
        "生物",
        "政治",
        "历史",
        "地理",
        "总分",
    ]

    # 月考各科年级排名列
    prev_rank_cols = {
        "语文": "校排名",
        "数学": "校排名.1",
        "英语": "校排名.2",
        "物理": "校排名.3",
        "化学": "校排名.4",
        "生物": "校排名.5",
        "政治": "校排名.6",
        "历史": "校排名.7",
        "地理": "校排名.8",
        "总分": "校排名.9",  # 月考总分年级排名
    }

    # 期中各科年级排名列
    curr_rank_cols = {
        "语文": "校排名",
        "数学": "校排名.1",
        "英语": "校排名.2",
        "物理": "校排名.3",
        "化学": "校排名.4",
        "生物": "校排名.5",
        "政治": "校排名.6",
        "历史": "校排名.7",
        "地理": "校排名.8",
        "总分": "校排名.9",
    }

    # 安全检查列是否存在
    for subj in subjects:
        if prev_rank_cols[subj] not in exam_prev.columns:
            raise ValueError(f"上次考试中缺少 {subj} 的排名列：{prev_rank_cols[subj]}")
        if curr_rank_cols[subj] not in exam_curr.columns:
            raise ValueError(f"本次考试中缺少 {subj} 的排名列：{curr_rank_cols[subj]}")

    # ===================== 2. 计算全年级各科人数 =====================
    rows = []
    for subj in subjects:
        prev_total = exam_prev[prev_rank_cols[subj]].notna().sum()
        curr_total = exam_curr[curr_rank_cols[subj]].notna().sum()
        rows.append(
            {
                "科目": subj,
                "上次考试总人数": int(prev_total),
                "本次考试总人数": int(curr_total),
            }
        )
    counts_df = pd.DataFrame(rows)

    # ===================== 3. 准备按班级筛选 & 对齐键 =====================
    if class_name:
        # 只分析指定班，如“6班”
        prev_df = exam_prev[exam_prev["班级"] == class_name].copy()
        curr_df = exam_curr[exam_curr["班级"] == class_name].copy()
        key_cols = ["姓名"]  # 班内通常姓名唯一
        label_prefix = class_name
    else:
        # 分析全年级所有班级
        prev_df = exam_prev.copy()
        curr_df = exam_curr.copy()
        key_cols = ["班级", "姓名"]  # 全年级时，按 (班级, 姓名) 对齐更稳妥
        label_prefix = "全部班级"

    records = []

    # ===================== 4. 各科计算百分比 & 进退步 =====================
    for subj in subjects:
        prev_total = counts_df.loc[counts_df["科目"] == subj, "上次考试总人数"].iloc[0]
        curr_total = counts_df.loc[counts_df["科目"] == subj, "本次考试总人数"].iloc[0]

        if prev_total == 0 or curr_total == 0:
            # 该科没有有效排名，跳过
            continue

        # 以 key_cols 作为索引进行对齐（视情况为 姓名 或 (班级, 姓名)）
        prev_s = prev_df.set_index(key_cols)[prev_rank_cols[subj]]
        curr_s = curr_df.set_index(key_cols)[curr_rank_cols[subj]]

        merged = pd.concat(
            [prev_s.rename("上次排名"), curr_s.rename("本次排名")],
            axis=1,
        )

        # 去掉任一考试缺失的学生
        merged = merged.dropna(subset=["上次排名", "本次排名"])

        if merged.empty:
            continue

        merged["上次总人数"] = prev_total
        merged["本次总人数"] = curr_total
        merged["上次百分比"] = merged["上次排名"] / prev_total
        merged["本次百分比"] = merged["本次排名"] / curr_total
        merged["进退步差值"] = merged["上次百分比"] - merged["本次百分比"]

        def judge_direction(delta: float) -> str:
            # delta = 上次百分比 - 本次百分比
            # 百分比越大越好，因此：
            #  delta >= -threshold -> 进步
            #  delta <= +threshold -> 退步
            if delta >= -threshold:
                return f"进步>{int(threshold*100)}%"
            elif delta <= threshold:
                return f"退步>{int(threshold*100)}%"
            else:
                return ""

        merged["变化方向"] = merged["进退步差值"].apply(judge_direction)
        merged["科目"] = subj

        merged = merged.reset_index()  # 把 key_cols 从索引还原为列

        # 统一整理输出列
        use_cols = (
            key_cols
            + [
                "科目",
                "上次排名",
                "上次总人数",
                "上次百分比",
                "本次排名",
                "本次总人数",
                "本次百分比",
                "进退步差值",
                "变化方向",
            ]
        )
        records.append(merged[use_cols])

    if not records:
        raise ValueError(f"{label_prefix} 在所有科目中都没有可比对且有记录的学生。")

    result_df = pd.concat(records, ignore_index=True)

    # 只保留进步/退步明显的记录
    significant_df = result_df[result_df["变化方向"] != ""]

    # ===================== 5. 写出结果 =====================
    output_path = Path(output_path)
    with pd.ExcelWriter(output_path) as writer:
        counts_df.to_excel(writer, sheet_name="全年级各科人数", index=False)

        sheet_all = f"{label_prefix}全部学生百分比"
        sheet_sig = f"{label_prefix}进退步>{int(threshold*100)}%"

        result_df.to_excel(writer, sheet_name=sheet_all, index=False)
        significant_df.to_excel(writer, sheet_name=sheet_sig, index=False)

    print(f"分析完成，结果已保存至：{output_path.resolve()}")


def main():
    parser = argparse.ArgumentParser(
        description="对两次考试（如月考 vs 期中）的年级排名进行进退步分析。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    # 互斥：单文件模式 vs 双文件模式
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--file",
        help="单个 Excel 文件，包含两次考试数据（用不同 sheet 表示）。",
    )
    group.add_argument(
        "--file-prev",
        help="上次考试 Excel 文件（与 --file-curr 联用）。",
    )

    parser.add_argument(
        "--file-curr",
        help="本次考试 Excel 文件（与 --file-prev 联用）。如果使用 --file，则可忽略。",
    )

    parser.add_argument(
        "--sheet-prev",
        default="月考",
        help="上次考试所在的 sheet 名。",
    )
    parser.add_argument(
        "--sheet-curr",
        default="期中",
        help="本次考试所在的 sheet 名。",
    )

    parser.add_argument(
        "--class",
        dest="class_name",
        default=None,
        help="要分析的班级名称，例如 '6班'。若不提供则分析全部班级。",
    )

    parser.add_argument(
        "--output",
        default="成绩进退步分析.xlsx",
        help="输出结果 Excel 文件名。",
    )

    parser.add_argument(
        "--threshold",
        type=float,
        default=0.2,
        help="判定进步/退步的百分比阈值(0.2)。",
    )

    args = parser.parse_args()

    # 读取数据
    if args.file:
        # 单文件 + 双 sheet 模式
        file_path = Path(args.file)
        if not file_path.exists():
            raise FileNotFoundError(f"找不到文件：{file_path.resolve()}")

        exam_prev = pd.read_excel(file_path, sheet_name=args.sheet_prev)
        exam_curr = pd.read_excel(file_path, sheet_name=args.sheet_curr)
    else:
        # 双文件模式
        if not args.file_prev or not args.file_curr:
            parser.error("使用 --file-prev 时必须同时提供 --file-curr。")

        file_prev = Path(args.file_prev)
        file_curr = Path(args.file_curr)
        if not file_prev.exists():
            raise FileNotFoundError(f"找不到上次考试文件：{file_prev.resolve()}")
        if not file_curr.exists():
            raise FileNotFoundError(f"找不到本次考试文件：{file_curr.resolve()}")

        exam_prev = pd.read_excel(file_prev, sheet_name=args.sheet_prev)
        exam_curr = pd.read_excel(file_curr, sheet_name=args.sheet_curr)

    analyze_scores(
        exam_prev=exam_prev,
        exam_curr=exam_curr,
        output_path=args.output,
        class_name=args.class_name,
        threshold=args.threshold,
    )


if __name__ == "__main__":
    main()
