#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
append_research.py — 通用用户访谈 CSV 横向追加脚本

【用法】
  python3 append_research.py <interview_data.json> --topic "<调研主题>"

【核心原则】
- CSV 横向格式：字段结构纵向（A列），每位访谈对象占一列（B、C、D...）
- 首次创建包含完整 A 列字段名和分类标题
- 后续访谈自动追加到已有文件的新列
- UTF-8 BOM 编码，Excel 双击打开不乱码
- 输出目录：`{workspace}/{主题}访谈记录/`

【字段结构】（共27行，与原模板行号对应）
  行1:   姓名          (A列空白)
  行2:   基本画像      (A列空白)
  行3:   访谈状态      (A列空白)
  行4:   待办任务
  行5:   现有状态
  行6:   [分类] 改变的力量
  行7:   触发事件（Triggering Event）
  行8:   转换触发事件（Conversion Event）
  行9:   期望结果
  行10:  利益攸关的问题
  行11:  [分类] 吸引力
  行12:  客户纳入考虑范围的解决方案
  行13:  最终选择的解决方案
  行14:  独特价值主张（承诺的结果）
  行15:  预期结果
  行16:  [分类] 改变现状所面对的阻力
  行17:  现有解决方案（如果没有则不填）
  行18:  哪里出了问题？
  行19:  发生转换的障碍
  行20:  [分类] 使用最终选择的解决方案所面对的阻力
  行21:  客户的焦虑
  行22:  采用解决方案的障碍
  行23:  其他细节（如果没有则不填）
  行24:  [分类] 下一步计划
  行25:  时间轴
  行26:  实际结果
  行27:  下一步行动
"""

import sys
import json
import os
import csv
import argparse
from pathlib import Path


# ============ 字段定义（27行） ============
# 格式：(行号, A列标签(字段名/分类标题/None), JSON key/None)
FIELD_DEFS = [
    (1,  None,              "姓名"),
    (2,  None,              "基本画像"),
    (3,  None,              "访谈状态"),
    (4,  "待办任务",         "待办任务"),
    (5,  "现有状态",         "现有状态"),
    (6,  "[分类] 改变的力量", None),
    (7,  "触发事件（Triggering Event）", "触发事件"),
    (8,  "转换触发事件（Conversion Event）", "转换触发事件"),
    (9,  "期望结果",         "期望结果"),
    (10, "利益攸关的问题",   "利益攸关的问题"),
    (11, "[分类] 吸引力",   None),
    (12, "客户纳入考虑范围的解决方案", "客户纳入考虑范围的解决方案"),
    (13, "最终选择的解决方案", "最终选择的解决方案"),
    (14, "独特价值主张（承诺的结果）", "独特价值主张"),
    (15, "预期结果",         "预期结果"),
    (16, "[分类] 改变现状所面对的阻力", None),
    (17, "现有解决方案（如果没有则不填）", "现有解决方案"),
    (18, "哪里出了问题？",   "哪里出了问题"),
    (19, "发生转换的障碍",   "发生转换的障碍"),
    (20, "[分类] 使用最终选择的解决方案所面对的阻力", None),
    (21, "客户的焦虑",       "客户的焦虑"),
    (22, "采用解决方案的障碍", "采用解决方案的障碍"),
    (23, "其他细节（如果没有则不填）", "其他细节"),
    (24, "[分类] 下一步计划", None),
    (25, "时间轴",           "时间轴"),
    (26, "实际结果",         "实际结果"),
    (27, "下一步行动",       "下一步行动"),
]

CSV_ROW_COUNT = len(FIELD_DEFS)


def load_interview_data(json_path: str) -> dict:
    """加载 JSON 访谈数据"""
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)


def read_existing_csv(csv_path: Path) -> list:
    """读取已有 CSV，返回二维列表。不存在则返回 None。"""
    if not csv_path.exists():
        return None
    rows = []
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        for row in reader:
            rows.append(row)
    return rows


def find_next_col(rows: list) -> int:
    """找到第一个空列的索引（0-based）"""
    if not rows:
        return 1  # B列（索引1）
    # 检查第一行（姓名行），找第一个空列
    for col in range(1, 200):  # 上限200列
        if col >= len(rows[0]) or rows[0][col].strip() == "":
            return col
    return len(rows[0])


def pad_rows(rows: list, target_rows: int, target_cols: int) -> list:
    """确保 rows 有 target_rows 行、每行至少 target_cols 列"""
    while len(rows) < target_rows:
        rows.append([])
    for i in range(len(rows)):
        while len(rows[i]) < target_cols:
            rows[i].append("")
    return rows


def init_csv(data: dict) -> list:
    """创建新 CSV 数据：A列=字段名/分类标题，B列=访谈数据"""
    rows = []
    for (_, a_label, b_key) in FIELD_DEFS:
        row = [a_label or "", data.get(b_key, "") if b_key else ""]
        rows.append(row)
    return rows


def append_to_csv(rows: list, col: int, data: dict) -> list:
    """在指定列追加一位访谈对象的数据"""
    for i, (_, _, b_key) in enumerate(FIELD_DEFS):
        # 确保行存在
        while len(rows) <= i:
            rows.append([])
        # 确保列存在
        while len(rows[i]) <= col:
            rows[i].append("")
        if b_key:
            rows[i][col] = data.get(b_key, "")
        else:
            rows[i][col] = ""
    return rows


def write_csv(csv_path: Path, rows: list):
    """写入 CSV（UTF-8 BOM）"""
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        for row in rows:
            writer.writerow(row)


def get_workspace() -> Path:
    """获取 workspace 路径"""
    workspace = os.environ.get("WORKBUDDY_WORKSPACE")
    if workspace:
        return Path(workspace)
    # 默认取当前对话的 workspace
    return Path.home() / "WorkBuddy" / "20260416085305"


def main():
    parser = argparse.ArgumentParser(description="通用用户访谈 CSV 追加脚本")
    parser.add_argument("json_path", help="访谈数据 JSON 文件路径")
    parser.add_argument("--topic", required=True, help="调研主题（如：Vibe Coding、GEO写作）")
    args = parser.parse_args()

    json_path = Path(args.json_path)
    if not json_path.exists():
        print(f"❌ JSON 文件不存在: {json_path}")
        sys.exit(1)

    data = load_interview_data(json_path)
    name = data.get("姓名", "未命名")

    # 动态获取 workspace 和输出目录
    workspace = get_workspace()
    output_dir = workspace / f"{args.topic}访谈记录"
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_path = output_dir / f"{args.topic}访谈汇总.csv"

    # 读取或初始化
    rows = read_existing_csv(csv_path)

    if rows is None:
        # 首次创建
        rows = init_csv(data)
        print(f"✅ 已创建新 CSV → {csv_path}")
    else:
        # 追加
        col = find_next_col(rows)
        rows = pad_rows(rows, CSV_ROW_COUNT, col + 1)
        rows = append_to_csv(rows, col, data)
        col_letter = chr(ord("A") + col)
        print(f"✅ 已追加到列 {col_letter}（{name}）→ {csv_path}")

    write_csv(csv_path, rows)
    print(f"📊 CSV 已保存（{len(rows)} 行）")


if __name__ == "__main__":
    main()
