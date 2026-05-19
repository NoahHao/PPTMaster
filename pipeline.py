#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPT智能模板助手 - 总控流水线
一键执行：分析输入 → 选模板 → 填内容 → 合并输出
"""

import sys
import json
import os
from pathlib import Path

# 添加当前目录到path
sys.path.insert(0, str(Path(__file__).parent))

from prompt_analyzer import PromptAnalyzer
from slide_selector import SlideSelector
from merge_ppt import SmartMerger


def run_pipeline(
    user_input: str,
    output_path: str = "output.pptx",
    structure: str = None,  # 可手动指定结构模式
    verbose: bool = True
) -> dict:
    """
    执行完整流水线

    Args:
        user_input: 用户输入的提纲文本
        output_path: 输出PPT文件路径
        structure: 手动指定结构模式（不指定则自动推断）
        verbose: 是否打印详细日志

    Returns:
        {"output": "output.pptx", "slides": N, "structure": "...", "sequence": [...]}
    """
    if verbose:
        print("=" * 60)
        print("PPT智能模板助手 v1.0")
        print("=" * 60)

    # === Step 1: 分析用户输入 ===
    if verbose:
        print("\n[Step 1] 分析输入内容...")
    
    analyzer = PromptAnalyzer()
    analysis = analyzer.parse(user_input)
    
    if structure:
        analysis["structure"] = structure
    
    if verbose:
        print(f"  识别结构: {analysis['structure']}")
        print(f"  标题: {analysis['title']}")
        print(f"  内容块: {analysis['stats']['blocks']}个")
        print(f"  层级深度: {analysis['stats']['depth']}")
        print(f"  数据条目: {analysis['stats']['data_count']}")

    # === Step 2: 选择模板页 ===
    if verbose:
        print("\n[Step 2] 选择模板页面...")
    
    selector = SlideSelector()
    sequence = selector.select_sequence(
        analysis["blocks"],
        analysis["structure"]
    )
    
    if verbose:
        print(f"  选中 {len(sequence)} 页:")
        for i, item in enumerate(sequence):
            c = item["content"]
            t = item["template"]
            print(f"    {i+1}. Slide {t['slide']:3d} [{t['type']:15s}] ← {c.get('title','')[:50]} [{item['slot_role']}]")

    # === Step 3: 克隆+填充+合并 ===
    if verbose:
        print(f"\n[Step 3] 合并输出到 {output_path} ...")
    
    merger = SmartMerger()
    result_path = merger.merge_with_content(
        sequence,
        output_path,
        title=analysis.get("title", ""),
        include_cover=True
    )
    
    slide_count = sequence.__len__()
    
    if verbose:
        print(f"\n{'=' * 60}")
        print(f"完成! 输出: {result_path}")
        print(f"总页数: {slide_count + 2} (含封面+结尾)")
        print(f"{'=' * 60}")

    return {
        "output": result_path,
        "slides": slide_count + 2,
        "structure": analysis["structure"],
        "sequence": [
            {
                "slide": item["template"]["slide"],
                "type": item["template"]["type"],
                "title": item["content"].get("title", ""),
                "role": item["slot_role"],
            }
            for item in sequence
        ]
    }


# ===== CLI入口 =====
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python pipeline.py <输入文件.md> [输出文件.pptx] [结构模式]")
        print("\n结构模式: 总分总 / 总分 / 并列 / 问题方案 / 数据驱动 / 叙事弧线 / 三步递进")
        print("\n示例:")
        print("  python pipeline.py 提纲.md output.pptx")
        print("  python pipeline.py 提纲.md output.pptx 总分总")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else "output.pptx"
    structure = sys.argv[3] if len(sys.argv) > 3 else None

    # 读取输入文件
    if not os.path.exists(input_file):
        print(f"错误: 文件不存在: {input_file}")
        sys.exit(1)

    with open(input_file, "r", encoding="utf-8") as f:
        content = f.read()

    result = run_pipeline(content, output_file, structure)
    
    print(f"\n最终输出: {result['output']}")
    print(f"结构模式: {result['structure']}")
    print(f"使用模板页: {[s['slide'] for s in result['sequence']]}")
