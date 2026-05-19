#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Prompt分析器：解析用户输入提纲，输出结构化内容树
支持Markdown格式和纯文本输入
"""

import re
import json
from typing import List, Dict, Optional


class ContentBlock:
    """内容块"""
    def __init__(self, role: str, title: str, body: str = "", 
                 level: int = 0, data_items: List[Dict] = None):
        self.role = role          # 语义角色: 总述/分述/总结/数据/过渡/封面
        self.title = title        # 块标题
        self.body = body          # 正文内容
        self.level = level        # 层级 (0=一级, 1=二级, 2=三级)
        self.data_items = data_items or []  # 数据条目 [{"label":"", "value":""}]
        self.children = []        # 子内容块

    def to_dict(self):
        return {
            "role": self.role,
            "title": self.title,
            "body": self.body,
            "level": self.level,
            "data_items": self.data_items,
            "children": [c.to_dict() for c in self.children]
        }

    def __repr__(self):
        return f"Block({self.role}, '{self.title[:30]}', level={self.level}, children={len(self.children)})"


class PromptAnalyzer:
    """分析用户输入的提纲文本，提取结构化内容"""

    # 角色关键词映射
    ROLE_KEYWORDS = {
        "总述": ["概述", "总览", "概要", "全景", "全貌", "总体", "整体", "一览", "概览"],
        "数据": ["数据", "指标", "数字", "KPI", "业绩", "成果", "统计", "占比", "增长率"],
        "问题": ["问题", "挑战", "痛点", "难点", "瓶颈", "困境", "风险"],
        "方案": ["方案", "解决", "对策", "措施", "策略", "规划", "计划", "路径"],
        "成效": ["成效", "效果", "价值", "收益", "成果", "回报", "提升"],
        "展望": ["展望", "未来", "规划", "愿景", "目标", "下一步", "方向"],
        "总结": ["总结", "回顾", "结语", "结论", "小结", "汇总"],
    }

    # 结构模式关键词
    PATTERN_KEYWORDS = {
        "总分总": ["首先", "其次", "最后", "总结", "概述", "分述", "综上"],
        "总分": ["首先", "其次", "最后", "概述", "展开"],
        "并列": ["一方面", "另一方面", "第一", "第二", "第三", "同时"],
        "问题方案": ["问题", "方案", "解决", "痛点", "对策"],
        "数据驱动": ["数据", "指标", "趋势", "分析", "洞察"],
        "叙事弧线": ["背景", "挑战", "转折", "突破", "展望"],
        "三步递进": ["是什么", "为什么", "怎么做", "定义", "原因", "方法"],
    }

    def __init__(self):
        pass

    def parse(self, text: str) -> Dict:
        """
        解析用户输入，返回结构化内容树
        
        Args:
            text: 用户输入的提纲文本(Markdown或纯文本)
        
        Returns:
            {
                "structure": "总分总",     # 推荐的结构模式
                "title": "主标题",
                "blocks": [ContentBlock...],  # 顶层内容块
                "stats": {"blocks": N, "depth": N, "data_count": N}
            }
        """
        text = text.strip()
        if not text:
            return {"structure": "总分总", "title": "", "blocks": [], "stats": {}}

        # Step 1: 解析Markdown层级
        blocks = self._parse_markdown_blocks(text)
        
        # Step 2: 识别每个块的语义角色
        for block in blocks:
            self._assign_role(block)
        
        # Step 3: 识别数据条目
        for block in blocks:
            self._extract_data_items(block)

        # Step 4: 推断整体结构模式
        structure = self._infer_structure(blocks, text)

        # 提取总标题
        title = blocks[0].title if blocks else ""

        # 统计
        def count_all(blist):
            total = len(blist)
            for b in blist:
                total += count_all(b.children)
            return total
        
        total_blocks = count_all(blocks)
        max_depth = max((b.level for b in blocks), default=0)
        data_count = sum(1 for b in blocks if b.role == "数据" or b.data_items)
        
        return {
            "structure": structure,
            "title": title,
            "blocks": [b.to_dict() for b in blocks],
            "stats": {
                "blocks": total_blocks,
                "depth": max_depth,
                "data_count": data_count,
            }
        }

    def _parse_markdown_blocks(self, text: str) -> List[ContentBlock]:
        """解析Markdown格式，识别 #/##/### 层级"""
        lines = text.split("\n")
        blocks = []
        current_block = None
        stack = []  # 层级栈: [(level, block)]
        first_block = True

        for line in lines:
            line = line.rstrip()
            if not line:
                continue

            # 检测标题层级
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if heading_match:
                level = len(heading_match.group(1)) - 1  # # → level 0, ## → level 1
                title = heading_match.group(2).strip()
                
                # 第一个一级标题 = 封面
                role = "封面" if (first_block and level == 0) else "分述"
                block = ContentBlock(role=role, title=title, level=level)
                first_block = False
                
                # 找到父块
                while stack and stack[-1][0] >= level:
                    stack.pop()
                
                if stack:
                    stack[-1][1].children.append(block)
                else:
                    blocks.append(block)
                
                stack.append((level, block))
                current_block = block
                continue

            # 检测列表项
            list_match = re.match(r'^[\s]*[-*+]\s+(.+)$', line)
            if list_match and current_block:
                item_text = list_match.group(1)
                current_block.body += (item_text + "\n") if current_block.body else item_text
                continue

            # 检测编号列表
            num_match = re.match(r'^[\s]*(\d+)[\.\)、]\s+(.+)$', line)
            if num_match and current_block:
                current_block.body += (num_match.group(2) + "\n") if current_block.body else num_match.group(2)
                continue

            # 检测键值对数据 (如: "营收: 100亿" 或 "营收 100亿")
            kv_match = re.match(r'^[\s]*([^:：\d]+)[：:]\s*(.+)$', line)
            if kv_match:
                label = kv_match.group(1).strip()
                value = kv_match.group(2).strip()
                if current_block:
                    current_block.body += f"{label}: {value}\n" if current_block.body else f"{label}: {value}"
                else:
                    # 没有当前块的游离文本
                    pass
                continue

            # 普通段落文本 - 追加到当前块
            if current_block:
                current_block.body += (line + "\n") if current_block.body else line
            elif line.strip():
                # 游离文本创建新块
                block = ContentBlock(role="分述", title=line[:50], body=line, level=0)
                blocks.append(block)
                current_block = block

        return blocks

    def _assign_role(self, block: ContentBlock):
        """根据标题和内容关键词分配语义角色"""
        full_text = block.title + " " + (block.body or "")
        
        # 数字+指标→数据角色
        has_numbers = bool(re.search(r'[\d,.]+\s*[万亿千百%个项次元]', full_text))
        has_data_kw = any(kw in block.title for kw in ["业绩", "指标", "数据", "成果", "KPI", "增长", "营收"])

        # 如果已有"封面"角色，保持不变
        if block.role == "封面":
            for child in block.children:
                self._assign_role(child)
            return

        for role, keywords in self.ROLE_KEYWORDS.items():
            for kw in keywords:
                if kw in block.title or (kw in full_text and block.level == 0):
                    block.role = role
                    break
            if block.role != "分述":
                break
        
        # 有大量数字的块→数据 (覆盖"总述")
        if block.role in ("分述", "总述") and (has_numbers or has_data_kw):
            block.role = "数据"

        # 递归处理子块
        for child in block.children:
            self._assign_role(child)

            # 如果子块有明确角色，父块设为"总述"（除非父块也是明确的）
            if child.role != "分述" and block.role == "分述":
                block.role = "总述"

    def _extract_data_items(self, block: ContentBlock):
        """从正文中提取数据条目"""
        if not block.body:
            return

        # 匹配 "标签: 数值" 或 "标签 数值"
        patterns = [
            r'([^:：\n]+)[：:]\s*([^\n]+)',   # 中文冒号/英文冒号
            r'([^\d\n]+?)\s+([\d,.]+[万亿千百%个项次元\d]*)',  # 数字+单位
        ]

        for pattern in patterns:
            matches = re.findall(pattern, block.body)
            for label, value in matches:
                label = label.strip().rstrip("，,。.")
                value = value.strip()
                if len(label) <= 20 and len(label) >= 2:
                    block.data_items.append({"label": label, "value": value})

        # 递归子块
        for child in block.children:
            self._extract_data_items(child)

    def _infer_structure(self, blocks: List[ContentBlock], original_text: str) -> str:
        """推断整体结构模式"""
        roles = [b.role for b in blocks]
        
        # 统计各模式关键词命中次数
        scores = {}
        for pattern, keywords in self.PATTERN_KEYWORDS.items():
            score = sum(original_text.count(kw) for kw in keywords)
            scores[pattern] = score

        # 基于角色分布调整
        has_data = "数据" in roles
        has_problem = any("问题" in b.role for b in blocks)
        has_solution = any("方案" in b.role for b in blocks)
        has_outlook = any("展望" in b.role for b in blocks)
        has_summary = any("总结" in b.role for b in blocks)
        
        if has_data:
            scores["数据驱动"] += 3
        if has_problem and has_solution:
            scores["问题方案"] += 3
        if has_summary:
            scores["总分总"] += 2
        if has_outlook:
            scores["叙事弧线"] += 2

        # 默认逻辑
        if len(blocks) <= 2:
            scores["总分"] += 2
        if len(blocks) >= 5 and all(r == "分述" for r in roles):
            scores["并列"] += 3

        # 选最高分
        best = max(scores, key=scores.get)
        if scores[best] == 0:
            return "总分总"  # 默认fallback
        
        return best


# ===== 测试 =====
if __name__ == "__main__":
    analyzer = PromptAnalyzer()
    
    test_input = """
# 2026年Q1工作汇报

## 总体业绩概览
- 营收达到 5000万，同比增长35%
- 客户数量突破200家
- 团队扩展至50人

## 核心项目进展
### 项目A：智能运维平台
- 完成第一阶段开发
- 上线3个客户，反馈良好

### 项目B：数据分析引擎
- 性能提升50%
- 完成POC验证

## 问题与挑战
- 人员招聘进度滞后
- 项目A的第三个模块延期

## 下季度规划
- 重点推进项目A商业化
- 启动项目C的技术预研

## 总结
本季度整体达成率85%，核心指标超预期。
"""
    
    result = analyzer.parse(test_input)
    print(json.dumps(result, ensure_ascii=False, indent=2))
