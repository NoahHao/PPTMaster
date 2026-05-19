#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
页面选择器：将内容块映射为具体模板页面编号
"""

import json
import random
from pathlib import Path
from typing import List, Dict, Optional


class SlideSelector:
    """根据内容角色从模板目录中选择合适的页面"""

    ROLE_TO_TYPE = {
        "封面": "封面-总标题",
        "总述": "内容-KPI/数据",      # 总览用KPI大字报
        "总览": "内容-KPI/数据",
        "数据": "内容-KPI/数据",      # 数据也用KPI
        "分述": "内容-图文卡片",      # 分项用图文卡片
        "问题": "内容-标准文本",      # 问题背景用标准文本
        "方案": "内容-复杂图示",      # 解决方案用图示
        "成效": "内容-KPI/数据",      # 成效用数据
        "总结": "内容-大图+标语",     # 总结用大图+标语
        "展望": "内容-大图+标语",
        "过渡": "内容-大图+标语",
        "背景": "内容-大图+标语",
        "结尾": "内容-大图+标语",
        "分项数据": "内容-图文卡片",
        "洞察建议": "内容-文本详细",
        "核心数据": "内容-KPI/数据",
        "数据趋势": "内容-图表型",
    }

    # 备选映射（当首选页型不够时）
    FALLBACK_TYPE = {
        "内容-KPI/数据": ["内容-标准文本", "内容-图文卡片"],
        "内容-图文卡片": ["内容-标准文本", "内容-复杂图示"],
        "内容-复杂图示": ["内容-图文卡片", "内容-标准文本"],
        "内容-标准文本": ["内容-图文卡片", "内容-文本详细"],
        "内容-大图+标语": ["内容-图文卡片", "内容-KPI/数据"],
        "内容-图表型": ["内容-KPI/数据", "内容-标准文本"],
        "内容-表格型": ["内容-标准文本", "内容-图文卡片"],
        "内容-文本详细": ["内容-标准文本", "内容-图文卡片"],
    }

    def __init__(self, catalog_path: str = None):
        if catalog_path is None:
            catalog_path = Path(__file__).parent.parent / "references" / "template_catalog_v2.json"
        
        with open(catalog_path, "r", encoding="utf-8") as f:
            self.catalog = json.load(f)
        
        # 为每种页型建立可用池
        self._pools = {}
        for ptype, info in self.catalog.items():
            self._pools[ptype] = list(info["slides"])  # 深拷贝
            # 随机打乱，保证每次选取不同
            random.shuffle(self._pools[ptype])
        
        # 使用计数器（避免同一页被重复选用）
        self._used = set()

    def select_for_block(self, role: str, title: str = "", 
                          body: str = "", data_items: List = None) -> Optional[Dict]:
        """
        为一个内容块选择最佳模板页
        
        Returns:
            {"slide": 页码, "type": "页型", "layout": "布局名", ...}
        """
        preferred = self.ROLE_TO_TYPE.get(role, "内容-标准文本")
        
        # 特殊情况：有数据条目的块用KPI页
        if data_items and len(data_items) >= 2:
            preferred = "内容-KPI/数据"
        
        # 特殊情况：极长文本用文本详细页
        if body and len(body) > 300:
            preferred = "内容-文本详细"

        # 尝试首选页型
        slide_info = self._pick_from_pool(preferred)
        if slide_info:
            return slide_info

        # 尝试备选页型
        fallbacks = self.FALLBACK_TYPE.get(preferred, ["内容-标准文本"])
        for fb in fallbacks:
            slide_info = self._pick_from_pool(fb)
            if slide_info:
                return slide_info

        # 最终兜底：任何可用页
        for ptype, slides in self._pools.items():
            if slides and ptype not in ["说明页"]:
                slide_info = slides.pop(0)
                if slide_info["slide"] not in self._used:
                    self._used.add(slide_info["slide"])
                    slide_info["type"] = ptype
                    return slide_info

        return None

    def select_sequence(self, blocks: List[Dict], structure: str) -> List[Dict]:
        """
        为一组内容块按结构模式选择模板页序列

        Args:
            blocks: 内容块列表 (来自PromptAnalyzer的输出)
            structure: 结构模式名 (如"总分总")

        Returns:
            [{"content": block_dict, "template": slide_info}, ...]
        """
        # 加载结构模板
        struct_path = Path(__file__).parent / "structure_templates.json"
        with open(struct_path, "r", encoding="utf-8") as f:
            struct_templates = json.load(f)

        mode = struct_templates["modes"].get(structure)
        if not mode:
            mode = struct_templates["modes"]["总分总"]  # 默认

        sequence = []
        
        # 展平所有内容块
        flat_blocks = self._flatten_blocks(blocks)

        # 按结构模板的slot分配
        template_slots = mode["sequence"]
        slot_pool = list(template_slots)  # 可消费的slot池
        slot_usage = {s["role"]: 0 for s in template_slots}
        
        # 预先计算每个slot的最大使用次数
        slot_max = {}
        for s in template_slots:
            c = s["count"]
            if c == "*":
                slot_max[s["role"]] = 999
            elif isinstance(c, str) and "-" in c:
                slot_max[s["role"]] = int(c.split("-")[1])
            else:
                slot_max[s["role"]] = int(c)
        
        for block in flat_blocks:
            role = block.get("role", "分述")
            
            # 找匹配的slot
            matched_slot = None
            for s in template_slots:
                if s["role"] == role and slot_usage[s["role"]] < slot_max.get(s["role"], 1):
                    matched_slot = s
                    break
            
            if not matched_slot:
                # 找任何未满的slot（优先选择分述/内容型）
                for s in template_slots:
                    if slot_usage.get(s["role"], 0) < slot_max.get(s["role"], 1):
                        if s["role"] in ("分述", "分项数据", "洞察建议"):
                            matched_slot = s
                            break
                # 还没找到就用第一个未满的
                if not matched_slot:
                    for s in template_slots:
                        if slot_usage.get(s["role"], 0) < slot_max.get(s["role"], 1):
                            matched_slot = s
                            break
            
            if matched_slot:
                slot_usage[matched_slot["role"]] += 1
                matched_role = matched_slot["role"]
            else:
                matched_role = role  # fallback
            
            slide_info = self.select_for_block(
                role=role,
                title=block.get("title", ""),
                body=block.get("body", ""),
                data_items=block.get("data_items", [])
            )
            
            if slide_info:
                sequence.append({
                    "content": block,
                    "template": slide_info,
                    "slot_role": matched_role
                })

        return sequence

    def _flatten_blocks(self, blocks: List[Dict]) -> List[Dict]:
        """展平嵌套内容块为线性序列"""
        result = []
        for block in blocks:
            result.append(block)
            if block.get("children"):
                result.extend(self._flatten_blocks(block["children"]))
        return result

    def _pick_from_pool(self, ptype: str) -> Optional[Dict]:
        """从页型池中取一个未使用的页面"""
        if ptype not in self._pools:
            return None
        
        pool = self._pools[ptype]
        original_len = len(pool)
        
        for _ in range(original_len):
            if not pool:
                return None
            slide_info = pool.pop(0)
            if slide_info["slide"] not in self._used:
                self._used.add(slide_info["slide"])
                slide_info["type"] = ptype
                return slide_info
            # 已用过就放回队尾
            pool.append(slide_info)
        
        return None

    def reset(self):
        """重置使用记录和页池"""
        self._used = set()
        for ptype in self._pools:
            random.shuffle(self._pools[ptype])


# ===== 测试 =====
if __name__ == "__main__":
    from prompt_analyzer import PromptAnalyzer
    
    analyzer = PromptAnalyzer()
    selector = SlideSelector()
    
    test = """
# Q1工作汇报

## 总体业绩
营收 5000万，增长35%
客户 200家
团队 50人

## 核心项目
### 智能运维平台
完成第一阶段，上线3客户

### 数据分析引擎
性能提升50%

## 问题与挑战
招聘滞后，模块延期

## 下季规划
推进商业化，启动预研

## 总结
达成率85%，超预期
"""
    
    result = analyzer.parse(test)
    print(f"识别结构: {result['structure']}")
    print(f"内容块数: {result['stats']['blocks']}")
    
    sequence = selector.select_sequence(result["blocks"], result["structure"])
    
    print(f"\n选中的模板页序列 ({len(sequence)}页):")
    for i, item in enumerate(sequence):
        c = item["content"]
        t = item["template"]
        print(f"  {i+1}. [{t['type']}] Slide {t['slide']} ← \"{c.get('title','')[:50]}\" ({item['slot_role']})")
