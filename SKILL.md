---
name: hw-slide-master-skill
description: >
  华为胶片大师。基于149页公司汇报提纲精选模板，自动分析用户内容结构，
  选取最匹配模板页，按逻辑结构替换文字内容，完全保留原有布局、字体、字号、颜色等一切格式。
  支持单页/多页生成，纯文字替换模式。TRIGGERS: 华为胶片大师、华为胶片、华为PPT、汇报模板、公司汇报、
  模板替换、幻灯片填充、PPT内容替换、逻辑结构匹配、槽位匹配、智能换字、模板目录、模板页选取、
  字数对齐替换、纯文字替换不碰格式、自定义模板、模板分析。
---

# 华为胶片大师

> 基于模板DNA的PPT智能填充引擎——分析模板结构 → 匹配内容逻辑 → 纯文字替换生成。

---

## 一、功能概述

### 1.1 核心能力

本技能将任意 PPTX 模板转化为**可编程的"槽位填充系统"**：用户提供模板（或使用内置华为模板149页）和内容描述，引擎自动完成结构分析、页面匹配、逐框文字替换，输出格式化 PPT。

| 能力 | 说明 |
|------|------|
| **模板无关** | 默认华为模板149页或用户自提供任意PPTX，统一分析流水线 |
| **零格式破坏** | 纯文本替换，不动 `<a:rPr>` 字体属性、不动 `<a:xfrm>` 坐标 |
| **逻辑匹配** | 按字号/粗体/字数/位置推断文本框角色，同角色内容对齐分配 |
| **智能蒸馏** | 三阶策略将任意长度文本适配到任意宽度槽位，永不在词语中间截断 |
| **满槽输出** | 轮循机制保证零空槽，兜底池防止内容枯竭 |

### 1.2 适用场景

- 企业汇报：封面→目录→KPI数据→分项论述→总结，全套骨架填充
- 主题展示：单页或多页，将自定义内容灌入预设排版
- 模板复用：把旧模板的文字换成新内容，保留所有设计与格式

### 1.3 技术边界

- **不生成新布局**：不创建形状、不调整位置、不修改配色。排版能力完全由模板决定
- **不操作图片/图表**：仅处理文本框架（`has_text_frame`），图片和SmartArt保留原样
- **不改变幻灯片数量**：如需多页，从模板选取多页后合并

---

## 二、初始化（每次会话强制执行）

**在回答任何PPT生成需求之前，必须先完成初始化选择。不允许跳过此步骤直接生成。**

向用户提出以下选择（使用原文，不可改写）：

> 请选择模板来源：
> 
> **1. 使用默认华为模板** — `template/hw_template.pptx`（149页·6种结构·157条DNA·35MB）
> 
> **2. 提供自定义模板** — 上传你的 PPTX 文件，我将完整分析其结构，生成模板目录后使用

### 选项1：默认模板
- 直接加载 `references/template_catalog_v2.json`
- 无需额外分析，立即就绪

### 选项2：自定义模板

**步骤A — 接收文件：**
要求用户提供 `.pptx` 文件路径。

**步骤B — 执行分析：**
```bash
cd skills/hw-slide-master-skill
python scripts/analyze_template.py \
  --template <用户PPT路径> \
  --output references/custom_template_catalog.json
```

**步骤C — 分析内容：**
1. **逐页DNA提取** — 每个文本框提取17个属性（见[模板DNA系统](#五模板dna系统)）
2. **结构类型推断** — 7级优先级链式分类器（见[结构类型分类器](#三结构类型分类器)）
3. **输出Catalog JSON** — 保存到 `references/custom_template_catalog.json`
4. **打印统计摘要** — 页数、结构分布柱状图、槽位范围、字数范围

**步骤D — 后续使用：**
所有生成操作使用自定义模板目录。模板PPT路径和Catalog JSON路径作为上下文传递。

### 持久化规则

| 来源 | 模板PPT | Catalog JSON |
|------|---------|-------------|
| 默认华为模板 | `template/hw_template.pptx` | `references/template_catalog_v2.json` |
| 用户自定义模板 | 用户提供路径 | `references/custom_template_catalog.json` |

> **⚠️ 如果用户跳过初始化直接提出生成需求，必须先拦截、追问选择，再继续。**

---

## 三、架构总览

### 3.1 组件关系图

```
┌────────────────────────────────────────────────────────────────┐
│                        SKILL 入口                              │
│   SKILL.md  →  初始化选择  →  工作流触发                        │
└──────────┬─────────────────────────────────────────────────────┘
           │
    ┌──────┴──────────────────────────────────────────┐
    │         analyze_template.py（分析引擎）           │
    │  · get_slide_dna()    — 逐页17属性提取            │
    │  · classify_structure_type() — 7级分类器          │
    │  · → template_catalog_v2.json / custom_catalog   │
    └──────┬──────────────────────────────────────────┘
           │
           ▼
    ┌──────────────────────────────────────────────────┐
    │              template_utils.py（工具库）          │
    │  · replace_text_keep_font()  纯文本替换          │
    │  · distill()                 三阶蒸馏            │
    │  · classify_slot_role()      角色推断            │
    │  · load_template_catalog()   加载目录JSON        │
    │  · extract_single_slide()    单页提取            │
    └──────┬──────────────────────────────────────────┘
           │
    ┌──────┴──────────────────────────────────────────┐
    │              生成流水线（脚本层）                 │
    │                                                  │
    │  prompt_analyzer.py  → 分析用户意图              │
    │  slide_selector.py   → 从Catalog选最佳模板页     │
    │  gen_logic_match.py  → 逻辑结构匹配+逐框替换     │
    │  content_filler.py   → 批量内容填充              │
    │  merge_ppt.py        → 合并多页输出              │
    └──────────────────────────────────────────────────┘
```

### 3.2 数据流

```
PPTX模板            →  analyze_template.py  →  Catalog JSON（模板DNA库）
                                                      │
用户内容描述         →  prompt_analyzer.py   →  内容结构（页×角色×内容）  
                                                      │
Catalog + 内容结构   →  slide_selector.py     →  选定模板页列表
                                                      │
模板页 + 内容块      →  gen_logic_match.py    →  角色对齐 + 蒸馏 + 替换
                                                      │
多页PPT             →  merge_ppt.py          →  最终输出PPTX
```

### 3.3 核心设计决策

| 决策 | 理由 | 替代方案（已否决） |
|------|------|-------------------|
| **纯文字替换** | 保护模板的设计资产（字体/配色/布局） | Shape复制→会丢失母版/主题 |
| **Catalog预计算** | 一次分析，多次复用；避免每次生成时遍历PPT | 实时分析→149页每次>30秒 |
| **槽宽优先匹配** | 槽位宽度决定内容适配质量，"宽槽多装，窄槽精装" | 结构相似匹配→窄槽截断灾难 |
| **角色分类分配** | 同角色内容对齐，保证语义一致性 | 顺序填充→标题槽填正文 |
| **蒸馏三阶策略** | 不同宽度槽需要不同截断粒度 | 一刀切→短槽破碎词 |

---

## 四、核心设计原则

### 4.1 纯文字替换，格式不动

替换前：
```xml
<a:r><a:rPr lang="zh-CN" sz="2400" b="1"><a:solidFill><a:srgbClr val="C8102E"/>
</a:solidFill></a:rPr><a:t>公司汇报提纲</a:t></a:r>
```

替换后（仅改 `<a:t>` 内容）：
```xml
<a:r><a:rPr lang="zh-CN" sz="2400" b="1"><a:solidFill><a:srgbClr val="C8102E"/>
</a:solidFill></a:rPr><a:t>区域AI工具建设建议</a:t></a:r>
```

- 所有 `<a:rPr>` 属性（字号、字体、颜色、粗斜体、语言）保持不变
- 所有 `<a:xfrm>` 坐标（位置、尺寸）保持不变
- 操作对象：`paragraph.runs[0].text`，通过 `python-pptx` 的 `_element` 访问

### 4.2 永不置空，绝不截尾

```
轮循策略:
  角色池 = [item1, item2, item3, ...]
  对于每页该角色的N个槽:
    slot[i] ← 角色池[i % len(角色池)]
  
  兜底池 (final_fallback):
    如果角色池已耗尽 → 从兜底池取
    兜底池 ≥ 角色数 × 槽数，确保永不枯竭
```

截断规则：
- **绝不在汉字中间截断** — `"区域AI工具建"` → 必须是 `"区域AI工具建设"`
- **绝不用省略号** — `"区域AI工具..."` → 必须是 `"区域AI工具建设建议"`
- 超长内容宁可整句丢弃，也不出半个词

### 4.3 模板选取：槽宽优先

选模板如择友，贵宽不贵多：

```
候选模板排序:
  1. struct_type 匹配权重: 30%（结构是否相似）
  2. avg_slot_len 匹配权重: 50%（每个槽平均能装多少字）
  3. slot_count 匹配权重: 20%（槽位数量是否匹配内容条数）

最终得分 = 0.3×结构分 + 0.5×槽宽分 + 0.2×数量分
```

反例教训：S3（40槽，平均6字，产品矩阵信息图）被误选为结束页 → 碎标签填不满、大空白。正确选择：S29（7槽，平均7字，简洁结束页）。

---

## 五、模板DNA系统

### 5.1 Catalog JSON 结构

```json
{
  "slide": 1,
  "struct_type": "并列结构",
  "confidence": 0.8,
  "description": "并列：6个同等级分项",
  "slot_count": 41,
  "total_chars": 271,
  "text_slots": [
    {
      "shape_name": "矩形 15",
      "para_index": 0,
      "text": "持续领先，成为最可信赖的服务伙伴",
      "char_count": 17,
      "font_size": 24.0,
      "font_name": "微软雅黑",
      "bold": true,
      "italic": false,
      "color": "C8102E",
      "left": 0.85,
      "top": 0.42,
      "width": 8.50,
      "height": 0.65
    }
  ]
}
```

### 5.2 字段说明

| 层级 | 字段 | 类型 | 说明 |
|------|------|------|------|
| 页级 | `slide` | int | 幻灯片序号（1-based，PPT实际索引） |
| 页级 | `struct_type` | string | 结构类型（并列/通用/总分总/KPI/总分/纯图） |
| 页级 | `confidence` | float | 分类置信度（0.5兜底 ~ 0.9高置信） |
| 页级 | `description` | string | 分类依据简述（人可读） |
| 页级 | `slot_count` | int | 总槽位数（文本段落数） |
| 页级 | `total_chars` | int | 原始模板总字数 |
| 槽级 | `shape_name` | string | PPT形状名（用于定位，⚠️可能重复） |
| 槽级 | `para_index` | int | 形状内段落序号（多段落shape的关键字段） |
| 槽级 | `text` | string | 原始文本内容 |
| 槽级 | `char_count` | int | 字数（蒸馏目标长度） |
| 槽级 | `font_size` | float | 字号（pt），角色分类主要依据 |
| 槽级 | `font_name` | string | 字体名称 |
| 槽级 | `bold` | bool | 粗体标记（角色推断关键信号） |
| 槽级 | `italic` | bool | 斜体标记 |
| 槽级 | `color` | string | 颜色值（RGB hex 或 "SCHEME"） |
| 槽级 | `left/top` | float | 左上角坐标（英寸） |
| 槽级 | `width/height` | float | 文本框宽高（英寸） |

### 5.3 已知陷阱

| 陷阱 | 表现 | 对策 |
|------|------|------|
| **shape_name 不唯一** | 多个"圆角矩形 254"、多个"智 慧 出 行" | 用原始文本内容+坐标双重匹配 |
| **多段落shape** | 一个shape含多个paragraph | 匹配全文本后按 `\n` 分割逐段替换 |
| **PPT索引偏移** | Catalog slide序号的索引值≠PPT实际 `prs.slides[idx]` 位置 | 始终用 `slide` 字段-1作为实际索引 |
| **字号为0** | 部分shape无显式字号（继承主题） | `classify_slot_role()` 处理 size=0 情况 |

---

## 六、结构类型分类器

### 6.1 分类决策树

```
输入: 一页的所有 text_slots
         │
    ┌────┴───── slots为空？
    │ YES         │ NO
    ▼             ▼
  纯图页      ┌─────────────────────┐
  (0.5)      │ 1. 封面/标题页?      │
             │   ≥24pt粗体 仅1个    │
             │   ≤3槽 <100字        │
             └───────┬─────────────┘
                     │ NO
             ┌───────┴─────────────┐
             │ 2. 章节分隔?          │
             │   ≥24pt粗体 仅1个    │
             │   无正文 <60字        │
             └───────┬─────────────┘
                     │ NO
             ┌───────┴─────────────┐
             │ 3. KPI数据展示?       │
             │   2-6个大数字(≥24pt  │
             │   非粗体 top≥1.5")    │
             │   ≤12槽 <300字        │
             └───────┬─────────────┘
                     │ NO
             ┌───────┴─────────────┐
             │ 4. 总分结构?          │
             │   1个大标题           │
             │   3-8个中等分项(16-24)│
             └───────┬─────────────┘
                     │ NO
             ┌───────┴─────────────┐
             │ 5. 总分总?            │
             │   标题 + ≥4段正文     │
             │   总字>200 + 尾部小结  │
             └───────┬─────────────┘
                     │ NO
             ┌───────┴─────────────┐
             │ 6. 并列结构?          │
             │   ≥3个同字号分项      │
             │   字号差≤4pt         │
             │   字数方差<30         │
             └───────┬─────────────┘
                     │ NO
                     ▼
              通用内容 (兜底, 0.5)
```

### 6.2 默认模板分布

| 结构类型 | 页数 | 占比 | 槽位范围 | 典型特征 |
|----------|------|------|----------|----------|
| 并列结构 | 79 | 53.0% | 4-61槽 | 多主题并行展示，含KPI标签模式 |
| 通用内容 | 67 | 45.0% | 1-118槽 | 灵活布局，封面/架构/图+文混排 |
| 总分总 | 5 | 3.4% | 18-44槽 | 汇报骨架：概述→分述→归纳 |
| KPI数据 | 3 | 2.0% | 9-11槽 | 关键指标大字报 |
| 总分 | 2 | 1.3% | 11-49槽 | 产品/方案单层展开 |
| 纯图 | 1 | 0.7% | 0槽 | 零文字过渡页 |

---

## 七、智能蒸馏算法

### 7.1 三阶策略

```
distill(target_len, text):
  │
  ├── text ≤ target_len? → 直接返回
  │
  ├── target_len ≤ 8 (短槽):
  │   在 target_len 范围内找最后一个自然断点:
  │     "，" "。" "；" "、" " " "】" "）" "：" ":"
  │   无断点 → 返回 target_len（整个词）
  │
  ├── target_len 9~20 (中槽):
  │   在 target_len 范围内找分句点:
  │     "。" "；" "，" "、"
  │   必须 ≥ target_len×0.5 才截断
  │   无标点 → 返回 target_len
  │
  └── target_len > 20 (长槽):
      允许115%超长: len(text) ≤ target_len×1.15? → 返回全文
      否则在 target_len×1.1 范围内找句号:
        "。" "；"
      必须 ≥ target_len×0.5 才截断
      无句号 → 返回 target_len
```

### 7.2 蒸馏示例

| 输入 | target | 策略 | 输出 | 说明 |
|------|--------|------|------|------|
| `区域AI工具建设建议与治理规范` | 8 | 短槽·标点断 | `区域AI工具建设` | 在"建议"后"与"前断——虽然无逗号但取8字整词 |
| `建立统一的数据治理规范，涵盖数据采集` | 14 | 中槽·逗号断 | `建立统一的数据治理规范，` | 逗号后断，12字 |
| `区域AI工具建设需要从组织架构、流程规范、技术平台三个维度协同推进，确保安全可控` | 25 | 长槽·115% | `区域AI工具建设需要从组织架构、流程规范、技术平台三个维度协同推进。` | 句号28字≤28.75(25×1.15) |
| `AI` | 4 | 短槽·无事可断 | `AI` | 已短于目标 |

### 7.3 蒸馏原则

1. **永不在词语中间截断** — 这是最高原则，高于所有其他规则
2. **优先在标点处截断** — "。" > "；" > "，" > "、"
3. **宁短勿碎** — 宁可少几个字，不出半个词
4. **绝不用省略号** — 不用"…"暗示截断
5. **长正文允许适度溢出** — 115%弹性空间

---

## 八、角色匹配系统

### 8.1 槽位角色分类 (`classify_slot_role`)

```
输入: slot字典 + 整页 all_slots（获取全局字号上下文）

     字号≥24pt + 粗体 + top<1.5"  →  主标题
     字号≥16pt + 粗体              →  分项标题
     字号≥24pt + 非粗 + top≥1.5"  →  数据值
     12≤字号<20 + 字数>12          →  描述正文
     12≤字号<20 + 粗体 + 字数≤12   →  分类标签
     10≤字号<20 + 字数≤12          →  短标签
     字号<10pt                      →  注释
     其他                          →  通用文本（兜底）
```

### 8.2 内容分配逻辑

```
对于每页选定模板:
  1. 获取所有槽位 → classify_slot_role() 分组
  2. 用户内容块按角色归入对应的内容池
  3. 分配:
     for 该角色的每个槽:
       slot ← 内容池[i % len(内容池)]
       if 内容池已空 → slot ← 兜底池[i % len(兜底池)]
  4. 蒸馏: slot ← distill(slot.char_count, slot.assigned_text)
  5. 替换: replace_text_keep_font(slot.paragraph, slot.distilled_text)
```

### 8.3 短-短匹配策略 (`fill_all_slots`)

短内容优先分配给短槽，避免"大槽装小词"的稀疏感：

```
1. 按槽宽升序排列槽位
2. 按内容字数升序排列内容
3. 一一对应分配
4. 剩余长内容 → 分配给剩余大槽
```

---

## 九、工作流（流水线）

### 9.1 完整流水线

```bash
# Step 1: 分析用户意图
python scripts/prompt_analyzer.py "用户内容描述"

# Step 2: 选取最佳模板页
python scripts/slide_selector.py --content result_from_step1.json

# Step 3: 逻辑结构匹配 + 逐框填字
python scripts/gen_logic_match.py --slides <selected> --content <content_json>

# Step 4: (可选) 批量内容填充
python scripts/content_filler.py

# Step 5: 合并多页为最终输出
python scripts/merge_ppt.py --slides <page1> <page2> ... --output 最终输出.pptx
```

### 9.2 快速单页生成

最简单的生成路径——直接调用核心引擎：

```bash
cd skills/hw-slide-master-skill/scripts
python gen_logic_match.py
```

脚本内编辑 `SLIDE_NUM`（模板页号）和内容池，运行即输出。

### 9.3 自定义模板生成

```
初始化选择"选项2"后:
  1. python scripts/analyze_template.py -t <用户PPT> -o references/custom_template_catalog.json
  2. 根据统计摘要选取目标页
  3. 使用 gen_logic_match.py（指定catalog路径和模板路径）
```

---

## 十、文件结构

```
skills/hw-slide-master-skill/
├── SKILL.md                        ← 本文档（技能说明+架构+使用指南）
├── TEMPLATE_PATH.txt               ← 默认模板路径记录
│
├── template/
│   └── hw_template.pptx            ← 默认华为模板（149页·35MB）
│
├── references/
│   ├── template_catalog_v2.json    ← 默认模板DNA目录（157条·~1MB）
│   ├── custom_template_catalog.json← 用户自定义模板DNA（由analyze生成）
│   └── structure_templates.json    ← 7种内容结构模式定义
│
├── scripts/
│   ├── analyze_template.py         ← 【通用】模板分析引擎（核心入口）
│   ├── template_utils.py           ← 【共享】工具库（distill/replace/classify/load）
│   ├── gen_logic_match.py          ← 【核心】逻辑匹配+替换引擎
│   ├── slide_selector.py           ← 【匹配】模板页选取器
│   ├── prompt_analyzer.py          ← 【解析】用户意图分析
│   ├── content_filler.py           ← 【填充】批量内容填充
│   ├── merge_ppt.py                ← 【合并】多页PPT合并
│   ├── analyze_structure_v2.py     ← 【历史】早期分析脚本（由analyze_template替代）
│   ├── gen_economy.py              ← 【示例】中国经济主题生成
│   └── gen_china_house.py          ← 【示例】中国房价主题生成
│
└── gen_*.py（根目录）              ← 独立生成脚本（美国文化等示例）
```

---

## 十一、关键函数参考

### 11.1 `template_utils.py` — 共享工具库

| 函数 | 签名 | 用途 |
|------|------|------|
| `replace_text_keep_font` | `(para, new_text: str)` | 替换段落文本，保留全部 `<a:rPr>` 格式 |
| `distill` | `(target_len: int, text: str) → str` | 三阶智能蒸馏到目标字数 |
| `classify_slot_role` | `(slot: dict, all_slots: list) → str` | 推断文本框逻辑角色 |
| `pick_from_pools` | `(target_len: int, idx: int) → str` | 从对应字数池取词 |
| `load_template_catalog` | `(path: str) → list` | 加载模板目录JSON |
| `get_slide_slots` | `(catalog: list, slide_num: int) → list` | 获取指定页所有槽位 |
| `match_shapes_to_plan` | `(slide, plan: list) → list` | 坐标+文本匹配shape到替换计划 |
| `extract_single_slide` | `(tmpl: str, out: str, n: int) → Presentation` | 从模板提取单页到新PPT（shape复制法） |

### 11.2 `analyze_template.py` — 通用分析引擎

```bash
# 分析默认模板
python scripts/analyze_template.py

# 分析自定义模板
python scripts/analyze_template.py -t /path/to/custom.pptx -o references/custom_template_catalog.json

# 查看帮助
python scripts/analyze_template.py --help
```

参数：
- `--template` / `-t`：模板PPTX路径（默认：`template/hw_template.pptx`）
- `--output` / `-o`：输出JSON路径（默认：`references/template_catalog_v2.json`）

### 11.3 `gen_logic_match.py` — 核心生成引擎

内置完整流水线：
1. 加载模板目录 → 定位目标页
2. 槽位角色分类 (`classify_slot_role`)
3. 内容池定义（短标签/分项标题/描述文/分类标签/注释，每个角色独立池）
4. 短-短匹配分配 (`fill_all_slots`)
5. 逐槽蒸馏 (`distill`)
6. 逐槽替换 (`replace_text_keep_font`)
7. 输出单页PPTX

---

## 十二、扩展与维护指南

### 12.1 新增结构类型

1. 在 `classify_structure_type()` 中添加新的判断分支（置于优先级链适当位置）
2. 在 `references/structure_templates.json` 中补充新模式定义
3. 更新本SKILL.md中的分类器决策树

### 12.2 新增蒸馏策略

在 `distill()` 中添加新的 `target_len` 区间分支：
```python
if target_len <= some_new_threshold:
    # 新的截断逻辑
```

保持三个核心约束不变：不在词中间截断、不用省略号、标点优先。

### 12.3 适配新模板

```
1. 获取模板 PPTX
2. python scripts/analyze_template.py -t <new_template> -o references/custom_template_catalog.json
3. 阅读统计摘要，了解结构分布
4. 使用 gen_logic_match.py 测试单页替换
```

### 12.4 已知技术壁垒

| 壁垒 | 现状 | 影响 |
|------|------|------|
| shape_name 重复 | 已绕过（文本+坐标匹配） | 不能仅凭shape_name定位 |
| extract_single_slide 原方案失效 | 已重写（shape复制法） | 旧代码（sldIdLst.remove）不可用 |
| 中英混合文件名 | Qwen会插入空格 | 模板已重命名为纯英文 `hw_template.pptx` |
| PPT索引偏移 | Catalog slide ≠ PPT actual index | 多余幻灯片被删后索引不连续 |

### 12.5 维护检查清单

- [ ] `template_catalog_v2.json` 与 `hw_template.pptx` 页数一致（149页→149条有效条目，当前157条含8个被删页残余）
- [ ] `analyze_template.py` 默认参数指向正确模板
- [ ] 所有 `.py` 脚本无硬编码绝对路径
- [ ] 所有路径引用使用 `_SKILL_DIR` 相对推导
- [ ] 每次修改核心函数后运行单页生成测试

---

## 十三、约束与限制

| 约束 | 说明 |
|------|------|
| **不写Excel** | 模板替换用 python-pptx，不同时调用 xlsx skill |
| **不创建新形状** | 只替换文字，不新增/删除/移动任何布局元素 |
| **不处理图片/SmartArt** | `has_text_frame=False` 的形状原样保留 |
| **模板决定一切** | 输出排版质量完全取决于模板设计质量 |
| **先分析再使用** | 自定义模板必须先通过 `analyze_template.py` 生成Catalog |
| **单页提取有损耗** | `extract_single_slide` 使用空白布局，不保留原始母版主题 |

---

*— 模板是骨架，内容是血肉，此技能只负责输血，不负责长骨⚡*
