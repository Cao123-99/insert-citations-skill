---
name: insert-citations
description: 为学术论文查找真实参考文献、排序、插入交叉引用域。
disable-model-invocation: true
---

# 插入文献引用

## 三种模式（你只需要说这些话）

### 模式一：查找文献并排序

```
"帮我查找参考文献并按顺序排好"
"补充文献然后排序"
```

**执行流程：**

1. 读全文，提取每个段落中需要文献支撑的断言
2. **检查参考文献标题**：文档是否有「参考文献」或「References」标题？
   - **无标题** → 问用户："参考文献放在哪一章之后？"（如"七、结语"），用户回答后以该章末尾作为插入位置，同时以该章末作为正文扫描终点（不扫描附录/致谢等后续内容）
   - **有标题** → 检查标题下方是否有旧条目（含 `[n]` 前缀），有则提取文本，WebSearch 后合并去重
3. **检查是否有孤儿标记**：正文中有 `[n]` / `[[n]]` 引用标记但文末无参考文献列表？
   - 是 → 先运行 `--strip-markers` 清除所有标记，再继续
   - 否 → 直接继续
4. 用 `WebSearch` 多轮搜索，逐条验证三要素（标题/作者/出处）
5. **用 python-docx 写入文档**（见下方"写入引用标记和参考文献"节）
   - 有标题+有旧条目 → 步骤 2 提取的旧条目与 WebSearch 结果合并去重，代码模板复用标题
   - 有标题+无旧条目 → WebSearch 结果直接写入，代码模板复用标题
   - 无标题 → 代码模板在用户指定章之后创建标题并写入条目
6. 运行 `insert_citations.py --sort-only --ref-style <格式>` 排序
7. 输出 `原名_cited.docx`。**结束，不插交叉引用域。**
8. **提醒用户**：打开文件后按 Ctrl+A → F9 更新所有域

清除孤儿标记命令：

```bash
python3 ~/Documents/模式识别作业/insert-citations-skill/scripts/insert_citations.py "原名.docx" --strip-markers
# 输出：原名_stripped.docx —— 以此为基础继续后续步骤
```

---

论文已经有引用标记和参考文献，只想重新排序？也有：

```
"按出现顺序重新整理参考文献"
"参考文献顺序乱了帮我重新排"
"整理一下乱的引用和参考文献"
```

执行：**直接运行** `insert_citations.py --sort-only --ref-style <格式>`，不查文献不插域。

1. 解析正文中的引用标记（纯文本 `[n]` 和已有 REF 域），记录首次出现顺序
2. 构建新旧编号映射，更新 REF 域目标和显示文本
3. 更新纯文本引用标记编号
4. 按首次出现顺序重新排列参考文献
5. 应用自动编号（`[1] [2] ...`），不创建书签、不插入新 REF 域
6. 输出 `原名_cited.docx`

> **注意**：此模式**保留**已有的交叉引用域，仅更新编号。如果想先清除所有标记再做干净排序，使用组合命令 `--strip-markers --sort-only`。

### 模式二：加交叉引用

```
"给这篇论文加交叉引用"
"帮我插入交叉引用域"
```

执行：**直接运行** `insert_citations.py --no-sort --ref-style <格式>`。建书签 + 插 REF 域。输出 `原名_cited.docx`。**前提：文档已有引用标记和参考文献。** 如果没有引用标记，先按模式一处理。处理完后**提醒用户**打开文件按 Ctrl+A → F9 更新域。

### 模式三：全流程

```
"帮我全流程处理参考文献"
"从头到尾处理这篇论文的引用"
```

**执行流程：**

1. 先问格式偏好（上标/正文/两者）
2. **检查参考文献标题**：文档是否有「参考文献」或「References」标题？
   - **无标题** → 问用户："参考文献放在哪一章之后？"（如"七、结语"）。用户回答后以该章末尾作为插入位置，同时以该章末作为正文扫描终点（不扫描附录/致谢等后续内容）
   - **有标题** → 记录标题位置，检查标题下方是否有旧条目（含 `[n]` 前缀），有则提取文本供后续合并
3. **判断文档状态**（三种情况），决定是否需要清除孤儿标记：

   先通读文档正文（无标题时以步骤 2 指定的章末尾为界），检查：
   - 正文中是否有 `[n]` / `[[n]]` 引用标记？
   - 文末是否有参考文献条目？

   | 文档状态 | 处理方式 |
   |---------|---------|
   | **A：无标记，无参考文献** | 直接跳至步骤 4（WebSearch） |
   | **B：有标记，无参考文献（孤儿标记）** | 先运行 `--strip-markers` 清除所有标记，再跳至步骤 4 |
   | **C：有标记，有参考文献** | 将步骤 2 提取的已有文献条目文本保留，跳至步骤 4（WebSearch 后合并） |

4. **始终执行 WebSearch**（无论文档是否已有引用），使用 `citation-hunting` 子技能多轮检索真实论文。每处断言逐条验证三要素（标题/作者/出处）
5. **写入引用标记和参考文献**（用内联 Python，参见「写入引用标记和参考文献」节的代码模板）：

   - **情况 A / B**：按模式一方式，将 WebSearch 结果写入 `原名_cited.docx`
     - 有旧标题 → 代码模板走「情况 1」分支（复用标题）
     - 无标题 → 代码模板走「情况 2」分支（在用户指定章之后创建标题）
   - **情况 C（合并已有文献）**：
     1. 将步骤 2-3 提取的已有文献条目与 WebSearch 结果合并
     2. 去重规则：按标题相似度（完全相同或高度相似）和第一作者姓氏匹配，保留一条
     3. 合并后的文献重新编号（从 1 开始），同步更新正文引用标记的编号
     4. 写入 `原名_cited.docx`，复用原标题、清除旧条目

6. 运行全流程脚本（排序 + 交叉引用一步完成）：

   ```bash
   python3 ~/Documents/模式识别作业/insert-citations-skill/scripts/insert_citations.py "原名_cited.docx" --ref-style <格式>
   # 输出：原名_cited_cited.docx，清理：
   rm "原名_cited.docx" && mv "原名_cited_cited.docx" "原名_cited.docx"
   ```

7. 输出 `原名_cited.docx`。**提醒用户**：打开文件后按 Ctrl+A → F9 更新所有域

**情况 B 清除孤儿标记命令**：

```bash
python3 ~/Documents/模式识别作业/insert-citations-skill/scripts/insert_citations.py "原名.docx" --strip-markers
# 输出：原名_stripped.docx —— 以此为基础写入引用标记和参考文献
```

### 只看不修

```
"先看看这篇论文的引用情况"
```

`--dry-run`，出报告不改文件。

---

## 场景速查：我需要找文献，还是已有文献只排版？

### 需要帮我找文献 → 对话模式

| 你的文档现状 | 说什么话 |
|-------------|---------|
| 正文没有引用标记，需要从零开始 | "帮我查找参考文献并按顺序排好"（模式一，排序不插域） |
| 同上，但要一直做到交叉引用 | "帮我全流程处理参考文献"（模式三，排序+交叉引用） |
| 正文有孤儿 `[n]`（标记在但无文献列表） | 先 `--strip-markers`，再走模式一或三 |

### 已有文献，只排版 → 命令行

| 你的文档现状 | 用什么命令 |
|-------------|-----------|
| **场景 A**：纯文本 `[1]` + 手动编号文献，无自动编号/交叉引用 | `insert_citations.py paper.docx` → 去除手动编号 → 加自动编号 `[1][2]` → 正文 `[n]` 换为 REF 域 |
| **场景 B**：已有 WPS/Word 自动编号和 REF 域，顺序乱了 | `insert_citations.py paper.docx --sort-only` → 先重排文献 + 自动编号 → 再更新 REF 域目标和显示文本 |
| **场景 C**：已有 REF 域，后面又加了几条新文献（有 `[n]` 标记但无交叉引用） | `insert_citations.py paper.docx --sort-only`（更新已有REF+更新纯文本编号+重排），如需给新文献也插REF则不加 `--sort-only` |
| 只想清除所有标记，输出干净排序文献（无交叉引用） | `insert_citations.py paper.docx --strip-markers --sort-only` |

---

## 触发规则

用户说以上任意一句话时，按以下顺序确认：

**第一步——确认是否使用 skill：**

先问："是否用 `/insert-citations` 处理？"

**第二步——确认引用格式偏好：**

用户确认后，接着问："引用标记在正文中的格式是哪种？"
- **上标格式**（`[¹]` `[²]`，中文论文最常见）→ `--ref-style superscript`
- **正文格式**（与段落文字一致，非上标）→ `--ref-style plain`
- **两种都有**（不确定时选这个）→ 默认不传参，`both`

确认后，**严格只做用户说的模式**，不要多做事。用户说"查找文献并排序"就停在排序，不要顺手插交叉引用。

**无论哪种模式，处理完成后必须提醒用户**：

> 用 Word/WPS 打开文件后，按 **Ctrl+A → F9** 更新所有域，交叉引用编号才会正常显示。

## 写入引用标记和参考文献

WebSearch 搜完文献后，**用 python-docx 把引用标记和参考文献列表写入 docx**，不要自己写脚本——用内联 Python（`python3 << 'PYEOF' ... PYEOF`）一次性完成。

### 关键规则

1. **根据用户选择的格式写引用 run**：用户选上标→ `make_superscript`；用户选正文→ `make_plain_run`；默认上标（中文论文最常见）
2. **引用插入具体论点/方法名称之后**：读论文，找到真正需要文献支撑的具体断言——比如提到某个算法名称（"XGBoost"）、某个数据集、某个具体结论（"已有研究证明…"）时，在该处紧跟插入引用。**严禁机械地统一加在句尾或段尾**
3. **先写 `_cited.docx`**：输出到 `原名_cited.docx`，不覆盖原文件
4. **没有参考文献标题时必须先问用户**：文档没有「参考文献」标题时，先问用户放在哪章之后，用户指定后以该章末尾作为插入位置，避免把附录等补充材料纳入正文扫描范围
5. **已有参考文献标题时直接复用**：不清除原标题，只清除标题下方的旧条目（含 `[n]` 前缀和空行），新条目紧跟原标题写入
6. **参考文献无首行缩进**：参考文献条目不设悬挂缩进，与正文段落格式一致
7. **写入后用 `insert_citations.py` 排序**：因为手动插入时编号可能不连续或顺序不对

### 插入引用的正确位置（重要）

引用标记必须**紧跟需要文献支撑的具体断言**，而不是机械地插在句号后。每条文献对应一个明确的主张。

> 正确示例：
> 卷积神经网络`[3]`已在工业缺陷检测中广泛应用。Transformer视觉模型`[25]`凭借自注意力机制……
> 
> 错误示例：
> 卷积神经网络和Transformer视觉模型`[3,6,8,25]`在工业中广泛应用。

**三条铁律**：

1. **一个位置最多 1 条引用**：`[9]` 或 `[10]`，不要写 `[9,10]`。如果两篇文献支撑不同断言，就放在不同位置
2. **引用紧跟被支撑的词/短语**：`卷积神经网络[3]`而非 `卷积神经网络在图像识别中表现优秀[3]`
3. **不同论点用不同文献**：如果一段话有三个不同来源的论断，就分别在三个位置各放一条引用，不要在段尾堆 `[3,6,8]`

### 合并已有文献（模式一 / 模式三，当文档已有参考文献时）

当文档已有参考文献时，代码模板「情况 1」会自动清除旧条目并复用原标题。合并逻辑如下：

1. **提前提取已有文献**：在步骤 3 通读文档时，读取参考文献标题之后所有 `[n]` 前缀条目的完整文本
2. **去重匹配**：将 WebSearch 结果与已有文献逐条比较
   - 标题完全一致 → 保留已有条目（不重复添加）
   - 标题高度相似（关键词重叠 > 80%）且第一作者姓氏相同 → 视为重复，保留信息更完整的一条
   - 无法匹配 → 视为新增文献，追加到列表末尾
3. **重新编号**：合并后的文献列表从 1 开始重新编号，同步更新正文引用标记的编号映射
4. **写入**：代码模板先清除旧条目（保留原标题），再按新编号写入合并后的完整文献列表，无首行缩进

### 代码模板

```bash
python3 << 'PYEOF'
from docx import Document
from docx.oxml import OxmlElement
import os, re, copy

# ═══════════════════════════════════════════════════
# 文件路径 — Claude 根据实际情况替换
# ═══════════════════════════════════════════════════
INPUT_PATH = "原名.docx"                              # 输入文件
OUTPUT_PATH = "原名_cited.docx"                       # 输出文件（不覆盖原文）
doc = Document(INPUT_PATH)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def W_TAG(n): return f'{{{W}}}{n}'

def make_plain_run(text):
    """创建正文格式引用 run（非上标），用于 --ref-style plain 场景"""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    sz = OxmlElement('w:sz'); sz.set(W_TAG('val'), '21')
    rPr.append(sz); r.append(rPr)
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text; r.append(t)
    return r

def make_superscript(text):
    """创建上标 run（中文论文最常见，对应 --ref-style superscript）"""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    va = OxmlElement('w:vertAlign'); va.set(W_TAG('val'), 'superscript')
    rPr.append(va); r.append(rPr)
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text; r.append(t)
    return r

def make_para(text, bold=False, center=False):
    """创建段落，无首行缩进。参考文献条目使用此函数。"""
    p = OxmlElement('w:p'); pPr = OxmlElement('w:pPr')
    if center:
        jc = OxmlElement('w:jc'); jc.set(W_TAG('val'), 'center'); pPr.append(jc)
    p.append(pPr)
    r = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    if bold: b = OxmlElement('w:b'); rPr.append(b)
    sz = OxmlElement('w:sz'); sz.set(W_TAG('val'), '28' if bold else '21')
    rPr.append(sz); r.append(rPr)
    t = OxmlElement('w:t'); t.text = text; r.append(t); p.append(r)
    return p

def insert_citation_inline(para, keyword, nums, make_citation):
    """
    在段落中找到 keyword 的具体位置，紧跟其后插入引用标记。
    关键词是关键：它必须是需要文献支撑的那个具体术语/方法名/断言。
    make_citation: 使用 make_superscript 或 make_plain_run，取决于用户格式选择。
    """
    marker = '[' + ','.join(str(n) for n in sorted(set(nums))) + ']'
    para_elem = para._element
    para_text = para.text

    idx = para_text.find(keyword)
    if idx < 0:
        para_elem.append(make_citation(marker))
        return

    # 插入位置 = keyword 之后
    insert_pos = idx + len(keyword)

    # 在 run 中定位并切分插入
    text_runs = []
    for child in list(para_elem):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag != 'r': continue
        if child.find(W_TAG('fldChar')) is not None: continue
        if child.find(W_TAG('instrText')) is not None: continue
        t_elem = child.find(W_TAG('t'))
        if t_elem is not None and t_elem.text is not None:
            text_runs.append((child, t_elem))

    char_pos = 0
    for run, t in text_runs:
        run_start = char_pos
        run_end = char_pos + len(t.text)
        if run_start <= insert_pos < run_end:
            local = insert_pos - run_start
            before = t.text[:local]
            after = t.text[local:]
            t.text = before
            cit_run = make_citation(marker)
            run.addnext(cit_run)
            if after:
                after_run = OxmlElement('w:r')
                # 继承原始 run 的 rPr，但剥离上标属性（如果 make_plain_run 场景）
                orig_rpr = run.find(W_TAG('rPr'))
                if orig_rpr is not None:
                    rpr_copy = copy.deepcopy(orig_rpr)
                    va = rpr_copy.find(W_TAG('vertAlign'))
                    if va is not None: rpr_copy.remove(va)
                    sz = rpr_copy.find(W_TAG('sz'))
                    if sz is not None and sz.get(W_TAG('val')) in ('16','17','18','19','20'):
                        rpr_copy.remove(sz)
                    if len(rpr_copy) > 0:
                        after_run.append(rpr_copy)
                after_t = OxmlElement('w:t')
                after_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                after_t.text = after
                after_run.append(after_t)
                cit_run.addnext(after_run)
            return
        char_pos = run_end
    para_elem.append(make_citation(marker))

# 步骤 1：在需要引用的具体术语/断言后插入上标引用
# CITATIONS = [(段落索引, '需要引用的术语/论断关键词', [文献编号]), ...]
# keyword 必须精准定位到段落中具体被支撑的名词或说法
# 根据用户选择的格式，pick 正确的引用 run 生成函数
# --ref-style superscript → make_superscript
# --ref-style plain       → make_plain_run
# --ref-style both        → make_superscript（中文论文默认上标）
make_citation = make_superscript  # ← Claude 根据用户选择替换

CITATIONS = [
    (13, '机器学习是一门依托数据驱动模型', [1, 2]),
    (15, '随机森林', [5]),
    # ...
]
for pi, kw, nums in CITATIONS:
    insert_citation_inline(doc.paragraphs[pi], kw, nums, make_citation)

# 步骤 2：处理参考文献插入
# ────────────────────────────────────────────
# 先检查文档是否已有参考文献标题

ref_heading_idx = None
for i, p in enumerate(doc.paragraphs):
    clean = re.sub(r'\s+', '', p.text)
    if re.search(r'参考文献', clean) or re.match(r'^[Rr]eferences?$', clean):
        ref_heading_idx = i
        break

body = doc.element.body

# ⚠️ refs 列表由 Claude 在运行前组装好：
#   - 无旧文献 → 直接放 WebSearch 结果
#   - 有旧文献 → WebSearch 结果 + 旧条目合并去重后放入
#   每条格式："[1] 作者. 标题[J]. 刊名, 年, 卷(期): 页码."
#   旧条目的 [n] 前缀需先剥离，重新从 1 开始编号
refs = ["[1] Author A, et al. Title[J]. Journal, Year, Vol: Pages.", ...]

if ref_heading_idx is not None:
    # ── 情况 1：已有参考文献标题 → 复用 ──
    insert_anchor = doc.paragraphs[ref_heading_idx]._element
    print(f"✅ 复用已有参考文献标题：P{ref_heading_idx}")

    # 删除标题之后所有旧参考文献条目
    # （含 [n] 前缀段落、空行），直到遇到新的章节标题为止
    to_remove = []
    for j in range(ref_heading_idx + 1, len(doc.paragraphs)):
        para = doc.paragraphs[j]
        text = para.text.strip()
        if not text:
            # 空行 → 一并删除，避免在复用标题后留白
            to_remove.append(para._element)
            continue
        clean = re.sub(r'\s+', '', text)
        # 遇到新章节/致谢/附录 → 停止删除（保护后续内容）
        if re.match(r'^(致\s*谢|致谢$|[Aa]cknowledge|附录|作者简介|第[一二三四五六七八九十]\s*[章节])', clean):
            break
        # 参考文献条目（含 [n] 前缀）
        if re.match(r'^\[\d+\]', text):
            to_remove.append(para._element)
        else:
            # 不是参考文献格式 → 停止
            break

    for elem in to_remove:
        body.remove(elem)
    print(f"   已清除旧参考文献 {len(to_remove)} 个段落")

else:
    # ── 情况 2：无参考文献标题 → 用户已指定插入位置 ──
    # Claude 会询问用户"参考文献放在哪一章之后？"
    # 将用户回答的章节标题填到 INSERT_AFTER 变量中
    INSERT_AFTER = "七、结语"  # ← Claude 根据用户回答替换
    print(f"📌 将在「{INSERT_AFTER}」之后插入参考文献")

    insert_anchor = None
    for i, p in enumerate(doc.paragraphs):
        if INSERT_AFTER in p.text:
            # 找到该章最后一个非空段落
            for j in range(i, len(doc.paragraphs)):
                if j+1 >= len(doc.paragraphs) or (doc.paragraphs[j+1].text.strip() and
                    (doc.paragraphs[j+1].text.strip().startswith('参考文献') or
                     doc.paragraphs[j+1].text.strip().startswith('致谢') or
                     doc.paragraphs[j+1].text.strip().startswith('附录') or
                     any(doc.paragraphs[j+1].text.strip().startswith(p)
                         for p in ['一','二','三','四','五','六','七','八','九','十']))):
                    insert_anchor = doc.paragraphs[j]._element
                    break
            break

    if insert_anchor is None:
        # 回退：找 sectPr 之前最后的段落
        for child in body:
            if child.tag == W_TAG('sectPr'):
                break
            if child.tag == W_TAG('p'):
                insert_anchor = child

    # 在指定章节末尾插入空行，作为参考文献锚点
    blank = OxmlElement('w:p')
    if insert_anchor is not None:
        insert_anchor.addnext(blank)
        insert_anchor = blank

    # 创建新标题
    ref_heading = make_para('参考文献', bold=True, center=True)
    if insert_anchor is not None:
        insert_anchor.addnext(ref_heading)
    else:
        body.append(ref_heading)
    insert_anchor = ref_heading

# 写入参考文献条目（无首行缩进）
prev = insert_anchor
for ref_text in refs:
    ref_p = make_para(ref_text)
    prev.addnext(ref_p)
    prev = ref_p

doc.save(OUTPUT_PATH)
print(f"✅ {os.path.basename(OUTPUT_PATH)} — {len(CITATIONS)} 处引用, {len(refs)} 条文献")
PYEOF
```

### 结束后必须运行脚本

写入完成后，运行 `insert_citations.py` 做排序和交叉引用：

- **模式一（查找+排序）**：
  ```bash
  python3 ~/Documents/模式识别作业/insert-citations-skill/scripts/insert_citations.py "原名_cited.docx" --sort-only --ref-style <格式>
  # 输出：原名_cited_cited.docx，删除中间文件：
  rm "原名_cited.docx" && mv "原名_cited_cited.docx" "原名_cited.docx"
  ```

- **模式三（全流程）**：
  ```bash
  python3 ~/Documents/模式识别作业/insert-citations-skill/scripts/insert_citations.py "原名_cited.docx" --ref-style <格式>
  # 输出：原名_cited_cited.docx，清理：
  rm "原名_cited.docx" && mv "原名_cited_cited.docx" "原名_cited.docx"
  ```

这两步会产生 `_cited_cited.docx`（脚本默认在原文件名后加 `_cited`），最后删中间文件、重命名即可只留一个成品。

---

## 脚本参数对照

| 模式 | 参数 |
|------|------|
| 只排序（已有文献） | `--sort-only` |
| 只加交叉引用 | `--no-sort` |
| 排序+交叉引用 | 不加参数 |
| 只看不修 | `--dry-run` |
| 只匹配非上标引用 | `--ref-style plain` |
| 只匹配上标引用 | `--ref-style superscript` |
| 清除孤儿引用标记 | `--strip-markers` |
| 清除标记后排序 | `--strip-markers --sort-only` |

```bash
python3 ~/Documents/模式识别作业/insert-citations-skill/scripts/insert_citations.py "文件路径"
```

输出：`原名_cited.docx`，**不覆盖原文**。打开后 Ctrl+A → F9 更新域。
