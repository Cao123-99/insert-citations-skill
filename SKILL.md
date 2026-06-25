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
2. 用 `WebSearch` 多轮搜索，逐条验证三要素（标题/作者/出处）
3. **用 python-docx 写入文档**（见下方"写入引用标记和参考文献"节）
4. 运行 `insert_citations.py --sort-only --ref-style <格式>` 排序
5. 输出 `原名_cited.docx`。**结束，不插交叉引用域。**
6. **提醒用户**：打开文件后按 Ctrl+A → F9 更新所有域

---

论文已经有引用标记和参考文献，只想重新排序？也有：

```
"按出现顺序重新整理参考文献"
"参考文献顺序乱了帮我重新排"
"整理一下乱的引用和参考文献"
```

执行：**直接运行** `insert_citations.py --sort-only --ref-style <格式>`，不查文献不插域。

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
2. 再判断：文档是否已有引用标记？
   - **有引用** → 直接运行 `insert_citations.py --ref-style <格式>`（排序 + 交叉引用一步完成）
   - **无引用** → 先执行模式一（WebSearch → 写入标记 → `--sort-only` 排序），再执行模式二（`--no-sort` 插域）。或者写完后一次性运行 `insert_citations.py --ref-style <格式>` 全流程

### 只看不修

```
"先看看这篇论文的引用情况"
```

`--dry-run`，出报告不改文件。

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

1. **引用标记是上标 run**：中文论文的 `[1]` `[1,2]` `[1,3-5]` 必须用上标 run 写入
2. **引用插入具体论点/方法名称之后**：读论文，找到真正需要文献支撑的具体断言——比如提到某个算法名称（"XGBoost"）、某个数据集、某个具体结论（"已有研究证明…"）时，在该处紧跟插入引用。**严禁机械地统一加在句尾或段尾**
3. **先写 `_cited.docx`**：输出到 `原名_cited.docx`，不覆盖原文件
4. **参考文献放在文章最后**：在结论章之后、文档末尾插入「参考文献」节及条目
5. **写入后用 `insert_citations.py` 排序**：因为手动插入时编号可能不连续或顺序不对

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

### 代码模板

```bash
python3 << 'PYEOF'
from docx import Document
from docx.oxml import OxmlElement
import os, re, copy

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def W_TAG(n): return f'{{{W}}}{n}'

def make_superscript(text):
    """创建上标 run"""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    va = OxmlElement('w:vertAlign'); va.set(W_TAG('val'), 'superscript')
    rPr.append(va); r.append(rPr)
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text; r.append(t)
    return r

def make_para(text, bold=False, center=False, indent=False):
    p = OxmlElement('w:p'); pPr = OxmlElement('w:pPr')
    if center:
        jc = OxmlElement('w:jc'); jc.set(W_TAG('val'), 'center'); pPr.append(jc)
    if indent:
        ind = OxmlElement('w:ind'); ind.set(W_TAG('left'), '420')
        ind.set(W_TAG('hanging'), '420'); pPr.append(ind)
    p.append(pPr)
    r = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    if bold: b = OxmlElement('w:b'); rPr.append(b)
    sz = OxmlElement('w:sz'); sz.set(W_TAG('val'), '28' if bold else '21')
    rPr.append(sz); r.append(rPr)
    t = OxmlElement('w:t'); t.text = text; r.append(t); p.append(r)
    return p

def insert_citation_inline(para, keyword, nums):
    """
    在段落中找到 keyword 的具体位置，紧跟其后插入上标引用标记。
    关键词是关键：它必须是需要文献支撑的那个具体术语/方法名/断言。
    例如：keyword='卷积神经网络' → '卷积神经网络[3,6]'
         keyword='随机森林模型' → '随机森林模型[5]'
    """
    marker = '[' + ','.join(str(n) for n in sorted(set(nums))) + ']'
    para_elem = para._element
    para_text = para.text

    idx = para_text.find(keyword)
    if idx < 0:
        para_elem.append(make_superscript(marker))
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
            sup_run = make_superscript(marker)
            run.addnext(sup_run)
            if after:
                after_run = OxmlElement('w:r')
                # 继承原始 run 的 rPr，但剥离上标属性
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
                sup_run.addnext(after_run)
            return
        char_pos = run_end
    para_elem.append(make_superscript(marker))

# 步骤 1：在需要引用的具体术语/断言后插入上标引用
# CITATIONS = [(段落索引, '需要引用的术语/论断关键词', [文献编号]), ...]
# keyword 必须精准定位到段落中具体被支撑的名词或说法
CITATIONS = [
    (13, '机器学习是一门依托数据驱动模型', [1, 2]),
    (15, '随机森林', [5]),
    # ...
]
for pi, kw, nums in CITATIONS:
    insert_citation_inline(doc.paragraphs[pi], kw, nums)

# 步骤 2：在「六、结论」之后插入参考文献节（文章最后）
body = doc.element.body
# 找到结论章的最后一个段落
conclusion_last = None
for i, p in enumerate(doc.paragraphs):
    if '六、结论' in p.text:
        # 从结论标题开始，找到该章最后一个非空段落
        for j in range(i, len(doc.paragraphs)):
            if j+1 >= len(doc.paragraphs) or (doc.paragraphs[j+1].text.strip() and
                (doc.paragraphs[j+1].text.strip().startswith('参考文献') or
                 any(doc.paragraphs[j+1].text.strip().startswith(p) for p in ['一','二','三','四','五','六','七','八']))):
                conclusion_last = doc.paragraphs[j]._element
                break
        break

# 在结论最后一个段落之后添加空行和参考文献
if conclusion_last is not None:
    insert_anchor = conclusion_last
else:
    # 回退：找 sectPr 之前最后的段落
    insert_anchor = None
    for child in body:
        if child.tag == W_TAG('sectPr'):
            break
        if child.tag == W_TAG('p'):
            insert_anchor = child

blank = OxmlElement('w:p')
if insert_anchor is not None:
    insert_anchor.addnext(blank)
    insert_anchor = blank

refs = ["[1] Author A, et al. Title[J]. Journal, Year, Vol: Pages.", ...]
ref_heading = make_para('参考文献', bold=True, center=True)
if insert_anchor is not None:
    insert_anchor.addnext(ref_heading)
else:
    body.append(ref_heading)
prev = ref_heading
for ref_text in refs:
    ref_p = make_para(ref_text, indent=True)
    prev.addnext(ref_p)
    prev = ref_p

doc.save(output_path)
print(f"✅ {os.path.basename(output_path)} — {len(CITATIONS)} 处引用, {len(refs)} 条文献")
PYEOF
```

### 结束后必须运行脚本

写入完成后，运行 `insert_citations.py` 做排序和交叉引用：

- **模式一（查找+排序）**：
  ```bash
  python3 ~/Documents/模式识别作业/insert-citations-skill/scripts/insert_citations.py "原名_cited.docx" --sort-only --ref-style superscript
  # 输出：原名_cited_cited.docx，删除中间文件：
  rm "原名_cited.docx" && mv "原名_cited_cited.docx" "原名_cited.docx"
  ```

- **模式三（全流程）**：
  ```bash
  python3 ~/Documents/模式识别作业/insert-citations-skill/scripts/insert_citations.py "原名_cited.docx" --ref-style superscript
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

```bash
python3 ~/Documents/模式识别作业/insert-citations-skill/scripts/insert_citations.py "文件路径"
```

输出：`原名_cited.docx`，**不覆盖原文**。打开后 Ctrl+A → F9 更新域。
