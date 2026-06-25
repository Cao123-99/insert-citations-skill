# insert-citations

学术论文参考文献处理工具——给 .docx 论文查找真实文献、按出现顺序排序、插入交叉引用域。兼容 Word 和 WPS。

---

## 快速开始

```bash
cd insert-citations-skill
./install.sh
```

安装后在对话中说一句话即可使用（见下方）。

---

## 三种模式

### 模式一：查找文献 + 排序

**触发语：**
- "帮我查找参考文献并按顺序排好"
- "补充文献然后排序"
- "按出现顺序重新整理参考文献"
- "参考文献顺序乱了帮我重新排"

**做什么：**
1. 通读全文，找出需要文献支撑的断言
2. WebSearch 检索真实论文（不编造）
3. 在正文插入 `[n]` 引用标记，在文末写入参考文献
4. 按正文首次出现顺序重排编号

**输出：** `原名_cited.docx`

---

### 模式二：加交叉引用域

**触发语：**
- "给这篇论文加交叉引用"
- "帮我插入交叉引用域"

**做什么：**
- 在参考文献段落上创建 `_Ref` 书签
- 将正文中的 `[n]` 替换为 REF 域
- 验证所有书签与 REF 域一一对应

**前提：** 论文已有引用标记和参考文献列表。

**输出：** `原名_cited.docx`，打开后按 `Ctrl+A` → `F9` 更新域。

---

### 模式三：全流程

**触发语：**
- "帮我全流程处理参考文献"
- "从头到尾处理这篇论文的引用"

**做什么：** 模式一（查文献+排序）→ 模式二（交叉引用），一步走完。

---

## 命令行

```bash
# 全部（排序 + 交叉引用）
python3 scripts/insert_citations.py paper.docx

# 只排序
python3 scripts/insert_citations.py paper.docx --sort-only

# 只加交叉引用
python3 scripts/insert_citations.py paper.docx --no-sort

# 只看报告不修改
python3 scripts/insert_citations.py paper.docx --dry-run

# 只匹配上标格式的引用（中文论文常见）
python3 scripts/insert_citations.py paper.docx --ref-style superscript

# 只匹配非上标格式的引用
python3 scripts/insert_citations.py paper.docx --ref-style plain
```

---

## 能力清单

| 功能 | 说明 |
|------|------|
| 引用格式兼容 | `[1]` `[[1]]` `[[17],[18]]` `[1,3-5]` — 自动识别，无需指定 |
| 表格引用 | 扫描正文段落 + 表格单元格，不遗漏 |
| 中英文论文 | 标题"参考文献""References""7.参考文献"等自动识别 |
| 格式继承 | 新插入的 REF 域与段落正文格式一致，不强制上标 |
| 重复前缀清理 | 自动编号与手动 `[n]` 共存时自动去除双重前缀 |
| 无效引用过滤 | 图表标签（如 `[20,40,60,80,100]`）自动跳过 |
| 跨 run 引用 | `[` `4` `]` 分散在多个 run 中也正确处理 |
| 覆盖统计 | 输出引用数/总数/百分比/未引用编号 |

---

## 发给别人

把整个文件夹打包发过去，对方解压后运行：

```bash
cd insert-citations-skill
./install.sh
```

即可安装使用。

---

## 文件结构

```
insert-citations-skill/
├── SKILL.md              # 主 skill 文件
├── FORMATS.md            # 引用格式参考（GB/T 7714、IEEE、APA）
├── README.md             # 本文件
├── install.sh            # 一键安装
├── citation-hunting/
│   └── SKILL.md          # 文献检索子技能
└── scripts/
    └── insert_citations.py   # 核心脚本
```
