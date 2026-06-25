# insert-citations

学术论文参考文献处理工具——给 .docx 论文查找真实文献、按出现顺序排序、插入交叉引用域（Word/WPS 标准"编号项"交叉引用）。兼容 Word 和 WPS。

---

## 快速开始

```bash
cd insert-citations-skill
./install.sh
```

安装后在对话中说一句话即可使用。

---

## 三种模式

### 模式一：查找文献 + 排序

**触发语：** "帮我查找参考文献并按顺序排好"、"补充文献然后排序"

**做什么：**
1. 通读全文，找出需要文献支撑的断言
2. WebSearch 检索真实论文（每条三要素可验证）
3. 在正文具体断言/术语后插入上标 `[n]` 引用标记
4. 在结论章之后插入参考文献列表
5. 运行脚本按正文首次出现顺序重排编号

**输出：** `原名_cited.docx`

已经有序号混乱的参考文献？说"按出现顺序重新整理参考文献"只排序，不查文献。

---

### 模式二：加交叉引用域

**触发语：** "给这篇论文加交叉引用"、"帮我插入交叉引用域"

**做什么：**
- 为参考文献段落创建 Word 自动编号（`[%1]` 格式）
- 在参考文献上创建 `_Ref` 书签
- 将正文中的 `[n]` 替换为 `REF \n \h` 域（编号项交叉引用）
- 验证所有书签与 REF 域一一对应

**前提：** 论文已有引用标记和参考文献列表。

**输出：** `原名_cited.docx`，打开后 `Ctrl+A → F9` 更新域即可看到效果。

---

### 模式三：全流程

**触发语：** "帮我全流程处理参考文献"、"从头到尾处理这篇论文的引用"

**做什么：** 模式一（查文献+排序）→ 模式二（交叉引用），一步走完。

---

## 命令行

```bash
# 全流程（排序 + 交叉引用）
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

| 参数 | 说明 |
|------|------|
| `--sort-only` | 只排序，不插交叉引用 |
| `--no-sort` | 只插交叉引用，不排序 |
| `--dry-run` | 仅分析报告，不修改 |
| `--ref-style plain` | 仅匹配非上标引用 |
| `--ref-style superscript` | 仅匹配上标引用 |
| `-o <path>` | 指定输出路径（默认 `原名_cited.docx`） |

---

## 能力清单

| 功能 | 说明 |
|------|------|
| 交叉引用 | `REF \n \h` 编号项交叉引用，Word/WPS 原生支持 |
| 自动编号 | 自动创建 `[1] [2] ...` 格式编号，无 numbering 文件的文档自动注入 |
| 引用格式兼容 | `[1]` `[[1]]` `[[17],[18]]` `[1,3-5]` — 自动识别 |
| 上标/非上标过滤 | `--ref-style` 按格式筛选，减少误识别 |
| 格式继承 | REF 域显示文本保留上标格式，切分后的后续文本不受影响 |
| 表格引用 | 扫描正文段落 + 表格单元格，不遗漏 |
| 中英文论文 | 标题"参考文献""References"等自动识别 |
| 重复前缀清理 | 自动编号与手动 `[n]` 共存时去除双重前缀 |
| 无效引用过滤 | 图表标签自动跳过 |
| 跨 run 引用 | `[` `4` `]` 分散在多个 run 中也正确处理 |
| 覆盖统计 | 输出引用数/总数/百分比/未引用编号 |

---

## 使用须知

1. **引用不要扎堆**：每处引用只放一条文献（`[5]` 而非 `[5,9]`），不同论断用不同引用
2. **引用紧跟术语/方法名**：放在被支撑的术语或断言后，不堆在段尾或句尾
3. **参考文献放文末**：结论章之后、文档末尾
4. **打开后更新域**：`Ctrl+A → F9`，交叉引用编号才会正常显示

---

## 发给别人

把整个文件夹打包发过去，对方解压后运行：

```bash
cd insert-citations-skill
./install.sh
```

---

## 更新

安装后是一个 git 仓库，symlink 到 `~/.claude/skills/`，更新只需 `git pull`：

```bash
cd insert-citations-skill
./update.sh              # 手动更新到最新版
./update.sh --schedule   # 设置每天自动检查（macOS/Linux cron）
./update.sh --unschedule # 取消自动检查
```

---

## 文件结构

```
insert-citations-skill/
├── SKILL.md              # 主 skill 文件（Claude 执行规则）
├── FORMATS.md            # 引用格式参考（GB/T 7714、IEEE、APA）
├── README.md             # 本文件
├── install.sh            # 一键安装
├── update.sh             # 一键更新 + 自动更新调度
├── citation-hunting/
│   └── SKILL.md          # 文献检索子技能
└── scripts/
    └── insert_citations.py   # 核心脚本
```

```
insert-citations-skill/
├── SKILL.md              # 主 skill 文件（Claude 执行规则）
├── FORMATS.md            # 引用格式参考（GB/T 7714、IEEE、APA）
├── README.md             # 本文件
├── install.sh            # 一键安装
├── citation-hunting/
│   └── SKILL.md          # 文献检索子技能
└── scripts/
    └── insert_citations.py   # 核心脚本
```
