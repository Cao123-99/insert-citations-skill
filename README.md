# insert-citations

学术论文参考文献处理工具——为 .docx 论文**查找真实文献**、**按出现顺序排序**、**插入 Word/WPS 交叉引用域**。兼容 Word 和 WPS。

---

## 为什么需要它

写论文最繁琐的部分之一：引用了文献但编号乱了、正文标记和文末列表对不上、想加交叉引用但不会操作。这个工具让你在 Claude 对话中说一句话即可全流程处理——从 WebSearch 查找真实论文、到插入引用标记、到排序、到交叉引用域，全程自动化。

**核心设计原则：**
- 📄 **不编造文献**：每条引用都经过 WebSearch 检索、三要素（标题/作者/出处）验证
- 🔧 **不覆盖原文**：输出 `原名_cited.docx`，原文件毫发无损
- 🧹 **自动处理脏数据**：孤儿标记检测与清除、旧条目替换、双重前缀清理、附录边界保护
- 🎯 **精确插入**：引用标记紧跟被支撑的术语/方法，不扎堆段尾
- 📏 **格式统一**：所有模式输出一致的自动编号（`[1] [2] ...`）+ 无首行缩进

---

## 快速开始

```bash
cd insert-citations-skill
./install.sh
```

安装后在 Claude 对话中说一句话即可使用。安装过程会自动：
- 将 skill 注册到 `~/.claude/skills/`
- 安装 Python 依赖（`python-docx`）

---

## 场景速查

### 需要帮我找文献 → 对话模式（WebSearch 检索）

| 文档现状 | 说什么话 |
|---------|---------|
| 正文无引用标记，需要从零开始 | "帮我查找参考文献并按顺序排好"（模式一） |
| 同上，但要一直做到交叉引用 | "帮我全流程处理参考文献"（模式三） |
| 正文有孤儿 `[n]`（标记在但无文献列表） | 先 `--strip-markers` 清除，再走模式一或三 |

### 已有文献，只排版 → 命令行（无 WebSearch）

| 场景 | 文档现状 | 命令 | 内部执行顺序 |
|------|---------|------|-------------|
| **A** | 纯文本 `[1]`，手动编号文献，无交叉引用 | `paper.docx`（无参数） | ① 重排文献 → 去除手动编号 → 加自动编号 ② 更新正文 `[n]` 编号 ③ 创建书签 ④ `[n]` → REF 域 |
| **B** | 已有 REF 域，顺序乱，想保留交叉引用 | `paper.docx --sort-only` | ① 重排文献 + 自动编号 ② 更新 REF 域目标和显示文本 ③ 停（不创建书签、不插新 REF） |
| **C** | 已有 REF + 新加了文献（有 `[n]` 无交叉引用） | `paper.docx`（无参数）<br>仅更新编号 → `--sort-only` | ① 重排文献 + 自动编号 ② 更新已有 REF 域 ③ 更新纯文本 `[n]` 编号 ④ 创建书签 ⑤ 新 `[n]` → REF 域 |
| **D** | 清除所有标记，输出干净排序文献（无交叉引用） | `--strip-markers --sort-only` | ① 清除纯文本 `[n]` + REF 域 ② 重排文献 + 自动编号 |

---

## 对话模式：三种模式

### 模式一：查找文献并排序

> "帮我查找参考文献并按顺序排好"
> "补充文献然后排序"

1. 通读全文，提取需要文献支撑的断言
2. 检查文档是否已有「参考文献」标题
   - 有 → 提取旧条目文本，WebSearch 后合并去重
   - 无 → 问用户"放在哪章之后？"，以该章末为正文终点（不误扫附录）
3. 检测孤儿标记（正文有 `[n]` 但文末无文献列表）→ 有则先清除
4. WebSearch 多轮检索真实论文，逐条验证标题/作者/出处
5. 内联 Python 写入引用标记和参考文献
6. 运行 `--sort-only` 脚本排序，输出 `原名_cited.docx`
7. 结束，不插交叉引用域

### 模式二：加交叉引用域

> "给这篇论文加交叉引用"
> "帮我插入交叉引用域"

**前提：** 文档已有引用标记和参考文献列表。

**做什么：** 为参考文献创建 Word 自动编号 → 创建 `_Ref` 书签 → 正文 `[n]` 替换为 `REF \n \h` 域 → 验证书签与 REF 域一一对应。

**输出：** `原名_cited.docx`，打开后 `Ctrl+A → F9` 更新域。

### 模式三：全流程（查找 + 排序 + 交叉引用）

> "帮我全流程处理参考文献"
> "从头到尾处理这篇论文的引用"

1. 确认引用标记格式（上标 / 正文 / 两者）
2. 检查参考文献标题（有则提取旧条目供合并、无则问插入位置）
3. 诊断文档状态：

| 状态 | 含义 | 处理 |
|------|------|------|
| A | 无标记、无文献 | 直接 WebSearch |
| B | 有标记、无文献（孤儿） | `--strip-markers` 先清除，再 WebSearch |
| C | 有标记、有文献 | 提取旧文献 → WebSearch → 新旧合并去重 |

4. **始终执行 WebSearch**，检索支撑断言的真实论文
5. 写入引用标记和参考文献列表（复用原标题、清除旧条目）
6. 运行全流程脚本（排序 + 交叉引用），输出 `原名_cited.docx`
7. 提醒用户：`Ctrl+A → F9` 更新域

---

## 命令行速查

```bash
# 最常用：排序 + 交叉引用（已有文献只排版）
python3 scripts/insert_citations.py paper.docx

# 只排序（重排文献+自动编号，保留已有交叉引用）
python3 scripts/insert_citations.py paper.docx --sort-only

# 先清除所有标记，再排序（输出干净文献列表，无交叉引用）
python3 scripts/insert_citations.py paper.docx --strip-markers --sort-only

# 只加交叉引用
python3 scripts/insert_citations.py paper.docx --no-sort

# 只看报告不修改
python3 scripts/insert_citations.py paper.docx --dry-run

# 指定引用格式
python3 scripts/insert_citations.py paper.docx --ref-style superscript
python3 scripts/insert_citations.py paper.docx --ref-style plain

# 清除孤儿标记（纯文本 [n] + REF 域）
python3 scripts/insert_citations.py paper.docx --strip-markers

# 指定输出路径
python3 scripts/insert_citations.py paper.docx -o output.docx
```

### 完整参数表

| 参数 | 说明 |
|------|------|
| `--sort-only` | 只排序，不插交叉引用（先重排→再更新编号） |
| `--no-sort` | 只插交叉引用，不排序 |
| `--dry-run` | 仅分析报告，不修改文档 |
| `--ref-style superscript` | 仅匹配上标引用（中文论文最常见） |
| `--ref-style plain` | 仅匹配非上标（正文格式）引用 |
| `--ref-style both` | 匹配所有格式（默认） |
| `--strip-markers` | 清除正文中所有 `[n]` 引用标记 **和 REF 域**，保留其余文本 |
| `--strip-markers --sort-only` | 先清除标记，再排序（干净输出，无交叉引用） |
| `-o <path>` | 指定输出路径（默认 `原名_cited.docx`） |

---

## 核心能力

### 文献检索（citation-hunting）
- 根据断言文本构造检索式，WebSearch 多轮搜索
- 逐条验证三要素（标题 / 全部作者 / 出处）
- 支持中文和英文论文，默认 GB/T 7714 格式
- 找不到的文献明确标注，不编造

### 引用处理（insert_citations.py）
- **自动编号**：创建 `[1] [2] ...` Word 自动编号，无 numbering 文件的文档自动注入
- **交叉引用**：`REF \n \h` 编号项交叉引用，Word/WPS 原生支持
- **排序**：按正文首次出现顺序重新排列参考文献
- **孤儿标记清理**：检测并清除无对应文献列表的残留 `[n]` 和 REF 域
- **参考文献标题智能处理**：有标题则复用并清除旧条目；无标题则询问用户插入位置
- **附录边界保护**：无标题时以用户指定章末为正文终点，不误扫附录/致谢
- **格式统一**：所有模式输出一致的自动编号 + 无首行缩进

### 兼容性
- 引用格式：`[1]` `[[1]]` `[[17],[18]]` `[1,3-5]` 自动识别
- 跨 run 引用：`[` `4` `]` 分散在多个 XML run 中也能正确处理
- 表格引用：扫描正文段落和表格单元格，不遗漏
- 中英文论文：标题"参考文献""References"等自动识别
- 双重前缀清理：自动编号与手动 `[n]` 共存时自动去除重复前缀
- 无效引用过滤：图题/表题等短标签自动跳过
- 覆盖统计：输出引用数/总数/百分比/未引用编号

### 输出规范
- 所有模式输出**统一格式**：`[1] [2] ...` 自动编号，无悬挂缩进
- REF 域显示文本继承正文段落格式
- 切分后的后续文本保留原格式（剥离上标属性）

---

## 引用格式参考

详见 [`FORMATS.md`](./FORMATS.md)，覆盖 GB/T 7714-2015（中文论文最常用）、IEEE、APA 7th 等标准，支持期刊 `[J]`、会议 `[C]`、专著 `[M]`、学位论文 `[D]`、在线资源 `[EB/OL]`、专利 `[P]`、预印本 `[J/OL]` 等文献类型。

---

## 使用须知

### 对话模式

1. 确认 skill 已安装：`cd insert-citations-skill && ./install.sh`
2. 在对话中把 `.docx` 文件/路径导入
3. 说出触发语，例如"帮我全流程处理参考文献"
4. 按提示确认格式偏好和插入位置
5. 等待处理完成，下载 `原名_cited.docx`
6. 用 Word/WPS 打开，`Ctrl+A → F9` 更新所有域

### 命令行模式（已有文献，只排版）

```bash
# 纯文本 [1] + 手动文献 → 一键搞定排序和交叉引用
python3 scripts/insert_citations.py 论文.docx

# 已有交叉引用但顺序乱 → 只重排编号
python3 scripts/insert_citations.py 论文.docx --sort-only

# 已有交叉引用但缺了几条 → 补齐并更新
python3 scripts/insert_citations.py 论文.docx

# 先看看引用情况再决定
python3 scripts/insert_citations.py 论文.docx --dry-run

# 清除所有标记，输出干净文献列表
python3 scripts/insert_citations.py 论文.docx --strip-markers --sort-only
```

### 常见问题

**Q: 输出文件在哪？** A: 与原文件同目录，文件名为 `原名_cited.docx`，不覆盖原文。

**Q: 打开后编号没变化？** A: `Ctrl+A → F9` 更新域。交叉引用编号需要手动刷新。

**Q: 想指定输出路径？** A: 加 `-o` 参数，如 `-o /path/to/output.docx`。

**Q: 引用标记在正文中格式不对？** A: 运行时指定 `--ref-style`：上标用 `superscript`，正文格式用 `plain`。

**Q: 文档有附录或致谢，怕误扫？** A: 无标题时 Claude 会问你放在哪章之后，以该章末为界，不扫后续内容。

---

## 安装与更新

```bash
cd insert-citations-skill
./install.sh       # 安装
./update.sh        # 手动更新
./update.sh --schedule   # 每天自动检查更新
./update.sh --unschedule # 取消自动更新
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
│   └── SKILL.md          # 文献检索子技能（WebSearch 检索）
└── scripts/
    └── insert_citations.py   # 核心脚本
```

---

## 发给别人

把整个文件夹打包发过去，对方解压后运行：

```bash
cd insert-citations-skill
./install.sh
```
