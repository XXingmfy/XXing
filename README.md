# XXing · Skills 仓库

存放个人打磨的 Agent Skills。当前收录：

| Skill | 说明 |
|---|---|
| [hbst-thesis](hbst-thesis/SKILL.md) | 依据真实项目源码，生成/改写**可直接提交**的河北科技师范学院本科毕业论文 `.docx`（含降 AIGC 痕迹优化） |

---

## hbst-thesis —— 本科毕业论文写作 Skill

把一套真实源码变成一篇符合学校模板、可直接提交的 Word 论文。内置从真实成稿逆向出的河北科技师范学院格式画像，也兼容采用同款结构的其他本科院校（封面–声明–摘要–目录–六章正文–参考文献–致谢–附录）。

### 它能做什么

- **读源码建证据**：扫描技术栈 / 功能模块 / 角色权限 / 数据库表 / API / 测试材料 / 截图，建立证据工作区，全程不编造（缺什么列什么）。
- **按规范成文**：内置《毕业论文写作通用规范》——分节页码 / 字体字号 / 六章骨架 / 三线表 / 图表编号 / 语言红线 / 提交前检查清单，逐章生成。
- **降 AIGC 痕迹**：内置《降 AIGC 痕迹写作规范 v2.0》——语义层重写优先、句层去模板、文档层嗓音一致，附 S1–S12 自检与披露模板。规范全文见 `references/`，一个字不删。
- **产出 Word**：`build_docx.py` 一键从 Markdown 稿生成含 6 分节、封面、声明页、罗马/阿拉伯页码域、页眉、自动目录域、三线表的 `.docx`；`check_*.py` 提供提交前确定性闸门。

### 工作流程

```
阶段一  通读源码 → 证据工作区 paper-context/（spec + 图表登记 + 缺证清单）
阶段二  按通用写作规范逐章成文 → 随章做降 AIGC 优化（A 语义层 > B 句层 > C 文档层）
阶段三  build_docx.py 产出 .docx → 三道闸门检查 → Word 中 F9 更新域 → 导出 PDF 版面核对
```

### 安装（三选一）

**通用（任意支持 Skills 的 Agent，如 Claude Code / Codex / ZCode）**：把 `hbst-thesis/` 目录放入你的 skills 目录：

| 位置 | 说明 |
|---|---|
| `<项目>/.zcode/skills/hbst-thesis/` | 项目级（ZCode） |
| `<项目>/.agents/skills/hbst-thesis/` | 项目级（通用） |
| `~/.agents/skills/hbst-thesis/` | 用户级（全局，推荐） |

**ZCode 用户**：也可在设置中把本仓库添加为 skill 市场/源后启用 `hbst-thesis`。

### 使用示例

```text
用 hbst-thesis 根据这个项目写毕业论文：
1. 项目代码路径：/path/to/project
2. 学校模板：暂无（用内置河北科技师范学院模板）
3. 数据库结构：/path/to/sql 或实体类
4. 系统截图：/path/to/screenshots
5. 测试材料：/path/to/tests
要求：不要编造功能/字段/接口/测试/文献；先给大纲和图表计划确认，再逐章成文；
第五章截图用真实截图，缺失就登记占位；最后产出 .docx 并附缺证清单。
```

或直接说：

```text
用 hbst-thesis 把 /path/to/project 写成论文，先出大纲和图表计划。
```

对已有论文做去 AI 味：

```text
用 hbst-thesis 优化这篇论文的 AI 痕迹：不改变事实与数据，按 A1 脱稿重写高危段，
不虚构文献与数据，输出 S1–S12 自检。
```

### 目录结构

```
hbst-thesis/
├── SKILL.md                          # 主流程与触发说明（入口）
├── references/
│   ├── 01-writing-standards.md       # ★《毕业论文写作通用规范》全文
│   ├── 02-humanize-standards.md      # ★《降 AIGC 痕迹写作规范 v2.0》全文
│   ├── 03-template-reverse.md        # 河北科技师范学院模板画像 + 模板逆向方法
│   ├── 04-thesis-voice.md            # 论文语气：可用句式 vs 禁入正文的措辞
│   └── 05-docx-production.md         # docx 分节/页码域/三线表/续表落法
└── scripts/
    ├── build_docx.py                 # Markdown 稿 → 可提交 .docx（6 分节/域/三线表）
    ├── evidence_scan.py              # 源码证据扫描
    ├── check_text_quality.py         # 文本闸门（缺证措辞/重复标点/字数/引用闭合）
    └── check_docx_structure.py       # docx 结构闸门（分节/标题样式/TOC域/PAGE域/三线表）
```

### 依赖

- Python 3.8+；`pip install python-docx`（`build_docx.py` 与 docx 检查用）；其余尽量标准库。
- 生成 `.docx` 后用 Word/WPS 打开，**全选按 F9 更新目录与页码域**，再导出 PDF 核对版面。

### 诚实声明

- Skill 内置的两份规范来源于真实论文分析 + 公开检测研究实证的整理，使用时请遵守所在学校/期刊的 AI 使用政策；**需要披露的场景请如实披露**。
- Skill 生成的论文内容是否合规、是否通过查重/检测，责任在使用者；Skill 不承诺"保证过检"。

### 本地开发

```bash
# 修改后做完整性检查
python -m py_compile hbst-thesis/scripts/*.py
```

---

## License

本仓库内容供个人学习使用；引用内置规范或脚本请保留来源说明。
