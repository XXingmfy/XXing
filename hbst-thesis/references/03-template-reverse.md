# 河北科技师范学院模板画像与逆向方法

> 本文件把从真实论文（《基于微信小程序的手机租赁系统设计与实现》，学士学位论文）逆向出的格式事实沉淀为**模板画像**。用户提供学校模板时按"逆向方法"重新提取并覆盖本画像；未提供模板时直接采用本画像作为默认基线。**学校下发模板/导师要求与本画像冲突时，以学校模板为准。**

## 1. 模板画像（默认基线，来自真实成稿逆向）

### 1.1 页面与分节

| 项 | 值 |
|---|---|
| 纸张 | A4（宽 11906 twips × 高 16838 twips） |
| 上边距 / 下边距 | 1418 / 1134 twips（约 2.5 / 2.0 cm） |
| 左 / 右边距 | 1134 / 1134 twips（约 2.0 cm） |
| 装订线 | 284 twips（gutter，约 0.5 cm） |
| 页眉距 / 页脚距 | 851 / 992 twips（1.5 / 1.75 cm 档） |

分节结构（6 节，各节独立控制页眉页脚与页码）：

| 节 | 内容 | 页码 |
|---|---|---|
| 1 | 封面（含校名横幅图、"学士学位论文"、题目、姓名/学号/院系/专业/指导教师、日期"二〇二六年五月二十日"式汉字日期） | 无页码、无页眉 |
| 2 | 学位论文原创性声明 + 学位论文版权使用授权书 | 无页码、无页眉 |
| 3 | 中文摘要（摘　　要 + 关键词） | 罗马 i 起（footer 为 PAGE 域，居中） |
| 4 | ABSTRACT（+ Key words） | 罗马续 |
| 5 | 目录（包在 `w:sdt` 中，含 `TOC \o "1-3"` 域，显示到三级） | 罗马续 |
| 6 | 正文第1章 ~ 附录（同一节） | 阿拉伯 1 起（第 1 章 = 逻辑第 1 页） |

### 1.2 样式定义（styles.xml 逆向）

| 用途 | 样式（中文名） | 中文字体 | 西文/数字 | 字号 | 对齐/其他 |
|---|---|---|---|---|---|
| 正文 | 论文正文q | 宋体 | Times New Roman | 小四（24 半磅） | 两端对齐；首行缩进 2 字符（firstLineChars=200 / firstLine=480）；行距固定 20 磅（line=400 exact） |
| 一级标题 | Heading 1（标题 1） | 黑体 | Times New Roman | 三号（30 半磅） | 加粗、居中；段前 800/段后 400 |
| 二级标题 | Heading 2（标题 2） | 黑体 | Times New Roman | 四号（28 半磅） | 左对齐；段前 480/段后 120 |
| 三级标题 | Heading 3（标题 3） | 黑体 | Times New Roman | 小四/13pt（26 半磅） | 加粗、左对齐；段前 240/段后 120 |
| 图题注 | 图标题q | 黑体 | Times New Roman | 五号（22 半磅） | 居中；段前 120/段后 240 |
| 表题注 | Subtitle | （同表文） | — | 五号档 | 居中、在表上方 |
| 参考文献正文 | 参考文献正文 | 宋体 | Times New Roman | 小四 | 基于正文；自动编号 `[n]`（numId=6，lvlText=`[%1] `） |

标题编号体系：章 `第X章　标题`（全角空格）、节 `X.X`、小节 `X.X.X`；"摘　　要 / ABSTRACT / 参考文献 / 致　　谢 / 附　　录"为不带编号的一级标题（加全角空格居中）。
非编号一级标题文字在源文档中以单字 run 形态存在（"摘/　/要"分 run），**渲染正常属字体混排，不是错误**；复现时按整词写入即可，不做手工分散对齐。

### 1.3 表格约定

- 数据库表统一 **7 列**：编号 | 字段名 | 类型 | 长度 | 是否非空 | 是否主键 | 注释。
- **三线表**：仅上、下两条粗线（sz 12，约 1.5pt）+ 栏目线（表头下细线）；无竖线、无内部横线（XML 表现：tblBorders 只有 top/bottom=single sz12，insideH/insideV=nil，left/right=nil；数据行单元格无 tcBorders，靠表头行单元格下边框或表头样式呈现栏目线——以逆向结果为准）。
- 跨页表：首行 `tblHeader` 重复表头 + 正文出现"续表4.X"独立段；**注意真实样例中行未设 `cantSplit`，属于可改进项**，本技能默认按 writing-standards §6.4 要求行不跨页断行。
- 表标题用 Subtitle 样式在表**上方**；字段注释要写全枚举（"0-未使用、1-已使用""USER-普通用户"）。

### 1.4 页眉页脚

- 封面/声明节：无页眉页脚。
- 摘要~附录：页脚居中 PAGE 域页码；正文节页眉居中"河北科技师范学院学士学位论文"（宋体小五）。
- 真实样例页脚域缓存了旧值（I/II/V/17），导出 PDF 前必须全选 F9 更新域；`updateFields` 未开启时不会自动刷。

### 1.5 元数据

真实样例 docProps：无 title/keywords，author 为"Administrator"残留、lastModifiedBy 为作者名——交付前应清理文档属性（见 writing-standards §9）。

## 2. 逆向方法（当用户给了学校模板时执行）

若拿到学校模板 `.docx`，不依赖肉眼看格式，按如下提取画像：

1. **拆包**：把 `.docx` 当 zip 解开，读 `word/document.xml`、`word/styles.xml`、`word/numbering.xml`、`word/settings.xml`、各 `header*.xml`/`footer*.xml`。
2. **页面与分节**：找所有 `w:sectPr`（含段内与 body 末尾），读 `pgSz/pgMar/gutter/header/footer`，记录每节 `pgNumType`（fmt/start）与 `titlePg`。
3. **样式表**：列 `styles.xml` 全部段落样式；重点记正文/标题1-3/图题注/表题注/参考文献的实际字体、字号（`w:sz` 半磅）、对齐、缩进（firstLine/firstLineChars）、行距（spacing line/lineRule）。
4. **标题编号与 TOC**：确认章标题文字格式、目录是否 `TOC` 域（包在 `w:sdt` 内找 `instrText`）。
5. **页眉页脚**：逐 header/footer 文件读文本与 PAGE 域格式开关（`PAGE \* MERGEFORMAT` 等）。
6. **表格**：抽样读 `w:tblPr/tblBorders`、首行 `tblHeader`、行 `cantSplit`、`w:tblW` 宽度；确认是否三线表。
7. **编号**：`numbering.xml` 里找参考文献用的 `abstractNum`（lvlText `[%1]`）。
8. **产出**：`paper-context/template-profile.yaml`，字段含：page(section list: margins/pgNumType)、fonts(body/heading1..3/caption/tableCaption/refBody)、toc(levels, isField)、headerFooter(text per section)、tables(columns/threeLine/headerRepeat/continuation)、references(numbering format)、docProps(cleanup notes)。并把与默认基线不同之处列成 diff 说明，写作全程以 profile 为准。

## 3. 默认章节映射（该校成稿实例）

| 逻辑页 | 内容 |
|---|---|
| 封面 1–2 | 封面、声明+授权书 |
| i–iii | 摘要 / ABSTRACT / 目录（真实样例 3 页目录） |
| 1–43 | 第1章 绪论 ~ 第6章 系统测试（六章在同一节连续排版） |
| 44–47 | 结论 / 参考文献(17 条) / 致谢 / 附录 |

真实样例约 47 逻辑页 / 54 物理页；正文含 33 个图题注 + 25 个表题注（数据库表 22 + 测试表 3，其中 4 张跨页拆"续表"）。篇幅因题而异，**达标判据用字数而非页数**（见 SKILL.md 闸门 1）。
