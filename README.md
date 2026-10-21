# 小野的申请执行skill

给 **Codex / ChatGPT Work** 用的 Agent Skill：读一份民事判决书，一次生成整套**申请强制执行材料**（Word）。

## 能做什么

- 从判决书提取法院、案号、案由、当事人、判项、诉讼费/保全费、生效日
- 计算一般债务利息与迟延履行加倍利息，附利息计算表
- 按地区生成账户材料（普通 / 广州越秀 / 广州海珠）
- 自动排版：落款独立成页、字体字号、黑色无高亮
- 交付前一键复核：页数、落款分页、裁行、串案、字体、颜色、编号

## 默认产出

01 执行申请书、01附 利息计算表、02 执行授权委托书、04 法定代表人身份证明书（公司申请执行人）、05 账户书、07 法院提交材料清单、08 立案审批表、09 律师事务所所函、10 被告（第三人）信息表、11 诉讼文书送达地址确认书、12 执行风险提示（仅南海法院）。

要素式《强制执行申请书》默认不生成，需要时把脚本里的 `WITH_ELEMENT` 设为 `True`。

## 依赖

1. **Codex**（或支持 Agent Skills 的客户端）
2. **Python 3.10+**：`python-docx`、`pdfplumber`、`pywin32`（Windows，调 WPS 导 PDF 复核）、`Pillow`
3. **WPS Office**（Windows）：编辑 Word、导出 PDF；接入 **wps-office MCP** 可让 AI 直接读写导出（不接也能用脚本生成）
4. **一套 .docx 模板**（需自备，清单见下）
5. 可选配套 skills：`case-filing-documents`（立案材料）、`resumption-of-execution`（恢复执行材料）

## 安装

```bash
git clone https://github.com/xiaoye0714/小野的申请执行skill.git
```

把内容复制到 skills 目录：

- Windows：`%USERPROFILE%\.codex\skills\execution-application-documents\`
- macOS / Linux：`~/.codex/skills/execution-application-documents/`

重启 Codex 即可。

## 需要的模板（自备）

在 `scripts/build_exec_scaffold.py` 顶部把 `TPL` 改成你的模板目录，并准备同名 `.docx`：

```
01执行申请书 .docx
02执行授权委托书.docx
04法定代表人身份证明书(公司用这个).docx
06收款帐户确认书（最好加上银行卡复印件）.docx
07法院提交材料清单.docx
08立案审批表.docx
09律师事务所所函.docx
10被告（第三人）信息表.docx
11诉讼文书送达地址确认书.docx
12执行风险提示（南海才需要）.docx
强制执行申请书（要素式·样本）.docx
广州法院材料材料/越秀区退费表.docx
```

海珠区专用承诺书已随 `assets/海珠区/` 附带空白官式文件。

## 用法

把判决书 PDF 交给 Codex，并告诉它生效日期，说“按申请执行 skill 写一套执行材料”。

```bash
python scripts/build_exec_scaffold.py
python scripts/review_case.py "<案件文件夹>" 旧案关键词...
```

## 地区规则

- **南海法院**：额外输出 12 执行风险提示
- **广州法院**：账户书用越秀格式《诉讼费退费账户确认书》
- **海珠法院**：专用《退费申请及收款账户承诺书》；企业信用报告须在申请日前 7 个自然日内；全部文件日期统一为实际网上申请日
- **其他法院**：默认收取执行款银行账户确认书

## 自定义

- 律所/律师信息：`scripts/build_exec_scaffold.py` 顶部 `FIRM`、`LAWYER`、`ASSIST`、`FIRM_ADDR`、`FIRM_TEL`、`MOBILE`
- 每案只改顶部「案件数据」块：申请人/被执行人整段、`REASON1`、`FINDINGS`、`ITEMS`、`TOTAL`、地区开关

## 隐私与免责

本仓库为**脱敏公开版**，不含客户信息、真实案号或律所个人资料；模板请使用你自己的空白版本。生成的文书仅供办案参考，格式以受案法院最新规定为准，不构成法律意见。
