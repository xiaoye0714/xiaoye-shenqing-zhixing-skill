"""执行材料交付前自检。

用法：python check_outputs.py <案件文件夹> [要排查的旧案关键词 ...]

逐份检查：
1. 字体字号：每个非空 run 是否都有显式的 rFonts 与 sz（并与模板对照，区分“我改坏的”和“模板本来如此”）；
2. 编号重复：段落带自动编号（numId≠0）时，文字里又手写序号；
3. 颜色/高亮：非黑色字体、文字高亮；
4. 串案：出现多个案号，或命中了传入的旧案关键词；
5. 落款：“暂合计”与“此致”之间有多余空段落；
6. 专项版式：02 委托事项的被执行人名称、案由下划线是否完整，05 户名前空格是否为 5 个。
"""

import re
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

try:  # 避免 PowerShell 控制台按 GBK 输出导致中文乱码
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

TEMPLATES = Path(r"C:\你的路径\执行申请模板")
ELEMENT_TPL = TEMPLATES / "强制执行申请书（要素式·样本）.docx"
# 法院制式表格自带的中性色（231F20，肉眼等同黑字）不算“非黑字”
OK_COLORS = (None, "000000", "auto", "231F20", "231f20")
# 半角括号也要认：模板样本里常见 "(20XX) 粤 XXXX 民初 XXXX 号" 这种写法
CASE_NO = re.compile(r"[（(]20\d\d[)）][^号]{0,24}号")


def runs_of(doc):
    for p in doc.paragraphs:
        for r in p.runs:
            yield p, r
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:
                    for r in p.runs:
                        yield p, r


def font_ok(run):
    rPr = run._element.rPr
    if rPr is None:
        return False
    return rPr.find(qn("w:rFonts")) is not None and rPr.find(qn("w:sz")) is not None


def count_missing_fonts(path):
    d = Document(path)
    n = 0
    for _, r in runs_of(d):
        if r.text.strip() and not font_ok(r):
            n += 1
    return n


def missing_font_texts(path):
    """返回缺字体字号的 run 文本集合（按文本比对，避免合并单元格重复计数造成的误报）。"""
    d = Document(path)
    out = set()
    for _, r in runs_of(d):
        if r.text.strip() and not font_ok(r):
            out.add(r.text.strip())
    return out


def check(path, avoid):
    d = Document(path)
    msgs = []
    notes = []
    missing = count_missing_fonts(path)
    if missing:
        prefix = path.name[:2]
        # 要素式申请书对阵要素式模板；其余按文件名前缀找同名模板
        tpl = [ELEMENT_TPL] if "要素式" in path.name else sorted(TEMPLATES.glob(prefix + "*.docx"))
        if not tpl:
            msgs.append(f"缺字体字号 {missing} 处（未找到对应模板，无法比对）")
        else:
            mine = missing_font_texts(path)
            theirs = missing_font_texts(tpl[0])
            added = mine - theirs
            if not added:
                notes.append(f"提示：缺字体字号 {missing} 处，均为模板原有，属原样")
            else:
                msgs.append(
                    f"缺字体字号 {missing} 处，其中 {len(added)} 条为新增"
                    "（疑似改动丢失，需修）：" + "；".join(sorted(added)[:5])
                )

    dup = 0
    for p in d.paragraphs:
        if not re.match(r"^[1-9][.、]", p.text.strip()):
            continue
        pPr = p._p.find(qn("w:pPr"))
        if pPr is None:
            continue
        npr = pPr.find(qn("w:numPr"))
        if npr is None:
            continue
        nid = npr.find(qn("w:numId"))
        if nid is not None and nid.get(qn("w:val")) not in (None, "0"):
            dup += 1
    if dup:
        msgs.append(f"编号重复 {dup} 处（自动编号 + 手写序号）")

    colored = 0
    for _, r in runs_of(d):
        rPr = r._element.rPr
        if rPr is None:
            continue
        if rPr.find(qn("w:highlight")) is not None:
            colored += 1
        c = rPr.find(qn("w:color"))
        if c is not None and c.get(qn("w:val")) not in OK_COLORS:
            colored += 1
    if colored:
        msgs.append(f"非黑字/高亮 {colored} 处")

    if path.name.startswith("02_"):
        p = next((q for q in d.paragraphs if q.text.startswith("现委托以上受委托人在我方与")), None)
        if p is None:
            msgs.append("未找到授权委托书“委托事项”正文")
        else:
            prefix = "现委托以上受委托人在我方与"
            suffix = "申请强制执行一案中"
            between = "之间"
            text = p.text
            if prefix in text and between in text and suffix in text:
                party_start = text.index(prefix) + len(prefix)
                party_end = text.index(between, party_start)
                cause_end = text.index(suffix, party_end + len(between))
                expected = [
                    text[party_start:party_end],
                    text[party_end + len(between):cause_end],
                ]
                groups, current = [], []
                for r in p.runs:
                    if r.underline:
                        current.append(r.text)
                    elif current:
                        groups.append("".join(current))
                        current = []
                if current:
                    groups.append("".join(current))
                missing_underlines = [s for s in expected if s and s not in groups]
                if missing_underlines:
                    msgs.append("02委托事项下划线缺失：" + "；".join(missing_underlines))
            else:
                msgs.append("02委托事项正文结构不完整，无法核对下划线")

    if path.name.startswith("05_收款帐户确认书_"):
        p = next((q for q in d.paragraphs if q.text.startswith("户           名：")), None)
        if p is None:
            msgs.append("未找到05账户确认书“户名”行")
        else:
            tail = p.text.split("：", 1)[1] if "：" in p.text else ""
            lead = len(tail) - len(tail.lstrip(" "))
            if lead != 5:
                msgs.append(f"05户名前空格应为5个，当前为{lead}个")

    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    text = re.sub(r"<[^>]+>", "", xml)
    nos = sorted(set(CASE_NO.findall(text)))
    if len(nos) > 1:
        msgs.append("出现多个案号：" + "，".join(nos))
    hits = [a for a in avoid if a and a in text]
    if hits:
        msgs.append("疑似旧案残留：" + "，".join(hits))

    ps = d.paragraphs
    i_sum = next((i for i, q in enumerate(ps) if q.text.startswith("（以上各项暂合计")), None)
    i_zhi = next((i for i, q in enumerate(ps) if q.text.startswith("此致")), None)
    if i_sum is not None and i_zhi is not None:
        empty = sum(1 for q in ps[i_sum + 1 : i_zhi] if not q.text.strip())
        if empty:
            msgs.append(f"落款前多余空段 {empty} 个")

    return msgs, notes


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    folder = Path(sys.argv[1])
    avoid = sys.argv[2:]
    files = sorted(
        f for f in folder.glob("*.docx") if not f.name.startswith("~$")
    )
    print(f"检查目录：{folder}")
    bad = 0
    for f in files:
        try:
            msgs, notes = check(f, avoid)
        except Exception as exc:  # 打不开的（锁文件、加密、损坏）单列，不中断整批
            bad += 1
            print(f"  打不开  {f.name}  -> {type(exc).__name__}: {exc}")
            continue
        if msgs:
            bad += 1
            print("  需处理  " + f.name + "  ->  " + "；".join(msgs))
        else:
            print("  通过    " + f.name + (("  " + "；".join(notes)) if notes else ""))
    print(f"共 {len(files)} 份，需处理 {bad} 份")
