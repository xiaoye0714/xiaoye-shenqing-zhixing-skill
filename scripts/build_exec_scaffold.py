# -*- coding: utf-8 -*-
"""执行申请材料生成器（按内容定位，模板换成"已填样本"也不崩）。

用法（新案件 4 步）：
1. 复制本文件到 %TEMP%（如 build_xxx.py）；
2. 只改下面「案件数据」块；
3. 运行，一次生成适用的那几份；
4. 跑 review_case.py 复核（已内含交付前自检），只有报 [!] 的才渲染 PNG 看图。

约定：
- 要素式默认不生成，用户明确要时才把 WITH_ELEMENT 置 True；
- 定位段落一律按内容锚点，不按下标（模板会被换成旧样本/样本那种已填样本）；
- 落款"申请执行人："必须完全匹配，不能用 startswith（会命中首部当事人行）；
- 删多余的请求项段落从末尾往前删；
- 写盘前用 open(path,'r+b') 试开目标文件，确认没被 WPS 占用。
"""

import copy
import io
import os
import re
import shutil
from decimal import Decimal, InvalidOperation

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# ================= 案件数据（只改这一块） =================
TPL = r"C:\你的路径\执行申请模板"
OUT = r"C:\Users\liyeye\Desktop\20261015某某诉某某某买卖合同纠纷（准备执行）"
JUDGMENT = r"C:\Users\liyeye\Desktop\判决书.pdf"

COURT_FULL = "广东省佛山市南海区人民法院"      # 抬头用全称
COURT_SHORT = "佛山市南海区人民法院"           # 此致/所函/表单抬头
CASE_NO = "（20XX）粤XXXX民初XXXX号"
CAUSE = "买卖合同纠纷"
APPLICANT = "某某公司"                       # 用于文件名
DOC_TYPE = "民事判决书"                      # "民事判决书" / "民事调解书" / "民事裁定书"等

APPLICANT_BLOCK = "申请执行人：某某公司，住所地……，统一社会信用代码……。法定代表人：某某，职务：总经理。"
DEF_BLOCKS = [
    "被执行人：某某，男，……，公民身份号码……。",
]
REASON1 = "申请执行人与被执行人" + CAUSE + "一案，经贵院审理作出" + CASE_NO + "民事判决，该判决已于……发生法律效力。"
FINDINGS = "生效判决确定：……现判决指定的履行期限已于……届满，被执行人未在指定期限内履行任何给付义务，经申请执行人催告后仍拒不履行，已严重侵害申请执行人合法权益。依据《中华人民共和国民事诉讼法》第二百四十七条、第二百六十四条及《最高人民法院关于执行程序中计算迟延履行期间的债务利息适用法律若干问题的解释》之规定，特向贵院申请强制执行，请求强制被执行人履行生效判决确定的全部义务，并加倍支付迟延履行期间的债务利息。具体执行请求如下："
ITEMS = [
    "1.强制被执行人向申请执行人支付……；",
    "2.强制被执行人向申请执行人支付……；",
]
TOTAL = "0.00"

LAWYER = "示例律师"
ASSIST = "示例助理"
FIRM = "示例律师事务所"
FIRM_ADDR = "（律所地址）"
FIRM_TEL = "0000-0000000"
MOBILE = "13800000000"
LAWYER_ID = "000000000000000000"               # 11 送达确认书代理人身份号码

# 主体类型：自然人申请执行人填下面两行；公司填 COMPANY 那组
APPLICANT_ADDR = ""
APPLICANT_ID = ""
APPLICANT_MOBILE = ""                         # 申请执行人本人手机号：11 送达确认书当事人栏必填，不掌握就问用户
APPLICANT_IS_COMPANY = True
APPLICANT_IS_INDIVIDUAL_BUSINESS = False      # 个体工商户：使用负责人身份证明书
APPLICANT_CODE = ""
APPLICANT_REP = ""
APPLICANT_REP_TITLE = "总经理"
APPLICANT_ACCOUNT = ""                        # 海珠退费承诺书：收款账号，不掌握留空
APPLICANT_BANK = ""                           # 海珠退费承诺书：开户行，必须具体到支行
DEFS_ARE_COMPANIES = [False]                  # 与 DEF_BLOCKS 一一对应

# 地区开关
RISK_NOTICE = False                           # 南海才出 12 执行风险提示
ACCOUNT = "收款书"                            # "收款书" / "越秀退费书" / "海珠退费承诺书"
WITH_ELEMENT = False                          # 要素式默认不出
DOC_DATE = ""                                 # 海珠填实际网上申请日期，如 "2026年10月25日"；普通案件留空
REFUND_TOTAL = ""                             # 海珠退费承诺书必填：可退诉讼费合计（只填数值）
REFUND_BREAKDOWN = ""                         # 可选：如 "其中案件受理费XXX元、保全费XXX元"

# 以下三项给规则里需要文字的地方用（要素式/立案审批表），也可直接在函数里写死
BIAODI = "本金及利息"                         # 立案审批表"标的额"
EFFECTIVE_DATE = "2026年1月1日"               # 生效日
# ================= 数据块结束 =================


def rpr_of(r):
    el = r.find(qn("w:rPr"))
    if el is None:
        el = OxmlElement("w:rPr")
        r.insert(0, el)
    return el


def set_u(r, val):
    el = rpr_of(r).find(qn("w:u"))
    if val is None:
        if el is not None:
            rpr_of(r).remove(el)
        return
    if el is None:
        el = OxmlElement("w:u")
        rpr_of(r).append(el)
    el.set(qn("w:val"), "single" if val else "none")


def set_b(r, val):
    el = rpr_of(r).find(qn("w:b"))
    if val is None:
        if el is not None:
            rpr_of(r).remove(el)
        return
    if el is None:
        el = OxmlElement("w:b")
        rpr_of(r).append(el)
    el.set(qn("w:val"), "1" if val else "0")


def set_color(r, val="000000"):
    el = rpr_of(r).find(qn("w:color"))
    if el is None:
        el = OxmlElement("w:color")
        rpr_of(r).append(el)
    el.set(qn("w:val"), val)
    for bad in ("w:themeColor", "w:themeTint", "w:themeShade"):
        if el.get(qn(bad)) is not None:
            el.attrib.pop(qn(bad))


def kill_hl(r):
    el = rpr_of(r).find(qn("w:highlight"))
    if el is not None:
        rpr_of(r).remove(el)


def set_sz(run, pt):
    rPr = rpr_of(run._r)
    for tag in ("w:sz", "w:szCs"):
        for el in rPr.findall(qn(tag)):
            rPr.remove(el)
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), str(int(pt * 2)))
        rPr.append(el)


def make_run(tmpl_run, text, u=None, b=None):
    r = copy.deepcopy(tmpl_run)
    for child in list(r):
        if child.tag != qn("w:rPr"):
            r.remove(child)
    set_u(r, u)
    set_b(r, b)
    set_color(r)
    kill_hl(r)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def set_segments(p, segs):
    runs = p._p.findall(qn("w:r"))
    assert runs, "空段落无法定位格式模板: %r" % p.text
    tmpl = copy.deepcopy(runs[0])
    for r in runs:
        p._p.remove(r)
    for text, u, b in segs:
        p._p.append(make_run(tmpl, text, u, b))


TAGS = ("w:rFonts", "w:sz", "w:szCs")


def style_fallback(p):
    got = {t: None for t in TAGS}
    chain, st = [], p.style
    while st is not None:
        chain.append(st)
        st = st.base_style
    for st_ in chain:
        el = st_.element.find(qn("w:rPr"))
        if el is None:
            continue
        for t in TAGS:
            if got[t] is None:
                f = el.find(qn(t))
                if f is not None:
                    got[t] = copy.deepcopy(f)
    try:
        dd = p.part.document.styles.element.find(qn("w:docDefaults"))
        if dd is not None:
            rp = dd.find(qn("w:rPrDefault"))
            rp = rp.find(qn("w:rPr")) if rp is not None else None
            if rp is not None:
                for t in TAGS:
                    if got[t] is None:
                        f = rp.find(qn(t))
                        if f is not None:
                            got[t] = copy.deepcopy(f)
    except Exception:
        pass
    return got


def fill_style_font(run_el, p):
    rPr = run_el.find(qn("w:rPr"))
    fb = style_fallback(p)
    for t in TAGS:
        if rPr.find(qn(t)) is None and fb.get(t) is not None:
            rPr.append(copy.deepcopy(fb[t]))


def set_para_text(p, text):
    runs = p._p.findall(qn("w:r"))
    assert runs, "空段落无法定位格式模板"
    tmpl = copy.deepcopy(runs[0])
    for r in runs:
        p._p.remove(r)
    r = make_run(tmpl, text)
    fill_style_font(r, p)
    p._p.append(r)


def rows_of(t):
    out = []
    for row in t.rows:
        seen, lst = [], []
        for c in row.cells:
            if id(c._tc) in [id(x._tc) for x in seen]:
                continue
            seen.append(c)
            lst.append(c)
        out.append(lst)
    return out


def cell_rpr(cell):
    for p in cell.paragraphs:
        for r in p._p.findall(qn("w:r")):
            rp = r.find(qn("w:rPr"))
            if rp is not None and rp.find(qn("w:rFonts")) is not None and rp.find(qn("w:sz")) is not None:
                return copy.deepcopy(rp)
    return None


def set_cell(cell, text, donor=None):
    ps = cell.paragraphs
    p = ps[0]
    for extra in ps[1:]:
        extra._p.getparent().remove(extra._p)
    runs = p._p.findall(qn("w:r"))
    if runs:
        tmpl_rpr = copy.deepcopy(runs[0].find(qn("w:rPr")))
        for r in runs:
            p._p.remove(r)
    else:
        tmpl_rpr = copy.deepcopy(donor) if donor is not None else None
    r = OxmlElement("w:r")
    r.append(tmpl_rpr if tmpl_rpr is not None else OxmlElement("w:rPr"))
    set_color(r)
    kill_hl(r)
    fill_style_font(r, p)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p._p.append(r)


def put_lines(cell, lines):
    p = cell.paragraphs[0]
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    runs = p._p.findall(qn("w:r"))
    tmpl = copy.deepcopy(runs[0].find(qn("w:rPr")))
    for r in runs:
        p._p.remove(r)
    for i, line in enumerate(lines):
        if i:
            br = OxmlElement("w:r")
            br.append(copy.deepcopy(tmpl))
            br.append(OxmlElement("w:br"))
            p._p.append(br)
        r = OxmlElement("w:r")
        r.append(copy.deepcopy(tmpl))
        set_color(r)
        kill_hl(r)
        tx = OxmlElement("w:t")
        tx.set(qn("xml:space"), "preserve")
        tx.text = line
        r.append(tx)
        p._p.append(r)


def put_rich_lines(cell, lines):
    """向单元格写多行富文本；每行可为字符串或 (文本, 下划线, 加粗) 片段列表。"""
    p = cell.paragraphs[0]
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    runs = p._p.findall(qn("w:r"))
    assert runs, "空单元格无法取得字体模板: %r" % cell.text
    template_run = copy.deepcopy(runs[0])
    tmpl_rpr = copy.deepcopy(runs[0].find(qn("w:rPr")))
    for r in runs:
        p._p.remove(r)
    for i, line in enumerate(lines):
        if i:
            br = OxmlElement("w:r")
            br.append(copy.deepcopy(tmpl_rpr))
            br.append(OxmlElement("w:br"))
            p._p.append(br)
        segments = [(line, None, None)] if isinstance(line, str) else line
        for text, u, b in segments:
            p._p.append(make_run(template_run, text, u, b))


def normalize(doc):
    roots = [doc.element.body]
    for s in doc.sections:
        roots.append(s.header._element)
        roots.append(s.footer._element)
    for root in roots:
        for r in root.iter(qn("w:r")):
            set_color(r)
            kill_hl(r)


BLANK_DATE_RE = re.compile(
    r"^[\s\u3000_]*(?:(?:19|20)\d{2})?[\s\u3000_]*年[\s\u3000_]*月[\s\u3000_]*日[\s\u3000_]*$"
)


def fill_date_runs(doc):
    """海珠区：只替换整段为标准空白日期的 run，避免误改正文中的具体日期。"""
    if not DOC_DATE.strip():
        return
    for r in doc.element.body.iter(qn("w:r")):
        texts = r.findall(qn("w:t"))
        value = "".join(t.text or "" for t in texts)
        if value.strip() and BLANK_DATE_RE.fullmatch(value):
            texts[0].text = DOC_DATE.strip()
            for extra in texts[1:]:
                extra.text = ""


def fill_date_fields(doc):
    """海珠区：补齐整段日期及“日期／提交日期”栏，不改正文中的具体日期叙述。"""
    if not DOC_DATE.strip():
        return
    date = DOC_DATE.strip()
    for node in doc.element.body.iter(qn("w:p")):
        p = Paragraph(node, doc)
        text = p.text.strip()
        if not text or len(text) > 100 or not p.runs:
            continue
        if BLANK_DATE_RE.fullmatch(text):
            set_para_text(p, date)
        elif re.fullmatch(r"日\s*期\s*[:：].*年.*月.*日\s*", text):
            set_para_text(p, "    日      期：" + date)
        elif re.fullmatch(r"提交日期\s*[:：].*接收法院盖章\s*[:：]?\s*", text):
            set_para_text(p, "提交日期：" + date + "                   接收法院盖章：")


def money_text(value):
    if value is None or not str(value).strip():
        return ""
    raw = str(value).replace(",", "").strip()
    try:
        return "{:,.2f}".format(Decimal(raw))
    except InvalidOperation:
        return str(value).strip()


def load(name):
    with open(os.path.join(TPL, name), "rb") as f:
        return Document(io.BytesIO(f.read()))


def load_haizhu_account():
    candidates = [
        os.path.join(TPL, "海珠区材料", "退费申请及收款帐户承诺书(2).docx"),
        r"C:\Users\liyeye\.codex\skills\execution-application-documents\assets\海珠区\退费申请及收款承诺书_空白官式.docx",
    ]
    for path in candidates:
        if os.path.isfile(path):
            with open(path, "rb") as f:
                return Document(io.BytesIO(f.read()))
    raise FileNotFoundError("找不到海珠区退费申请及收款账户承诺书模板")


def save(doc, name):
    target = os.path.join(OUT, name)
    tmp = target + ".tmp"
    fill_date_fields(doc)
    fill_date_runs(doc)
    doc.save(tmp)
    os.replace(tmp, target)
    print("written:", name)


def find_para(doc, pred, what=""):
    for p in doc.paragraphs:
        if pred(p.text):
            return p
    raise AssertionError("找不到段落: " + what)


def set_spacing(doc, line=500):
    for p in doc.paragraphs:
        ppr = p._p.get_or_add_pPr()
        sp = ppr.find(qn("w:spacing"))
        if sp is None:
            sp = OxmlElement("w:spacing")
            ppr.append(sp)
        sp.set(qn("w:line"), str(line))
        sp.set(qn("w:lineRule"), "exact")


def clear_first_line_indent(p):
    ppr = p._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        ppr.remove(ind)


def set_first_line_chars(p, chars):
    ppr = p._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:firstLine"), str(chars * 280))
    ind.set(qn("w:firstLineChars"), str(chars * 100))


def tail_from(doc, anchor):
    out, started = [], False
    for p in doc.paragraphs:
        if not started and p.text.strip().startswith(anchor):
            started = True
        if started:
            out.append(p)
    return out


def tick_first_box(cell):
    for p in cell.paragraphs:
        for r in p._p.findall(qn("w:r")):
            for t in r.findall(qn("w:t")):
                if t.text and "□" in t.text:
                    t.text = t.text.replace("□", "☑", 1)
                    return True
    return False


def untick_all(cell):
    for p in cell.paragraphs:
        for r in p._p.findall(qn("w:r")):
            for t in r.findall(qn("w:t")):
                if t.text and "☑" in t.text:
                    t.text = t.text.replace("☑", "□")


def defendant_names():
    return "、".join(b.replace("被执行人：", "").split("，")[0] for b in DEF_BLOCKS)


def split_def(blk):
    """从"被执行人：姓名，性别，生日，汉族，住址，公民身份号码XX。"里抠字段。"""
    body = blk.replace("被执行人：", "")
    parts = [s.strip("。 ") for s in body.split("，") if s.strip("。 ")]
    name = parts[0]
    sex = parts[1] if len(parts) > 1 and parts[1] in ("男", "女") else ""
    birth = ""
    for s in parts:
        if "出生" in s:
            birth = s.replace("出生", "").replace("生", "")
    addr = ""
    for s in parts:
        if s.startswith("住所地"):
            addr = s[len("住所地"):]
        elif s.startswith("住"):
            addr = s[1:]
    idno = ""
    for s in parts:
        if "公民身份号码" in s:
            idno = s.replace("公民身份号码", "").strip("。 ")
        if "统一社会信用代码" in s:
            idno = s.replace("统一社会信用代码", "").strip("。 ")
    # 公司/个体户整段里“统一社会信用代码…；法定代表人：某某”连写时，会把法代一起带进
    # 10 表的证件号格（2026-10-19 合喜案踩过），这里在第一个分隔符处截断。
    for sep in ("；", ";", "法定代表人"):
        idno = idno.split(sep)[0].strip()
    return name, sex, birth, addr, idno


# ---------- 01 执行申请书 ----------
def build_01():
    doc = load("01执行申请书 .docx")
    ps = doc.paragraphs
    assert ps[0].text.strip() == "执行申请书", ps[0].text
    set_segments(ps[1], [(APPLICANT_BLOCK, None, None)])
    p_def = find_para(doc, lambda t: t.startswith("被执行人："), "被执行人")
    for extra in [p for p in doc.paragraphs if p.text.startswith("被执行人：")][1:]:
        extra._p.getparent().remove(extra._p)
    set_segments(p_def, [(DEF_BLOCKS[0], None, None)])
    from docx.text.paragraph import Paragraph
    cur = p_def
    for blk in DEF_BLOCKS[1:]:
        new = copy.deepcopy(p_def._p)
        cur._p.addnext(new)
        cur = Paragraph(new, p_def._parent)
        set_segments(cur, [(blk, None, None)])
    set_segments(find_para(doc, lambda t: t.startswith("案由"), "案由"), [("案由：" + CAUSE, None, None)])
    # 注意：这两处要写"标题的下一段"（正文），不能写标题段本身，否则标题被正文顶掉、
    # 模板样本里的旧正文（含旧案号、旧生效日）会留在文件里（2026-10-15 某案案踩过）。
    allp = doc.paragraphs
    i = next(i for i, p in enumerate(allp) if p.text.strip() == "申请执行依据")
    set_segments(allp[i + 1], [(COURT_FULL + CASE_NO + DOC_TYPE + "。", None, None)])
    allp = doc.paragraphs
    i = next(i for i, p in enumerate(allp) if p.text.strip() == "申请执行的理由及内容")
    set_segments(allp[i + 1], [(REASON1, None, None)])
    set_segments(find_para(doc, lambda t: t.strip().startswith("生效判决确定"), "生效判决确定段"), [(FINDINGS, None, None)])
    allp = doc.paragraphs
    i_first = next(i for i, p in enumerate(allp) if p.text.strip().startswith("1."))
    i_total = next(i for i, p in enumerate(allp) if p.text.strip().startswith("（以上各项暂合计"))
    for p in allp[i_first + len(ITEMS):i_total]:
        p._p.getparent().remove(p._p)
    allp = doc.paragraphs
    i_first = next(i for i, p in enumerate(allp) if p.text.strip().startswith("1."))
    for k, text in enumerate(ITEMS):
        set_segments(allp[i_first + k], [(text, None, None)])
    p_total = find_para(doc, lambda t: t.strip().startswith("（以上各项暂合计"), "暂合计")
    set_segments(p_total, [("（以上各项暂合计：" + TOTAL + "元）", None, True)])
    clear_first_line_indent(p_total)
    # “此致”与法院名必须分两段：写在同一个 w:t 里的换行符 WPS 会渲成同一行
    # （2026-10-19 用户改稿：此致缩进两字，法院名另起一段顶格）。
    p_cizhi = find_para(doc, lambda t: t.strip().startswith("此致"), "此致")
    set_segments(p_cizhi, [("此致", None, None)])
    prev = p_cizhi._p.getprevious()
    while prev is not None:
        p_prev = Paragraph(prev, p_cizhi._parent)
        if p_prev.text.strip():
            break
        before = prev.getprevious()
        prev.getparent().remove(prev)
        prev = before
    court_el = copy.deepcopy(p_cizhi._p)
    p_cizhi._p.addnext(court_el)
    from docx.text.paragraph import Paragraph
    p_court = Paragraph(court_el, p_cizhi._parent)
    set_segments(p_court, [(COURT_SHORT, None, None)])
    p_court.paragraph_format.page_break_before = False
    _ppr = p_court._p.find(qn("w:pPr"))
    if _ppr is not None:                     # 法院名要顶格：连 w:firstLineChars 一起清掉
        _ind = _ppr.find(qn("w:ind"))
        if _ind is not None:
            _ppr.remove(_ind)
    p_cizhi.paragraph_format.page_break_before = True
    p_sign = find_para(doc, lambda t: t.strip() == "申请执行人：", "落款申请执行人")
    set_segments(p_sign, [("申请执行人：", None, None)])
    set_first_line_chars(p_sign, 9)
    set_segments(find_para(doc, lambda t: "月" in t and "日" in t and len(t.strip()) < 20, "落款日期"),
                 [("年\u3000\u3000月\u3000\u3000日", None, None)])
    set_spacing(doc, 500)
    p_total.paragraph_format.keep_together = True
    p_total.paragraph_format.keep_with_next = False
    # 落款已由“此致”段前分页固定到末页；暂合计不能再 keep_with_next，
    # 否则 Word/WPS 会把它一起带到新页，并在段前分页后再多出一张空白页。
    tail = tail_from(doc, "此致")
    for i, p in enumerate(tail):
        p.paragraph_format.keep_together = True
        if i < len(tail) - 1:
            p.paragraph_format.keep_with_next = True
    normalize(doc)
    save(doc, "01_执行申请书_%s.docx" % APPLICANT)


# ---------- 01 要素式（默认不出；用户要时把 WITH_ELEMENT 置 True） ----------
def build_01e():
    doc = load("强制执行申请书（要素式·样本）.docx")
    t = doc.tables[0]
    rc = rows_of(t)

    def fill(cell, idx, text):
        while len(cell.paragraphs) <= idx:
            cell.add_paragraph()
        set_para_text(cell.paragraphs[idx], text)

    blanks_a = ["名称：", "住所地（主要办事机构所在地）：", "注册地/登记地： ",
                "法定代表人/负责人：\u2003职务：\u2003\u2003\u2003\u2003联系电话： ",
                "统一社会信用代码：", "银行账号： ", "开户名：\u2003\u2003\u2003\u2003开户行： "]
    blanks_b = ["名称：", "住所地（主要办事机构所在地）：", "注册地/登记地： ",
                "法定代表人/负责人：\u2003职务：\u2003\u2003\u2003\u2003联系电话： ", "统一社会信用代码："]
    blanks_nat = ["姓名：", "性别：男□ 女□ ",
                  "出生日期：\u2003\u2003年\u2003\u2003 月\u2003\u2003 日\u2003\u2003\u2003\u2003民族： ",
                  "工作单位：\u2003\u2003\u2003\u2003\u2003\u2003职务：\u2003\u2003\u2003\u2003\u2003联系电话： ",
                  "住所地（户籍所在地）：", "经常居住地： ", "证件类型：", "证件号码："]
    if APPLICANT_IS_COMPANY:
        for i, txt in enumerate(blanks_a):
            fill(rc[2][1], i, txt)
        c = rc[3][1]
        fill(c, 0, "名称：" + APPLICANT)
        fill(c, 1, "住所地（主要办事机构所在地）：" + APPLICANT_ADDR)
        fill(c, 2, "注册地/登记地： ")
        fill(c, 3, "法定代表人/负责人：" + APPLICANT_REP + "\u2003职务：" + APPLICANT_REP_TITLE
                   + "\u2003\u2003\u2003\u2003联系电话： ")
        fill(c, 4, "统一社会信用代码：" + APPLICANT_CODE)
        fill(c, 5, "银行账号： ")
        fill(c, 6, "开户名：\u2003\u2003\u2003\u2003开户行： ")
    else:
        name, sex, birth, addr, idno = split_def(APPLICANT_BLOCK.replace("申请执行人：", "被执行人："))
        for i, txt in enumerate(blanks_nat):
            fill(rc[2][1], i, txt)
        fill(rc[2][1], 0, "姓名：" + name)
        fill(rc[2][1], 1, "性别：男%s 女%s " % ("" if sex == "男" else "□", "" if sex == "女" else "□"))
        fill(rc[2][1], 2, "出生日期：" + birth + "\u2003\u2003\u2003\u2003民族：汉族")
        fill(rc[2][1], 4, "住所地（户籍所在地）：" + addr)
        fill(rc[2][1], 6, "证件类型：居民身份证")
        fill(rc[2][1], 7, "证件号码：" + idno)
        fill(rc[2][1], 8, "银行账号： ")
        fill(rc[2][1], 9, "开户名：\u2003\u2003\u2003\u2003\u2003\u2003\u2003\u2003\u2003开户行： ")
        for i, txt in enumerate(blanks_a):
            fill(rc[3][1], i, txt)
    c = rc[4][1]
    fill(c, 0, "有")
    fill(c, 1, "姓名：" + LAWYER)
    fill(c, 2, "单位：" + FIRM + "    职务：律师    联系电话：" + MOBILE)
    fill(c, 3, "代理权限：一般授权□    特别授权")
    fill(c, 4, "无□")
    base_row = None
    for ri, row in enumerate(rc):
        if row[0].text.replace("\n", "").replace(" ", "").startswith("被执行人（自然人"):
            base_row = ri
            break
    assert base_row is not None, "找不到被执行人（自然人）行"
    n_nat = max(1, sum(1 for x in DEFS_ARE_COMPANIES if not x))
    tr = t.rows[base_row]._tr
    last = tr
    for _ in range(n_nat - 1):
        new = copy.deepcopy(tr)
        last.addnext(new)
        last = new
    rc = rows_of(t)
    k = 0
    for i, is_comp in enumerate(DEFS_ARE_COMPANIES):
        if is_comp:
            continue
        name, sex, birth, addr, idno = split_def(DEF_BLOCKS[i])
        c = rc[base_row + k][1]
        fill(c, 0, "姓名：" + name)
        fill(c, 1, "性别：男%s 女%s " % ("" if sex == "男" else "□", "" if sex == "女" else "□"))
        fill(c, 2, "出生日期：" + birth + "\u2003\u2003\u2003\u2003民族：汉族")
        fill(c, 3, "工作单位：\u2003\u2003\u2003\u2003\u2003\u2003职务：\u2003\u2003\u2003\u2003\u2003 联系电话： ")
        fill(c, 4, "住所地（户籍所在地）：" + addr)
        fill(c, 5, "经常居住地：  ")
        fill(c, 6, "证件类型：居民身份证")
        fill(c, 7, "证件号码：" + idno)
        k += 1
    legal_row = base_row + n_nat
    for i, txt in enumerate(blanks_b):
        fill(rc[legal_row][1], i, txt)
    fill(rc[legal_row][1], 10, "所有制性质：国有□（控股□ 参股□） 民营□ 其他            ")
    base = legal_row + 2
    fill(rc[base][1], 1, "判决书 裁定书□ 调解书□ 制裁决定□ 支付令□ ")
    fill(rc[base + 1][1], 0, COURT_FULL)
    fill(rc[base + 2][1], 0, CAUSE)
    fill(rc[base + 3][1], 0, CASE_NO)
    fill(rc[base + 4][1], 0, ELEMENT_DATE)
    fill(rc[base + 5][1], 0, ELEMENT_PANJUE)
    fill(rc[base + 7][0], 0, ELEMENT_NARRATIVE)
    c = rc[base + 8][1]
    lines = ELEMENT_MONEY if isinstance(ELEMENT_MONEY, (list, tuple)) else ELEMENT_MONEY.split("||")
    while len(c.paragraphs) < len(lines):
        c.add_paragraph()
    for i, txt in enumerate(lines):
        fill(c, i, txt)
    c = rc[base + 9][1]
    fill(c, 0, "有" if HAS_BAOQUAN else "有□")
    fill(c, 1, "保全案号： ")
    fill(c, 2, "保全措施最早到期时间：\u2003\u2003年\u2003\u2003月\u2003\u2003日 ")
    fill(c, 3, "无□" if HAS_BAOQUAN else "无")
    fill(rc[base + 10][1], 0, ELEMENT_PROPERTY)
    for row in t.rows:
        trPr = row._tr.find(qn("w:trPr"))
        if trPr is None:
            continue
        h = trPr.find(qn("w:trHeight"))
        if h is not None and h.get(qn("w:hRule")) == "exact":
            h.set(qn("w:hRule"), "atLeast")
    normalize(doc)
    save(doc, "01_执行申请书（要素式）_%s.docx" % APPLICANT)


# ---------- 02 授权委托书 ----------
def build_02():
    doc = load("02执行授权委托书.docx")
    ps = doc.paragraphs
    set_segments(ps[1], [("委托人：" + APPLICANT_BLOCK.replace("申请执行人：", ""), None, None)])
    set_segments(ps[6], [("受委托人：", None, True)])
    work_ps = [p for p in doc.paragraphs if p.text.strip().startswith("工作单位：")]
    assert len(work_ps) >= 2, "授权委托书缺少第二位受委托人的工作单位栏"
    set_segments(work_ps[1], [("工作单位：示例律师事务所    职务：", None, None)])
    set_segments(ps[11], [
        ("现委托以上受委托人在我方与", None, None),
        (defendant_names(), True, None),
        ("之间", None, None),
        (CAUSE, True, None),
        ("申请强制执行一案中，作为我方的执行代理人。", None, None),
    ])
    normalize(doc)
    save(doc, "02_执行授权委托书_%s.docx" % APPLICANT)


# ---------- 04 法定代表人／负责人身份证明书 ----------
def build_04():
    if APPLICANT_IS_INDIVIDUAL_BUSINESS:
        doc = load("05负责人身份证明书(个体户用这个）.docx")
        p = find_para(
            doc,
            lambda t: "在我店担任" in t or "在我厂担任" in t,
            "负责人姓名行",
        )
        name_hit = title_hit = seen_title_marker = False
        for r in p.runs:
            if "在我店担任" in r.text or "在我厂担任" in r.text:
                seen_title_marker = True
            elif seen_title_marker and not title_hit and r.text.strip():
                r.text = "经营者"
                title_hit = True
            elif not seen_title_marker and not name_hit and r.text.strip():
                r.text = APPLICANT_REP
                name_hit = True
        assert name_hit and title_hit, "未找到负责人身份证明书样例字段"

        nonempty = [q for q in doc.paragraphs if q.text.strip()]
        date_i = next(
            i for i, q in enumerate(nonempty)
            if q.text.strip().startswith("日期") and "年" in q.text and "月" in q.text
        )
        comp = nonempty[date_i - 1]
        for r in reversed(comp.runs):
            if r.text.strip():
                leading = r.text[: len(r.text) - len(r.text.lstrip(" "))]
                r.text = leading + APPLICANT
                break
        normalize(doc)
        save(doc, "04_负责人身份证明书_%s.docx" % APPLICANT)
        return

    doc = load("04法定代表人身份证明书(公司用这个).docx")
    p = find_para(doc, lambda t: "在我公司担任" in t, "姓名行")
    hit = False
    for r in p.runs:
        if r.text.strip() and "我公司" not in r.text and "职务" not in r.text and not hit:
            r.text = APPLICANT_REP
            hit = True
    assert hit, "未找到姓名 run"
    comp = find_para(doc, lambda t: "年" in t and "月" in t and "日" in t, "落款行")
    for r in comp.runs:
        if "公司" in r.text:
            tail = len(r.text) - len(r.text.lstrip(" "))
            r.text = " " * max(0, tail - 1) + APPLICANT
            break
    normalize(doc)
    save(doc, "04_法定代表人身份证明书_%s.docx" % APPLICANT)


# ---------- 05 账户书（默认收款书；越秀退费书；海珠区专用退费承诺书） ----------
def build_05():
    if ACCOUNT == "海珠退费承诺书":
        if "海珠" not in COURT_FULL:
            raise ValueError("海珠专用退费承诺书只能用于广州市海珠区人民法院")
        if not str(REFUND_TOTAL).strip():
            raise ValueError("海珠区退费承诺书必须设置 REFUND_TOTAL（可退诉讼费合计）")
        doc = load_haizhu_account()
        set_segments(doc.paragraphs[2], [(COURT_FULL + "：", None, None)])
        total = money_text(REFUND_TOTAL)
        body = ("    " + CASE_NO + "案审理终结后，原告同意将应退回给原告的受理费等诉讼费用"
                "合计人民币" + total + "元通过以下账户予以接收。")
        if REFUND_BREAKDOWN.strip():
            body += REFUND_BREAKDOWN.strip().rstrip("；;。") + "；以上合计人民币" + total + "元。"
        body += "具体退费范围以生效裁判及法院核定为准。"
        set_segments(doc.paragraphs[3], [(body, None, None)])

        cell = doc.tables[0].rows[1].cells[0]
        ps = cell.paragraphs

        def fill_underline(p, value):
            if not str(value).strip():
                return
            for r in p.runs:
                if "_" in r.text:
                    width = max(len(r.text), len(str(value)) + 6)
                    r.text = "   " + str(value).strip() + " " * max(0, width - len(str(value).strip()) - 3)
                    return
            raise AssertionError("找不到下划线填空: " + p.text)

        fill_underline(ps[0], APPLICANT)
        fill_underline(ps[1], APPLICANT_ACCOUNT)
        if APPLICANT_BANK.strip():
            # 开户行通常较长，去掉模板末尾的“(具体到支行名称)”提示，避免折成两行。
            set_segments(ps[2], [("开  户  行：", None, None),
                                 ("  " + APPLICANT_BANK.strip(), True, None)])
        else:
            fill_underline(ps[2], APPLICANT_BANK)
        fill_underline(ps[3], APPLICANT_MOBILE)
        if APPLICANT_CODE:
            # 公司填统一社会信用代码；只替换方框 run，避免把法定代表人姓名带进证件栏。
            for r in ps[4].runs:
                if "□" in r.text:
                    r.text = "   " + APPLICANT_CODE + "   "
                    break
        normalize(doc)
        save(doc, "05_退费申请及收款账户承诺书_%s.docx" % APPLICANT)
        return

    if ACCOUNT == "越秀退费书":
        with open(os.path.join(TPL, "广州法院材料材料", "越秀区退费表.docx"), "rb") as f:
            doc = Document(io.BytesIO(f.read()))
        set_segments(doc.paragraphs[0], [(COURT_FULL, None, None)])
        t = doc.tables[0]
        rc = rows_of(t)
        donor = cell_rpr(rc[0][0])
        set_cell(rc[0][1], APPLICANT, donor)
        set_cell(rc[0][3], CASE_NO, donor)
        set_cell(rc[1][1], APPLICANT, donor)
        set_cell(rc[1][3], "", donor)
        set_cell(rc[2][1], "", donor)
        set_cell(rc[4][1], "", donor)
        normalize(doc)
        save(doc, "05_诉讼费退费账户确认书_%s.docx" % APPLICANT)
        return
    doc = load("06收款帐户确认书（最好加上银行卡复印件）.docx")
    ps = doc.paragraphs
    set_segments(ps[0], [(COURT_FULL, None, None)])
    set_segments(ps[3], [(COURT_FULL + "：", None, None)])
    set_segments(ps[4], [("为保证我的执行案件将来准确无误收取到贵院执行到位的执行款，特向贵院确认以下银行账户信息，作为我收取",
                          None, None)])
    set_segments(ps[5], [(CASE_NO, True, None),
                         ("案件执行款的唯一账户。除我另行书面变更外，如因银行账户信息错误导致的汇款风险一律由本人承担。", False, None)])
    set_segments(ps[6], [("开户银行（网点）：", None, None), (" " * 29, True, None)])
    set_sz(ps[6].runs[1], 16)
    set_segments(ps[7], [("账           号：", None, None), (" " * 30, True, None)])
    set_sz(ps[7].runs[1], 16)
    # 14 磅下每侧空格 = 17 - 名称字数，使户名下划线右缘与开户行、账号齐平。
    pad = max(1, 17 - len(APPLICANT))
    set_segments(ps[8], [("户           名：", None, None),
                         (" " * pad + APPLICANT + " " * pad, True, None)])
    set_sz(ps[8].runs[1], 14)
    normalize(doc)
    save(doc, "05_收款帐户确认书_%s.docx" % APPLICANT)


# ---------- 07 法院提交材料清单 ----------
def build_07():
    doc = load("07法院提交材料清单.docx")
    ps = doc.paragraphs
    set_segments(ps[0], [(COURT_SHORT, None, None)])
    set_segments(ps[3], [("当事人：" + APPLICANT + "  " + "联系电话：" + MOBILE, None, None)])
    t = doc.tables[0]
    rc = rows_of(t)
    set_cell(rc[2][1], COURT_FULL + CASE_NO + DOC_TYPE, cell_rpr(rc[2][1]))
    for row in list(t.rows):
        if not RISK_NOTICE and row.cells[1].text.strip() == "风险提示书":
            row._tr.getparent().remove(row._tr)
    normalize(doc)
    save(doc, "07_法院提交材料清单_%s.docx" % APPLICANT)


# ---------- 08 立案审批表 ----------
def build_08():
    doc = load("08立案审批表(给琼姐盖章的时候写，上传企业微信附件).docx")
    t = doc.tables[0]
    cells = rows_of(t)
    fb = cell_rpr(cells[3][0])

    def put(c, s):
        set_cell(c, s, cell_rpr(c) or fb)

    put(cells[3][1], APPLICANT)
    put(cells[4][1], defendant_names())
    put(cells[5][1], CAUSE)
    put(cells[5][3], BIAODI)
    put(cells[6][1], "□固定 □固定+提成   □免费 ")
    put(cells[7][3], "主办：" + LAWYER)
    put(cells[8][3], "协办1 " + ASSIST)
    put(cells[10][1], APPLICANT + "诉" + defendant_names() + CAUSE + "案。" + COURT_FULL + CASE_NO
        + DOC_TYPE + "于" + EFFECTIVE_DATE + "生效，拟申请强制执行。")
    normalize(doc)
    save(doc, "08_立案审批表_%s.docx" % APPLICANT)


# ---------- 09 所函 ----------
def build_09():
    doc = load("09示例律师事务所所函.docx")
    ps = doc.paragraphs
    set_segments(ps[2], [(COURT_SHORT + "：", None, None)])
    set_segments(ps[3], [
        ("本所接受" + APPLICANT + "委托，现派", None, None),
        (LAWYER + "律师、" + " " * 6, True, None),
        ("担任其与", None, None),
        (defendant_names(), True, None),
        ("之间", None, None),
        (CAUSE, True, None),
        ("申请强制执行一案的诉讼代理人。", None, None),
    ])
    set_segments(ps[8], [("年\u3000\u3000月\u3000\u3000日", None, None)])
    normalize(doc)
    save(doc, "09_示例律师事务所所函_%s.docx" % APPLICANT)


# ---------- 10 被告（第三人）信息表 ----------
def build_10():
    doc = load("10被告（第三人）信息表.docx")
    set_segments(doc.paragraphs[0], [(COURT_SHORT, None, True)])
    t = doc.tables[0]
    cells = rows_of(t)
    fb = cell_rpr(cells[1][1])

    def put(c, s):
        donor = cell_rpr(c)
        set_cell(c, s, donor if donor is not None else fb)

    put(cells[1][1], CASE_NO)
    put(cells[1][3], APPLICANT)
    blocks = [(2, 3, 4, 5), (9, 10, 11, 12), (16, 17, 18, 19), (23, 24, 25, 26)]
    for blk, (bn, b1, b2, blive) in zip(DEF_BLOCKS, blocks):
        name, sex, birth, addr, idno = split_def(blk)
        put(cells[bn][2], name)
        put(cells[b1][2], idno)
        put(cells[b2][2], idno)
        tick_first_box(cells[bn][0])
        if "统一社会信用代码" in blk:
            # 单位被执行人没有户籍所在地，改填登记／经常联系地址。
            untick_all(cells[blive][2])
            put(cells[blive][2], "☑经常居住地址：" + addr)
    for i in range(len(DEF_BLOCKS), len(blocks)):
        bn, b1, b2, blive = blocks[i]
        untick_all(cells[bn][0])
        untick_all(cells[blive][2])
    normalize(doc)
    save(doc, "10_被告（第三人）信息表_%s.docx" % APPLICANT)


# ---------- 11 诉讼文书送达地址确认书 ----------
def build_11():
    doc = load("11诉讼文书送达地址确认书.docx")
    set_segments(doc.paragraphs[0], [(COURT_SHORT, None, True)])
    t = doc.tables[0]
    cells = rows_of(t)
    fb = cell_rpr(cells[2][2])

    def put(c, s):
        donor = cell_rpr(c)
        set_cell(c, s, donor if donor is not None else fb)

    addr = APPLICANT_ADDR
    if APPLICANT_IS_COMPANY:
        for s in APPLICANT_BLOCK.split("，"):
            if s.startswith("住所地") or s.startswith("申请执行人：") and "住所地" in s:
                addr = s.split("住所地")[-1].lstrip("：: ")
    put(cells[0][3], CAUSE)
    put(cells[2][2], APPLICANT)
    put(cells[3][2], "☑同意手机短信送达，手机号码：         □同意微信送达，微信号：              ")
    # 电子送达手机号和身份号码必须填写并加下划线；企业身份号码栏填统一社会信用代码。
    applicant_econtact = MOBILE if APPLICANT_IS_COMPANY else APPLICANT_MOBILE
    applicant_identity = APPLICANT_CODE if APPLICANT_IS_COMPANY else APPLICANT_ID
    if not applicant_econtact.strip():
        raise ValueError("11表当事人电子送达手机号为空，请补充 APPLICANT_MOBILE/经办人手机号")
    if not applicant_identity.strip():
        raise ValueError("11表当事人身份号码为空：企业请补充 APPLICANT_CODE，自然人请补充 APPLICANT_ID")
    put_rich_lines(cells[4][2], [
        "☑同意采用人民法院在线服务网（可通过微信小程序“人民法院在线服务”登录）等诉讼服务平台电子送达",
        [("手机号码：", None, None), ("  " + applicant_econtact + "  ", True, None),
         ("公民身份号码：", None, None), ("  " + applicant_identity + "  ", True, None)],
    ])
    put(cells[5][3], addr)
    put(cells[6][3], (" " + APPLICANT_MOBILE + " ") if APPLICANT_MOBILE else "")
    put(cells[7][2], LAWYER)
    put_rich_lines(cells[8][2], [[
        ("☑同意手机短信送达，手机号码：", None, None),
        ("  " + MOBILE + "  ", True, None),
        ("□同意微信送达，微信号：              ", None, None),
    ]])
    put_rich_lines(cells[9][2], [
        "☑同意采用人民法院在线服务网（可通过微信小程序“人民法院在线服务”登录）等诉讼服务平台电子送达",
        [("手机号码：", None, None), ("  " + MOBILE + "  ", True, None),
         ("公民身份号码：", None, None), ("  " + LAWYER_ID + "  ", True, None)],
    ])
    put(cells[10][3], FIRM_ADDR)
    put(cells[11][3], " " + MOBILE + "  ")
    put_lines(cells[12][1], [
        "同意以上述邮寄送达地址作为接收诉讼文书以及诉讼活动信息的送达地址；电子送达须另行勾选确认。",
        "签名或捺印：_______________",
        "_____年____月____日",
    ])
    normalize(doc)
    save(doc, "11_诉讼文书送达地址确认书_%s.docx" % APPLICANT)


# ---------- 12 南海执行风险提示（仅南海） ----------
def build_12():
    doc = load("12执行风险提示（南海才需要）.docx")
    normalize(doc)
    save(doc, "12_南海区执行风险提示_%s.docx" % APPLICANT)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    target = os.path.join(OUT, os.path.basename(JUDGMENT))
    if not os.path.exists(target):
        shutil.copy2(JUDGMENT, target)
        print("copied judgment")
    todo = [build_01, build_02, build_05, build_07, build_08, build_09, build_10, build_11]
    if APPLICANT_IS_COMPANY:
        todo.insert(2, build_04)
    if RISK_NOTICE:
        todo.append(build_12)
    if WITH_ELEMENT:
        todo.insert(1, build_01e)
    for fn in todo:
        fn()
