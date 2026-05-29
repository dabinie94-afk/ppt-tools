# -*- coding: utf-8 -*-
"""
병리검사실 청소계획안 - 씨젠의료재단 CI 고도화 버전
기존 (4) 버전 구조 유지, 디자인 전면 개선
슬라이드: 1.제목+결재 / 2.목적+청소주기+비용비교 / 3.구역담당자
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ── CI 컬러 ─────────────────────────────────────
RED       = RGBColor(0xCB, 0x09, 0x13)
RED_PALE  = RGBColor(0xFD, 0xEC, 0xED)
RED_DARK  = RGBColor(0xA0, 0x07, 0x10)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
CHARCOAL  = RGBColor(0x24, 0x19, 0x16)   # 최소 사용
GRAY_BG   = RGBColor(0xF7, 0xF7, 0xF7)
GRAY_LINE = RGBColor(0xE2, 0xE2, 0xE2)
TEXT_DARK = RGBColor(0x1A, 0x1A, 0x1A)
TEXT_GRAY = RGBColor(0x66, 0x66, 0x66)

W = Inches(13.33)
H = Inches(7.5)
LOGO = r"C:\Users\최다빈\Desktop\_ci07.png"

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank = prs.slide_layouts[6]


# ════════════════════════════════════════
# 헬퍼
# ════════════════════════════════════════
def rh(rgb): return '{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])

def box(slide, l, t, w, h, fill=None, lc=None, lw=Pt(0.75)):
    s = slide.shapes.add_shape(1, l, t, w, h)
    if fill: s.fill.solid(); s.fill.fore_color.rgb = fill
    else: s.fill.background()
    if lc: s.line.width = lw; s.line.color.rgb = lc
    else: s.line.fill.background()
    return s

def txt(slide, text, l, t, w, h,
        size=Pt(11), bold=False, color=TEXT_DARK,
        align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = "맑은 고딕"
    r.font.size = size; r.font.bold = bold
    r.font.italic = italic; r.font.color.rgb = color
    return tb

def logo(slide, l, t, h=Inches(0.42)):
    from PIL import Image
    img = Image.open(LOGO)
    w = int(h * img.width / img.height)
    slide.shapes.add_picture(LOGO, l, t, width=w, height=h)

def cbg(cell, rgb):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    for o in tcPr.findall(qn('a:solidFill')): tcPr.remove(o)
    sf = etree.SubElement(tcPr, qn('a:solidFill'))
    sr = etree.SubElement(sf, qn('a:srgbClr')); sr.set('val', rh(rgb))

def cborder(cell, color=GRAY_LINE, w=Pt(0.5)):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    for side in ('a:lnL','a:lnR','a:lnT','a:lnB'):
        ln = etree.SubElement(tcPr, qn(side)); ln.set('w', str(int(w)))
        sf = etree.SubElement(ln, qn('a:solidFill'))
        sr = etree.SubElement(sf, qn('a:srgbClr')); sr.set('val', rh(color))

def sc(cell, text, size=Pt(10), bold=False, color=TEXT_DARK,
       align=PP_ALIGN.CENTER, bg=None, bc=GRAY_LINE, wrap=True, italic=False):
    cell.text = ""
    tf = cell.text_frame; tf.word_wrap = wrap
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.name = "맑은 고딕"; r.font.size = size
    r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    if bg: cbg(cell, bg)
    cborder(cell, bc)

def frame(slide, title, page):
    """공통 레이아웃: 레드 상단 헤더 + 레드 하단 바"""
    box(slide, 0, 0, W, Inches(0.78), fill=RED)
    box(slide, 0, 0, Inches(0.07), Inches(0.78), fill=RED_DARK)
    txt(slide, title, Inches(0.3), Inches(0.14), Inches(9.5), Inches(0.52),
        size=Pt(20), bold=True, color=WHITE)
    txt(slide, page, Inches(12.0), Inches(0.22), Inches(1.2), Inches(0.36),
        size=Pt(9.5), color=RGBColor(0xFF,0xBB,0xBB), align=PP_ALIGN.RIGHT)
    logo(slide, Inches(10.65), Inches(0.18), h=Inches(0.44))
    box(slide, 0, H - Inches(0.32), W, Inches(0.32), fill=RED)
    txt(slide, "경영관리본부  |  총무팀  |  대외비",
        Inches(0.35), H-Inches(0.3), Inches(6), Inches(0.26),
        size=Pt(8.5), color=WHITE)


# ════════════════════════════════════════
# SLIDE 1 — 제목 + 결재표
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)
box(slide, 0, 0, W, H, fill=WHITE)

# 상단 레드 영역
box(slide, 0, 0, W, Inches(3.2), fill=RED)
# 좌측 어두운 레드 세로 포인트
box(slide, 0, 0, Inches(0.12), Inches(3.2), fill=RED_DARK)

# 로고
logo(slide, Inches(10.4), Inches(0.28), h=Inches(0.56))

# 부서/날짜
txt(slide, "경영관리본부  |  총무팀",
    Inches(0.55), Inches(0.25), Inches(7), Inches(0.4),
    size=Pt(11), color=RGBColor(0xFF,0xBB,0xBB))

# 메인 제목
txt(slide, "본사 병리검사실",
    Inches(0.55), Inches(0.9), Inches(11), Inches(1.0),
    size=Pt(40), bold=True, color=WHITE)
txt(slide, "환경 위생 및 청소 계획(안)",
    Inches(0.55), Inches(1.82), Inches(11), Inches(0.9),
    size=Pt(32), bold=False, color=RGBColor(0xFF,0xDD,0xDE))

# 날짜 뱃지
box(slide, Inches(0.55), Inches(2.72), Inches(2.2), Inches(0.38),
    fill=RED_DARK)
txt(slide, "2026. 04. 14",
    Inches(0.55), Inches(2.72), Inches(2.2), Inches(0.38),
    size=Pt(11), bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# 흰색 영역 — 결재표 타이틀
box(slide, Inches(0.35), Inches(3.45), Inches(0.06), Inches(0.36), fill=RED)
txt(slide, "결  재",
    Inches(0.5), Inches(3.45), Inches(3), Inches(0.36),
    size=Pt(13), bold=True, color=RED)

# 결재표
sign_cols = ["담  당", "총무팀장", "경영관리본부", "경영관리부문", "기획조정실", "행정원장", "이사장"]
cw_sign = [Inches(1.5), Inches(1.5), Inches(1.7), Inches(1.7), Inches(1.5), Inches(1.5), Inches(1.5)]
tbl_sign = slide.shapes.add_table(
    2, 7, Inches(0.35), Inches(3.9),
    sum(cw_sign), Inches(1.8)
).table
for ci, cw in enumerate(cw_sign):
    tbl_sign.columns[ci].width = cw

# 헤더행
for ci, v in enumerate(sign_cols):
    sc(tbl_sign.cell(0, ci), v, size=Pt(10.5), bold=True, color=WHITE, bg=RED, bc=WHITE)
# 서명행
for ci in range(7):
    sc(tbl_sign.cell(1, ci), "", bg=WHITE, bc=GRAY_LINE)

# 하단 레드 라인
box(slide, 0, H-Inches(0.1), W, Inches(0.1), fill=RED)


# ════════════════════════════════════════
# SLIDE 2 — 목적 + 청소주기 + 비용비교
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)
box(slide, 0, 0, W, H, fill=WHITE)
frame(slide, "본사 병리검사실  환경 위생 및 청소 계획(안)", "02 / 03")

# ── [01] 목적 ─────────────────────────────────
box(slide, Inches(0.3), Inches(0.88), Inches(0.06), Inches(1.85), fill=RED)
txt(slide, "01  목적",
    Inches(0.5), Inches(0.88), Inches(4), Inches(0.36),
    size=Pt(13), bold=True, color=RED)

purposes = [
    ("신사옥 자산 가치 보존",
     "신사옥 이전 초기 집중 관리를 통해\n병리검사실 시설 노후화 방지 및 최상급 환경 품질 유지"),
    ("검사 신뢰도 및 위생 강화",
     "청소 주기 단축(격월→매월)을 통한\n맞춤형 위생 관리 확립"),
    ("근무 환경 및 생산성 제고",
     "쾌적한 위생 환경 조성을 통한\n임직원 보건 관리 증진 및 검사 업무 정확도 극대화"),
]
for j, (t, b) in enumerate(purposes):
    lft = Inches(0.5) + j * Inches(4.25)
    bw  = Inches(4.0)
    box(slide, lft, Inches(1.3), bw, Inches(1.35),
        fill=WHITE, lc=GRAY_LINE, lw=Pt(0.8))
    box(slide, lft, Inches(1.3), bw, Inches(0.38), fill=RED)
    txt(slide, t, lft+Inches(0.12), Inches(1.33), bw-Inches(0.2), Inches(0.32),
        size=Pt(10.5), bold=True, color=WHITE)
    txt(slide, b, lft+Inches(0.12), Inches(1.73), bw-Inches(0.2), Inches(0.88),
        size=Pt(9.5), color=TEXT_GRAY, wrap=True)

# ── [02] 청소 주기 ─────────────────────────────
box(slide, Inches(0.3), Inches(2.82), Inches(0.06), Inches(4.1), fill=RED)
txt(slide, "02  청소 주기 상세  (신사옥 기준)",
    Inches(0.5), Inches(2.82), Inches(7), Inches(0.36),
    size=Pt(13), bold=True, color=RED)

clean_rows = [
    ("실  명",                      "테이블/선반\n(현행→계획)",             "장  비\n(현행→계획)",               "바  닥\n(현행→계획)"),
    ("육안검사실 / 유해화학물질보관실", "해당없음/매주 → 해당없음/매일",       "매일/매주 → 매일/매일",               "2~3개월 → 매월"),
    ("포매실 / 조직침투기실",         "매일/매주 → 매일/매일",              "매일 → 매일",                         "2~3개월 → 매월"),
    ("조직표본제작실 / 검체접수실",    "매일 → 매일",                        "매일/해당없음 → 매일/해당없음",        "2~3개월 → 매월/격월"),
    ("염색실 / 면역병리실",           "매주/매일 → 매일/매일",              "매일 → 매일",                         "2~3개월 → 격월"),
    ("디지털병리실 / 의무정보실",      "매일 → 매일",                        "매일/해당없음 → 매일/해당없음",        "2~3개월/6개월 → 격월"),
    ("병리자료실 / 시약보관실",        "매주 → 매주",                        "해당없음 → 해당없음",                  "해당없음 → 분기"),
    ("검체보관실 / 세포표본제작실",    "매주/매일 → 매일/매일",              "해당없음/매일 → 해당없음/매일",        "해당없음/6개월 → 분기"),
]
cw_c = [Inches(2.8), Inches(2.4), Inches(2.4), Inches(2.1)]
tc = slide.shapes.add_table(
    len(clean_rows), 4,
    Inches(0.5), Inches(3.26), sum(cw_c), Inches(3.55)
).table
for ci, cw in enumerate(cw_c):
    tc.columns[ci].width = cw
for ri, row in enumerate(clean_rows):
    for ci, v in enumerate(row):
        c = tc.cell(ri, ci)
        if ri == 0:
            sc(c, v, size=Pt(9.5), bold=True, color=WHITE, bg=RED, bc=WHITE)
        else:
            bg = WHITE if ri % 2 == 1 else GRAY_BG
            is_change = '→' in v
            sc(c, v, size=Pt(8.5),
               bold=(ci == 0),
               color=RED if is_change else (TEXT_DARK if ci==0 else TEXT_GRAY),
               bg=RED_PALE if is_change else bg, bc=GRAY_LINE, wrap=True)

# ── [03] 장비/비품 비용 비교 ────────────────────
box(slide, Inches(10.05), Inches(0.88), Inches(0.06), Inches(2.7), fill=RED)
txt(slide, "03  장비·비품 관리 및 비용 비교",
    Inches(10.25), Inches(0.88), Inches(3.0), Inches(0.36),
    size=Pt(13), bold=True, color=RED)

cost_rows = [
    ("구  분",        "변경 전 (현행)",  "변경 후 (예정)",  "비  고"),
    ("청소 주기",     "2개월 1회",       "1개월 1회",       "격월 → 매월"),
    ("청소 비용(월)", "1,155,000원",     "1,155,000원",     "동일"),
    ("연간 비용",     "6,930,000원",     "13,860,000원",    "+6,930,000원"),
]
cw_cost = [Inches(1.4), Inches(1.1), Inches(1.1), Inches(1.1)]
tcost = slide.shapes.add_table(
    4, 4, Inches(10.2), Inches(1.3), sum(cw_cost), Inches(1.9)
).table
for ci, cw in enumerate(cw_cost):
    tcost.columns[ci].width = cw
for ri, row in enumerate(cost_rows):
    for ci, v in enumerate(row):
        c = tcost.cell(ri, ci)
        if ri == 0:
            sc(c, v, size=Pt(9), bold=True, color=WHITE, bg=RED, bc=WHITE)
        else:
            bg = WHITE if ri % 2 == 1 else GRAY_BG
            is_highlight = (ri == 3)
            sc(c, v, size=Pt(9),
               bold=is_highlight,
               color=RED if is_highlight else (TEXT_DARK if ci==0 else TEXT_GRAY),
               bg=RED_PALE if is_highlight else bg, bc=GRAY_LINE)

# 추가비용 강조 카드
box(slide, Inches(10.2), Inches(3.35), Inches(2.9), Inches(2.5),
    fill=RED, lc=None)
txt(slide, "연간 추가 소요 예산",
    Inches(10.2), Inches(3.5), Inches(2.9), Inches(0.42),
    size=Pt(10), color=WHITE, align=PP_ALIGN.CENTER)
txt(slide, "+6,930,000원",
    Inches(10.2), Inches(3.98), Inches(2.9), Inches(0.7),
    size=Pt(26), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
box(slide, Inches(10.5), Inches(4.72), Inches(2.3), Inches(0.02), fill=RGBColor(0xFF,0xAA,0xAA))
txt(slide, "청소 주기 변경 (격월→매월) 시",
    Inches(10.2), Inches(4.8), Inches(2.9), Inches(0.32),
    size=Pt(8.5), color=RGBColor(0xFF,0xDD,0xDE), align=PP_ALIGN.CENTER)
txt(slide, "연간 발생 추가 비용",
    Inches(10.2), Inches(5.12), Inches(2.9), Inches(0.32),
    size=Pt(8.5), color=RGBColor(0xFF,0xDD,0xDE), align=PP_ALIGN.CENTER)

# 중앙 세로 구분선
box(slide, Inches(9.9), Inches(0.88), Inches(0.02), Inches(5.9), fill=GRAY_LINE)


# ════════════════════════════════════════
# SLIDE 3 — 조직병리팀 구역 담당자
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)
box(slide, 0, 0, W, H, fill=WHITE)
frame(slide, "조직병리팀  구역 담당자", "03 / 03")

box(slide, Inches(0.3), Inches(0.9), Inches(0.06), Inches(0.38), fill=RED)
txt(slide, "각 구역별 담당자 현황  (방 번호 기준)",
    Inches(0.5), Inches(0.9), Inches(11), Inches(0.38),
    size=Pt(13), bold=True, color=RED)

# 두 테이블 공통 설정
left_data = [
    ("방 번호", "방  명",              "구역 담당자", "청소 담당자"),
    ("1527",   "유해화학물질 보관실",  "신정윤",     "박경태"),
    ("1528",   "침투기실",             "이창덕",     "오영권, 김지현"),
    ("1529",   "검체접수실",           "김영도",     "김문규, 조은빈"),
    ("1530",   "육안검사실",           "강시영",     "김선재, 김장미"),
    ("1531",   "포매실",               "이창덕",     "이성윤, 최선우"),
    ("1532",   "조직표본제작실",       "조기철",     "송하철, 이동현"),
    ("1533",   "면역병리실",           "오원호",     "한지원"),
    ("1534",   "염색실",               "류현규",     "황민영, 장두찬"),
]
right_data = [
    ("방 번호", "방  명",              "구역 담당자", "청소 담당자"),
    ("1535",   "디지털병리실",         "조기철",     "김영일"),
    ("1545",   "병리자료실",           "김영도",     "송영찬"),
    ("1546",   "시약보관실",           "이창덕",     "이동현"),
    ("1547",   "검체보관실",           "최수빈",     "김건희"),
    ("—",      "신관 방향 붙박이장",   "오원호",     "김성찬"),
    ("—",      "염색실 앞 붙박이장",   "조기철",     "이윤남"),
    ("—",      "여자 탈의실",          "강시영",     "김민경"),
    ("—",      "남자 탈의실",          "류현규",     "류현규"),
]

cw_d = [Inches(1.1), Inches(2.5), Inches(1.4), Inches(1.4)]
tw = sum(cw_d)

for table_idx, (data, left_pos) in enumerate([
    (left_data,  Inches(0.35)),
    (right_data, Inches(6.85)),
]):
    t = slide.shapes.add_table(
        9, 4, left_pos, Inches(1.42), tw, Inches(5.6)
    ).table
    for ci, cw in enumerate(cw_d):
        t.columns[ci].width = cw

    for ri, row in enumerate(data):
        for ci, v in enumerate(row):
            c = t.cell(ri, ci)
            if ri == 0:
                sc(c, v, size=Pt(10), bold=True, color=WHITE, bg=RED, bc=WHITE)
            else:
                bg = WHITE if ri % 2 == 1 else GRAY_BG
                sc(c, v, size=Pt(10),
                   bold=(ci == 0),
                   color=TEXT_DARK if ci in (0,1) else TEXT_GRAY,
                   bg=bg, bc=GRAY_LINE)

# 중앙 세로 구분선
box(slide, Inches(6.63), Inches(1.38), Inches(0.04), Inches(5.7), fill=RED_PALE)

# 비고
txt(slide, "※  구역 담당자: 해당 구역의 관리 책임자  |  청소 담당자: 실제 청소 수행 인원",
    Inches(0.35), Inches(7.16), Inches(12.5), Inches(0.28),
    size=Pt(8.5), color=TEXT_GRAY, italic=True)


# ════════════════════════════════════════
# 저장
# ════════════════════════════════════════
out = r"C:\Users\최다빈\Desktop\PPT\병리검사실_청소계획안_고도화.pptx"
prs.save(out)
print("완료:", out)
