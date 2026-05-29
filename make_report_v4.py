# -*- coding: utf-8 -*-
"""
CEO 보고용 PPT (4장) - 씨젠의료재단 CI
컬러: 메인 레드 #CB0913 / 서브 차콜 #241916 (최소) / 흰색 기반
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ── 컬러 팔레트 ─────────────────────────────────
RED       = RGBColor(0xCB, 0x09, 0x13)   # 메인
CHARCOAL  = RGBColor(0x24, 0x19, 0x16)   # 서브 (최소 사용)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
RED_PALE  = RGBColor(0xFD, 0xEC, 0xED)   # 레드 연한 배경
RED_MID   = RGBColor(0xE8, 0x30, 0x3D)   # 레드 밝은 변형
GRAY_LINE = RGBColor(0xE8, 0xE8, 0xE8)   # 구분선
GRAY_BG   = RGBColor(0xF9, 0xF9, 0xF9)   # 표 짝수행
TEXT_DARK = RGBColor(0x1A, 0x1A, 0x1A)   # 본문 텍스트 (거의 검정)
TEXT_GRAY = RGBColor(0x66, 0x66, 0x66)   # 보조 텍스트

W = Inches(13.33)
H = Inches(7.5)
LOGO_PATH = r"C:\Users\최다빈\Desktop\_ci07.png"

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank = prs.slide_layouts[6]


# ════════════════════════════════════════
# 헬퍼
# ════════════════════════════════════════
def rgb_hex(rgb):
    return '{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])

def box(slide, l, t, w, h, fill=None, line_c=None, lw=Pt(0.75)):
    shp = slide.shapes.add_shape(1, l, t, w, h)
    if fill:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if line_c:
        shp.line.width = lw; shp.line.color.rgb = line_c
    else:
        shp.line.fill.background()
    return shp

def txt(slide, text, l, t, w, h,
        size=Pt(11), bold=False, color=TEXT_DARK,
        align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = "맑은 고딕"
    r.font.size = size
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb

def add_logo(slide, l, t, h=Inches(0.42)):
    from PIL import Image
    img = Image.open(LOGO_PATH)
    w = int(h * img.width / img.height)
    slide.shapes.add_picture(LOGO_PATH, l, t, width=w, height=h)

def cell_bg(cell, rgb):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('a:solidFill')):
        tcPr.remove(old)
    sf = etree.SubElement(tcPr, qn('a:solidFill'))
    sr = etree.SubElement(sf,   qn('a:srgbClr'))
    sr.set('val', rgb_hex(rgb))

def cell_border(cell, color=GRAY_LINE, w=Pt(0.5)):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for side in ('a:lnL','a:lnR','a:lnT','a:lnB'):
        ln = etree.SubElement(tcPr, qn(side))
        ln.set('w', str(int(w)))
        sf = etree.SubElement(ln, qn('a:solidFill'))
        sr = etree.SubElement(sf, qn('a:srgbClr'))
        sr.set('val', rgb_hex(color))

def sc(cell, text, size=Pt(10), bold=False,
       color=TEXT_DARK, align=PP_ALIGN.CENTER,
       bg=None, bc=GRAY_LINE, wrap=True, italic=False):
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = "맑은 고딕"
    r.font.size = size
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    if bg:
        cell_bg(cell, bg)
    cell_border(cell, bc)

def common_frame(slide, title_text, page_text):
    """공통 프레임: 상단 레드 헤더 + 하단 레드 라인"""
    # 상단 레드 헤더 바
    box(slide, 0, 0, W, Inches(0.82), fill=RED)
    # 헤더 좌측 흰색 세로 포인트
    box(slide, Inches(0.3), Inches(0.14), Inches(0.05), Inches(0.54), fill=WHITE)
    txt(slide, title_text,
        Inches(0.5), Inches(0.14), Inches(9.5), Inches(0.56),
        size=Pt(22), bold=True, color=WHITE)
    txt(slide, page_text,
        Inches(11.8), Inches(0.25), Inches(1.4), Inches(0.35),
        size=Pt(10), color=RGBColor(0xFF, 0xBB, 0xBB), align=PP_ALIGN.RIGHT)
    add_logo(slide, Inches(10.8), Inches(0.2), h=Inches(0.42))
    # 하단 레드 라인
    box(slide, 0, H - Inches(0.32), W, Inches(0.32), fill=RED)
    txt(slide, "GA Intelligence  |  총무팀  |  대외비",
        Inches(0.35), H - Inches(0.3), Inches(6), Inches(0.26),
        size=Pt(8.5), color=WHITE)


# ════════════════════════════════════════
# SLIDE 1 — 제목
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)

# 흰색 배경
box(slide, 0, 0, W, H, fill=WHITE)

# 상단: 레드 전체 상단 영역
box(slide, 0, 0, W, Inches(2.6), fill=RED)

# 로고 (상단 레드 영역 우측)
add_logo(slide, Inches(10.5), Inches(0.25), h=Inches(0.52))

# 상단 레드 안에 작은 흰색 텍스트
txt(slide, "GA Intelligence  |  총무팀",
    Inches(0.6), Inches(0.2), Inches(6), Inches(0.38),
    size=Pt(10), color=RGBColor(0xFF, 0xBB, 0xBB))

# 메인 제목 (레드 영역 안, 흰색 텍스트)
txt(slide, "신규입사자 지원물품\n지급기준 및 개선방안",
    Inches(0.6), Inches(0.72), Inches(11), Inches(1.7),
    size=Pt(38), bold=True, color=WHITE)

# 흰색 영역 — 부제 및 메타
txt(slide, "사원증 재발급 기준 현실화를 통한 비용 절감 및 출입 보안 강화",
    Inches(0.6), Inches(2.82), Inches(11), Inches(0.5),
    size=Pt(14), color=RED, bold=False)

# 레드 구분선
box(slide, Inches(0.6), Inches(3.45), Inches(7.5), Inches(0.04), fill=RED)

# 보고 메타 정보
meta = [
    ("보  고  부  서", "GA Intelligence  |  총무팀"),
    ("보  고  일  자", "2026년 5월"),
    ("문  서  등  급", "대  외  비"),
]
for i, (k, v) in enumerate(meta):
    ty = Inches(3.65) + i * Inches(0.58)
    # 키 배경 박스
    box(slide, Inches(0.6), ty, Inches(1.9), Inches(0.42), fill=RED)
    txt(slide, k, Inches(0.6), ty, Inches(1.9), Inches(0.42),
        size=Pt(10), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txt(slide, v, Inches(2.65), ty, Inches(6), Inches(0.42),
        size=Pt(10), color=TEXT_DARK, bold=False)

# 하단 레드 라인
box(slide, 0, H - Inches(0.12), W, Inches(0.12), fill=RED)


# ════════════════════════════════════════
# SLIDE 2 — 목차
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)
box(slide, 0, 0, W, H, fill=WHITE)
common_frame(slide, "목  차  /  Contents", "2 / 4")

toc = [
    ("01", "추진 배경 및 목적",
     "사원증 무상 재지급 남용 사례 증가에 따른 제도 개선 필요성"),
    ("02", "현행 및 변경(안) 비교",
     "조정 대상 3종(사원증·케이스·목걸이)의 기준 변경 내용"),
    ("03", "기대 효과 및 향후 일정",
     "비용 절감·보안 강화·직원 수용성 확보 및 2026.06.01 시행 계획"),
    ("04", "신규입사자 지원물품 지급기준표",
     "전체 14개 품목 지급기준 총괄표 (현행 → 변경안 비교)"),
]

for i, (num, title, desc) in enumerate(toc):
    top = Inches(1.0) + i * Inches(1.42)

    # 번호 레드 박스
    box(slide, Inches(0.5), top + Inches(0.06),
        Inches(0.7), Inches(0.7), fill=RED)
    txt(slide, num,
        Inches(0.5), top + Inches(0.06), Inches(0.7), Inches(0.7),
        size=Pt(17), bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # 제목
    txt(slide, title,
        Inches(1.38), top + Inches(0.06), Inches(11), Inches(0.4),
        size=Pt(16), bold=True, color=RED)

    # 설명
    txt(slide, desc,
        Inches(1.38), top + Inches(0.48), Inches(11), Inches(0.32),
        size=Pt(10.5), color=TEXT_GRAY)

    # 구분선 (마지막 제외)
    if i < 3:
        box(slide, Inches(0.5), top + Inches(0.95),
            W - Inches(1.0), Inches(0.02), fill=RED_PALE)


# ════════════════════════════════════════
# SLIDE 3 — 내용 (Sheet1 기반)
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)
box(slide, 0, 0, W, H, fill=WHITE)
common_frame(slide, "주요 내용", "3 / 4")

# 중앙 수직 구분선
box(slide, Inches(6.65), Inches(0.95), Inches(0.02), Inches(6.1), fill=GRAY_LINE)

# ── 좌측 ──────────────────────────────────────
# [01] 추진 배경 및 목적
box(slide, Inches(0.3), Inches(1.0), Inches(0.06), Inches(2.05), fill=RED)
txt(slide, "01  추진 배경 및 목적",
    Inches(0.5), Inches(1.0), Inches(6.0), Inches(0.38),
    size=Pt(13), bold=True, color=RED)

bg_items = [
    ("자산 관리 책임감 부여",
     "사원증 및 부속품(케이스·목걸이)의 횟수 제한 없는\n무상 재지급 제도를 악용한 부주의 구제 사례 증가"),
    ("비용 절감 및 행정 효율화",
     "무분별한 재발급에 따른 총무팀 행정 소요 감소 및\n소모성 재경 비용 누수 방지"),
]
for j, (t, b) in enumerate(bg_items):
    ty = Inches(1.46) + j * Inches(0.82)
    box(slide, Inches(0.5), ty + Inches(0.04),
        Inches(0.28), Inches(0.28), fill=RED)
    txt(slide, t, Inches(0.88), ty, Inches(5.6), Inches(0.32),
        size=Pt(10.5), bold=True, color=TEXT_DARK)
    txt(slide, b, Inches(0.88), ty + Inches(0.3), Inches(5.6), Inches(0.5),
        size=Pt(9.5), color=TEXT_GRAY, wrap=True)

# 조정대상 강조 박스
box(slide, Inches(0.5), Inches(3.1), Inches(6.0), Inches(0.44),
    fill=RED, line_c=None)
txt(slide, "▶  조정 대상 : 사원증  /  사원증 케이스  /  사원증 목걸이  (총 3종)",
    Inches(0.65), Inches(3.14), Inches(5.8), Inches(0.35),
    size=Pt(10), bold=True, color=WHITE)

# [02] 현행 및 변경(안) 비교
box(slide, Inches(0.3), Inches(3.72), Inches(0.06), Inches(3.15), fill=RED)
txt(slide, "02  현행 및 변경(안) 비교",
    Inches(0.5), Inches(3.72), Inches(6.0), Inches(0.38),
    size=Pt(13), bold=True, color=RED)

compare = [
    ("구  분",         "현  행",                              "변경(안)"),
    ("사원증",         "분실·파손 시 횟수 제한 없이\n무상 재지급", "급여 100% 공제\n후 재지급"),
    ("케이스·목걸이",  "좌동",                                "좌동 → 급여 공제"),
    ("예외 조항",      "(없음)",                              "근속 5년↑ 자연 노후화\n→ 1회 무상 교체"),
]
cw3 = [Inches(1.45), Inches(2.35), Inches(2.2)]
t3 = slide.shapes.add_table(4, 3,
    Inches(0.5), Inches(4.18), sum(cw3), Inches(2.6)).table
for ci, cw in enumerate(cw3):
    t3.columns[ci].width = cw
for ri, row in enumerate(compare):
    for ci, v in enumerate(row):
        c = t3.cell(ri, ci)
        if ri == 0:
            sc(c, v, size=Pt(9.5), bold=True, color=WHITE, bg=RED, bc=WHITE)
        else:
            bg = GRAY_BG if ri % 2 == 1 else WHITE
            is_new = (ci == 2 and ri in (1, 2))
            sc(c, v, size=Pt(9),
               bold=is_new,
               color=RED if is_new else (TEXT_DARK if ci == 0 else TEXT_GRAY),
               bg=RED_PALE if is_new else bg,
               bc=GRAY_LINE, wrap=True)

# ── 우측 ──────────────────────────────────────
# [03] 기대 효과
box(slide, Inches(6.88), Inches(1.0), Inches(0.06), Inches(2.45), fill=RED)
txt(slide, "03  기대 효과",
    Inches(7.08), Inches(1.0), Inches(6.1), Inches(0.38),
    size=Pt(13), bold=True, color=RED)

effects = [
    ("재경 비용 감소",   "연간 사원증·부속품 제작비 절감,\n총무팀 행정 처리 공수 감소"),
    ("출입 보안 강화",   "분실 경각심 제고로 병원·재단\n비인가 출입 보안 사고 예방"),
    ("직원 수용성 확보", "근속 5년↑ 자연 노후화 1회 무상 예외\n조항으로 반발 심리 최소화"),
]
for j, (t, b) in enumerate(effects):
    ty = Inches(1.46) + j * Inches(0.82)
    box(slide, Inches(7.08), ty + Inches(0.04),
        Inches(0.28), Inches(0.28), fill=RED)
    txt(slide, t, Inches(7.46), ty, Inches(5.7), Inches(0.32),
        size=Pt(10.5), bold=True, color=TEXT_DARK)
    txt(slide, b, Inches(7.46), ty + Inches(0.3), Inches(5.6), Inches(0.5),
        size=Pt(9.5), color=TEXT_GRAY, wrap=True)

# [04] 향후 일정
box(slide, Inches(6.88), Inches(3.72), Inches(0.06), Inches(3.15), fill=RED)
txt(slide, "04  향후 일정",
    Inches(7.08), Inches(3.72), Inches(6.0), Inches(0.38),
    size=Pt(13), bold=True, color=RED)

timeline = [
    ("2026. 05",     "사내 규정 개정 및 재경·인사팀 프로세스 연동"),
    ("2026. 05",     "전사 공지 및 전파 (유예기간 1주일 부여)"),
    ("2026. 06. 01", "변경 기준 전격 시행"),
]
for j, (date, desc) in enumerate(timeline):
    ty = Inches(4.22) + j * Inches(0.96)
    # 날짜 레드 박스
    box(slide, Inches(7.08), ty, Inches(1.65), Inches(0.4), fill=RED)
    txt(slide, date, Inches(7.08), ty, Inches(1.65), Inches(0.4),
        size=Pt(9.5), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # 내용 박스
    box(slide, Inches(8.8), ty, Inches(4.1), Inches(0.4),
        fill=RED_PALE, line_c=RED_PALE)
    txt(slide, desc, Inches(8.95), ty, Inches(3.9), Inches(0.4),
        size=Pt(9.5), color=TEXT_DARK)
    # 수직 점선 연결
    if j < 2:
        box(slide, Inches(7.97), ty + Inches(0.42),
            Inches(0.03), Inches(0.52), fill=GRAY_LINE)


# ════════════════════════════════════════
# SLIDE 4 — Sheet2 표 (원본 그대로)
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)
box(slide, 0, 0, W, H, fill=WHITE)
common_frame(slide, "신규입사자 지원물품 지급기준표", "4 / 4")

# 표 데이터
rows = [
    ("사원증",                   "입사 시 1매 지급",          "분실·파손 시\n횟수 제한 없이 무상 재지급",   "좌동", "분실·파손 시\n급여 100% 공제 후 재지급",  "총무팀",    "8,250원\n(재발급 46,750원)", True),
    ("사원증 케이스",            "입사 시 1개 지급",          "분실·파손 시\n횟수 제한 없이 무상 재지급",   "좌동", "분실·파손 시\n급여 100% 공제 후 재지급",  "총무팀",    "30,800원",                  True),
    ("사원증 목걸이",            "입사 시 1개 지급",          "분실·파손 시\n횟수 제한 없이 무상 재지급",   "좌동", "분실·파손 시\n급여 100% 공제 후 재지급",  "총무팀",    "7,700원",                   True),
    ("검사화",                   "입사 시 1켤레 지급",        "구매요청서 작성 시\n발주 및 지급",           "좌동", "좌동",                                    "각 부서",   "49,500~151,800원",          False),
    ("검사가운",                 "입사 시 2벌 지급",          "구매요청서 작성 시\n발주 및 지급",           "좌동", "좌동",                                    "각 부서",   "27,500원",                  False),
    ("검사복",                   "입사 시 1벌 지급",          "구매요청서 작성 시\n발주 및 지급",           "좌동", "좌동",                                    "각 부서",   "74,800원",                  False),
    ("전문의가운",               "입사 시 2벌 지급",          "구매요청서 작성 시\n발주 및 지급",           "좌동", "좌동",                                    "각 부서",   "77,000원",                  False),
    ("동계피복(점퍼)",           "입사 년도 최초 지급",       "없음",                                       "좌동", "좌동",                                    "총무팀",    "100,000원",                 False),
    ("동계피복(조끼)",           "입사 년도 최초 지급",       "없음",                                       "좌동", "좌동",                                    "총무팀",    "90,300원",                  False),
    ("명함",                     "오피스디포 개별 신청",      "필요 시 오피스디포\n개별 신청",              "좌동", "좌동",                                    "—",         "10,309원",                  False),
    ("다이어리",                 "입사 시 1권 지급",          "없음",                                       "좌동", "좌동",                                    "학술홍보팀","—",                         False),
    ("탁상달력",                 "입사 시 1부 지급",          "없음",                                       "좌동", "좌동",                                    "학술홍보팀","—",                         False),
    ("법인폰\n(지점 해당)",      "입사 시 1개 지급",          "없음",                                       "좌동", "좌동",                                    "총무팀",    "30,000원",                  False),
    ("법인차량\n(임원·전문의·지점)", "기준에 따른 지급",      "—",                                          "좌동", "—",                                       "—",         "직급별 상이",                False),
    ("태블릿\n(임원·전문의)",    "입사 시 1개 지급",          "없음",                                       "좌동", "좌동",                                    "—",         "1,200,000원",               False),
]

# 열: 품목명 | 현재-최초 | 현재-재지급 | 변경-최초 | 변경-재지급 | 예산부서 | 단가
cw = [Inches(1.4), Inches(1.5), Inches(2.22), Inches(1.25), Inches(2.22), Inches(1.05), Inches(1.6)]

# 헤더 2행 + 데이터 15행
tbl = slide.shapes.add_table(
    17, 7,
    Inches(0.28), Inches(1.0),
    sum(cw), Inches(6.12)
).table
for ci, c in enumerate(cw):
    tbl.columns[ci].width = c

# 헤더 Row 0: 그룹명
h0 = ["품목명", "현재 지급 기준", "", "변경 기준(안)", "", "예산부서", "단가(원)\n*1개당 기준"]
for ci, v in enumerate(h0):
    sc(tbl.cell(0, ci), v, size=Pt(10), bold=True, color=WHITE, bg=RED, bc=WHITE)

# 헤더 Row 1: 서브명
h1 = ["", "최초 지급", "재 지급", "최초 지급", "재 지급", "", ""]
for ci, v in enumerate(h1):
    bg = RGBColor(0xE0, 0x20, 0x2B) if v else RED  # 서브헤더는 조금 밝은 레드
    sc(tbl.cell(1, ci), v, size=Pt(9.5), bold=True, color=WHITE, bg=bg, bc=WHITE)

# 데이터 행
for ri, (name, c1, c2, c3, c4, dept, price, changed) in enumerate(rows):
    ridx = ri + 2
    base = WHITE if ri % 2 == 0 else GRAY_BG
    for ci, val in enumerate([name, c1, c2, c3, c4, dept, price]):
        cell = tbl.cell(ridx, ci)
        if changed:
            if ci == 0:
                sc(cell, val, size=Pt(9), bold=True, color=RED, bg=RED_PALE, bc=GRAY_LINE)
            elif ci == 4:
                sc(cell, val, size=Pt(8.5), bold=True, color=RED, bg=RED_PALE, bc=RED, wrap=True)
            else:
                sc(cell, val, size=Pt(8.5), color=TEXT_GRAY, bg=RED_PALE, bc=GRAY_LINE, wrap=True)
        else:
            sc(cell, val,
               size=Pt(9) if ci == 0 else Pt(8.5),
               bold=(ci == 0),
               color=TEXT_DARK if ci == 0 else TEXT_GRAY,
               bg=base, wrap=True)

txt(slide,
    "※  붉은색 강조 항목 = 이번 개정 대상 (사원증 · 사원증 케이스 · 사원증 목걸이 3종)  |  나머지 11개 품목은 현행 기준 유지",
    Inches(0.28), Inches(7.15), Inches(12.5), Inches(0.28),
    size=Pt(8.5), color=TEXT_GRAY, italic=True)


# ════════════════════════════════════════
# 저장
# ════════════════════════════════════════
out = r"C:\Users\최다빈\Desktop\PPT\신규입사자_지원물품_CEO보고서_v4.pptx"
prs.save(out)
print("완료:", out)
