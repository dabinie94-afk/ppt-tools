# -*- coding: utf-8 -*-
"""
CEO 보고용 PPT (4장) - 씨젠의료재단 CI 적용
슬라이드: 1.제목 / 2.목차 / 3.내용 / 4.Sheet2 표
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ── CI 브랜드 색상 ──────────────────────────────────────
CI_RED    = RGBColor(0xCB, 0x09, 0x13)   # 씨젠 시그니처 레드
CI_DARK   = RGBColor(0x24, 0x19, 0x16)   # 씨젠 차콜 다크
CI_GRAY   = RGBColor(0xF4, 0xF4, 0xF4)   # 배경 연회색
CI_LINE   = RGBColor(0xDD, 0xDD, 0xDD)   # 구분선
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_GRAY = RGBColor(0x55, 0x55, 0x55)
RED_LIGHT = RGBColor(0xFB, 0xEC, 0xED)   # 레드 연한 배경

W = Inches(13.33)
H = Inches(7.5)
LOGO = r"C:\Users\최다빈\Desktop\_ci07.png"   # 로고 (심볼+국문+영문)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H
blank = prs.slide_layouts[6]


# ════════════════════════════════════════
# 공통 헬퍼
# ════════════════════════════════════════
def rgb_hex(rgb):
    return '{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2])

def box(slide, l, t, w, h, fill=None, line_color=None, line_w=Pt(0.75)):
    shp = slide.shapes.add_shape(1, l, t, w, h)
    if fill:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if line_color:
        shp.line.width = line_w
        shp.line.color.rgb = line_color
    else:
        shp.line.fill.background()
    return shp

def txt(slide, text, l, t, w, h,
        size=Pt(11), bold=False, color=CI_DARK,
        align=PP_ALIGN.LEFT, italic=False, wrap=True):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name   = "맑은 고딕"
    r.font.size   = size
    r.font.bold   = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb

def logo(slide, l, t, h=Inches(0.45)):
    """로고 이미지 삽입 (비율 유지)"""
    from PIL import Image
    img = Image.open(LOGO)
    ratio = img.width / img.height
    slide.shapes.add_picture(LOGO, l, t, width=int(h * ratio), height=h)

def cell_bg(cell, rgb):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('a:solidFill')):
        tcPr.remove(old)
    sf = etree.SubElement(tcPr, qn('a:solidFill'))
    sr = etree.SubElement(sf,   qn('a:srgbClr'))
    sr.set('val', rgb_hex(rgb))

def cell_border_all(cell, color=CI_LINE, w=Pt(0.5)):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for side in ('a:lnL','a:lnR','a:lnT','a:lnB'):
        ln = etree.SubElement(tcPr, qn(side))
        ln.set('w', str(int(w)))
        sf = etree.SubElement(ln, qn('a:solidFill'))
        sr = etree.SubElement(sf, qn('a:srgbClr'))
        sr.set('val', rgb_hex(color))

def sc(cell, text, size=Pt(10), bold=False,
       color=CI_DARK, align=PP_ALIGN.CENTER,
       bg=None, border_c=CI_LINE, wrap=True, italic=False):
    """셀 스타일링"""
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name   = "맑은 고딕"
    r.font.size   = size
    r.font.bold   = bold
    r.font.italic = italic
    r.font.color.rgb = color
    if bg:
        cell_bg(cell, bg)
    cell_border_all(cell, border_c)

def page_frame(slide, show_logo=True):
    """공통 레이아웃: 상단 레드바 + 하단 회색바 + 로고"""
    # 상단 레드 바
    box(slide, 0, 0, W, Inches(0.08), fill=CI_RED)
    # 하단 바
    box(slide, 0, H - Inches(0.38), W, Inches(0.38), fill=CI_DARK)
    txt(slide, "GA Intelligence  |  총무팀  |  대외비",
        Inches(0.35), H - Inches(0.35), Inches(6), Inches(0.3),
        size=Pt(8.5), color=RGBColor(0xAA, 0xAA, 0xAA))
    if show_logo:
        logo(slide, Inches(11.6), H - Inches(0.5), h=Inches(0.32))


# ════════════════════════════════════════
# SLIDE 1 — 제목
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)

# 배경
box(slide, 0, 0, W, H, fill=WHITE)

# 상단 레드 바 (두껍게)
box(slide, 0, 0, W, Inches(0.12), fill=CI_RED)

# 왼쪽 세로 레드 라인
box(slide, Inches(1.2), Inches(2.2), Inches(0.06), Inches(3.0), fill=CI_RED)

# 메인 제목
txt(slide, "신규입사자 지원물품",
    Inches(1.5), Inches(2.3), Inches(10), Inches(1.1),
    size=Pt(40), bold=True, color=CI_DARK)
txt(slide, "지급기준 및 개선방안",
    Inches(1.5), Inches(3.3), Inches(10), Inches(1.1),
    size=Pt(40), bold=True, color=CI_RED)

# 구분선
box(slide, Inches(1.5), Inches(4.55), Inches(6.5), Inches(0.03), fill=CI_LINE)

# 보고 정보
meta = [
    ("보고부서", "GA Intelligence  |  총무팀"),
    ("보고일자", "2026년 5월"),
    ("문서등급", "대  외  비"),
]
for i, (k, v) in enumerate(meta):
    top = Inches(4.7) + i * Inches(0.45)
    txt(slide, k, Inches(1.5), top, Inches(1.3), Inches(0.4),
        size=Pt(10), bold=True, color=CI_RED)
    txt(slide, v, Inches(2.85), top, Inches(5.0), Inches(0.4),
        size=Pt(10), color=TEXT_GRAY)

# 로고 (우하단)
logo(slide, Inches(9.8), Inches(5.8), h=Inches(0.72))

# 하단 레드 바
box(slide, 0, H - Inches(0.1), W, Inches(0.1), fill=CI_RED)


# ════════════════════════════════════════
# SLIDE 2 — 목차
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)
box(slide, 0, 0, W, H, fill=WHITE)
page_frame(slide)

# 헤더 영역
box(slide, 0, Inches(0.08), W, Inches(0.82), fill=CI_DARK)
txt(slide, "목  차  /  Contents",
    Inches(0.5), Inches(0.18), Inches(8), Inches(0.65),
    size=Pt(22), bold=True, color=WHITE)
logo(slide, Inches(11.0), Inches(0.2), h=Inches(0.44))

# 목차 항목
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
    top = Inches(1.25) + i * Inches(1.4)
    # 번호 박스
    box(slide, Inches(0.5), top, Inches(0.75), Inches(0.75), fill=CI_RED)
    txt(slide, num,
        Inches(0.5), top, Inches(0.75), Inches(0.75),
        size=Pt(17), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # 제목
    txt(slide, title,
        Inches(1.45), top + Inches(0.04), Inches(8), Inches(0.4),
        size=Pt(16), bold=True, color=CI_DARK)
    # 설명
    txt(slide, desc,
        Inches(1.45), top + Inches(0.42), Inches(11), Inches(0.32),
        size=Pt(10.5), color=TEXT_GRAY)
    # 구분선 (마지막 제외)
    if i < 3:
        box(slide, Inches(0.5), top + Inches(0.88),
            W - Inches(1.0), Inches(0.02), fill=CI_LINE)


# ════════════════════════════════════════
# SLIDE 3 — 내용 (Sheet1)
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)
box(slide, 0, 0, W, H, fill=WHITE)
page_frame(slide)

# 헤더
box(slide, 0, Inches(0.08), W, Inches(0.82), fill=CI_DARK)
txt(slide, "주요 내용",
    Inches(0.5), Inches(0.18), Inches(8), Inches(0.65),
    size=Pt(22), bold=True, color=WHITE)
logo(slide, Inches(11.0), Inches(0.2), h=Inches(0.44))

# ── 좌측 컬럼 ─────────────────────────────────
# [1] 추진 배경 및 목적
box(slide, Inches(0.35), Inches(1.08), Inches(0.06), Inches(1.85), fill=CI_RED)
txt(slide, "01  추진 배경 및 목적",
    Inches(0.55), Inches(1.08), Inches(5.8), Inches(0.38),
    size=Pt(13), bold=True, color=CI_RED)

items_bg = [
    ("자산 관리 책임감 부여",
     "사원증 및 부속품(케이스·목걸이)의 횟수 제한 없는 무상 재지급 제도를\n악용한 부주의 구제 사례 증가"),
    ("비용 절감 및 행정 효율화",
     "무분별한 재발급에 따른 총무팀 행정 소요 감소 및\n소모성 재경 비용 누수 방지"),
]
for j, (t_title, t_body) in enumerate(items_bg):
    ty = Inches(1.5) + j * Inches(0.72)
    txt(slide, f"·  {t_title}",
        Inches(0.6), ty, Inches(5.8), Inches(0.3),
        size=Pt(10.5), bold=True, color=CI_DARK)
    txt(slide, t_body,
        Inches(0.75), ty + Inches(0.28), Inches(5.65), Inches(0.44),
        size=Pt(9.5), color=TEXT_GRAY, wrap=True)

# 조정 대상 강조
box(slide, Inches(0.55), Inches(2.95), Inches(5.8), Inches(0.46),
    fill=RED_LIGHT, line_color=CI_RED)
txt(slide, "▶  조정 대상 : 사원증  /  사원증 케이스  /  사원증 목걸이  (총 3종)",
    Inches(0.7), Inches(3.0), Inches(5.6), Inches(0.36),
    size=Pt(10), bold=True, color=CI_RED)

# [2] 현행 및 변경(안) 비교
box(slide, Inches(0.35), Inches(3.6), Inches(0.06), Inches(1.0), fill=CI_RED)
txt(slide, "02  현행 및 변경(안) 비교",
    Inches(0.55), Inches(3.6), Inches(5.8), Inches(0.38),
    size=Pt(13), bold=True, color=CI_RED)

compare_rows = [
    ("구  분",    "현  행",               "변경(안)"),
    ("사원증",    "분실·파손 시 무상 재지급\n(횟수 제한 없음)", "급여 100% 공제 후 재지급"),
    ("케이스·목걸이", "좌동",             "좌동 → 급여 공제"),
    ("예외 조항", "(없음)",               "근속 5년↑ 자연 노후화 시\n1회 무상 교체"),
]
cw3 = [Inches(1.4), Inches(2.2), Inches(2.2)]
tbl_s = slide.shapes.add_table(4, 3,
    Inches(0.55), Inches(4.05), sum(cw3), Inches(2.98)).table
for ci, cw in enumerate(cw3):
    tbl_s.columns[ci].width = cw
for ri, row in enumerate(compare_rows):
    for ci, v in enumerate(row):
        c = tbl_s.cell(ri, ci)
        if ri == 0:
            sc(c, v, size=Pt(9.5), bold=True, color=WHITE, bg=CI_DARK)
        else:
            bg = CI_GRAY if ri % 2 == 1 else WHITE
            is_change = (ci == 2 and ri in (1, 2))
            sc(c, v, size=Pt(9), bold=is_change,
               color=CI_RED if is_change else (CI_DARK if ci==0 else TEXT_GRAY),
               bg=RED_LIGHT if is_change else bg, wrap=True)

# ── 우측 컬럼 ─────────────────────────────────
# [3] 기대 효과
box(slide, Inches(7.0), Inches(1.08), Inches(0.06), Inches(2.3), fill=CI_RED)
txt(slide, "03  기대 효과",
    Inches(7.2), Inches(1.08), Inches(5.8), Inches(0.38),
    size=Pt(13), bold=True, color=CI_RED)

effects = [
    ("💰 재경 비용 감소",  "연간 사원증·부속품 제작비 절감"),
    ("🔒 출입 보안 강화",  "분실 경각심 고취 → 보안 사고 예방"),
    ("✅ 직원 수용성 확보", "5년↑ 자연 노후화 1회 무상 예외 조항 포함"),
]
for j, (t_title, t_body) in enumerate(effects):
    ty = Inches(1.5) + j * Inches(0.72)
    box(slide, Inches(7.2), ty + Inches(0.06),
        Inches(0.3), Inches(0.3), fill=CI_RED)
    txt(slide, t_title,
        Inches(7.65), ty, Inches(5.5), Inches(0.32),
        size=Pt(10.5), bold=True, color=CI_DARK)
    txt(slide, t_body,
        Inches(7.65), ty + Inches(0.3), Inches(5.4), Inches(0.36),
        size=Pt(9.5), color=TEXT_GRAY)

# [4] 향후 일정
box(slide, Inches(7.0), Inches(3.6), Inches(0.06), Inches(3.38), fill=CI_RED)
txt(slide, "04  향후 일정",
    Inches(7.2), Inches(3.6), Inches(5.8), Inches(0.38),
    size=Pt(13), bold=True, color=CI_RED)

timeline = [
    ("2026. 05", "사내 규정 개정 및 재경·인사팀 프로세스 연동"),
    ("2026. 05", "전사 공지 및 전파 (유예기간 1주일 부여)"),
    ("2026. 06. 01", "변경 기준 전격 시행  ★"),
]
for j, (date, desc) in enumerate(timeline):
    ty = Inches(4.08) + j * Inches(0.95)
    # 날짜 뱃지
    box(slide, Inches(7.2), ty, Inches(1.6), Inches(0.36), fill=CI_RED)
    txt(slide, date,
        Inches(7.2), ty, Inches(1.6), Inches(0.36),
        size=Pt(9.5), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # 내용
    box(slide, Inches(8.85), ty, Inches(4.1), Inches(0.36),
        fill=CI_GRAY, line_color=CI_LINE)
    txt(slide, desc,
        Inches(9.0), ty, Inches(3.9), Inches(0.36),
        size=Pt(9.5), color=CI_DARK)
    # 연결선
    if j < 2:
        box(slide, Inches(7.98), ty + Inches(0.38),
            Inches(0.03), Inches(0.55), fill=CI_LINE)

# 중앙 구분선
box(slide, Inches(6.7), Inches(1.08), Inches(0.02), Inches(6.0), fill=CI_LINE)


# ════════════════════════════════════════
# SLIDE 4 — Sheet2 표 (원본 그대로)
# ════════════════════════════════════════
slide = prs.slides.add_slide(blank)
box(slide, 0, 0, W, H, fill=WHITE)
page_frame(slide)

# 헤더
box(slide, 0, Inches(0.08), W, Inches(0.82), fill=CI_DARK)
txt(slide, "신규입사자 지원물품 지급기준표",
    Inches(0.5), Inches(0.18), Inches(10), Inches(0.65),
    size=Pt(22), bold=True, color=WHITE)
logo(slide, Inches(11.0), Inches(0.2), h=Inches(0.44))

# ── 표 구성 ───────────────────────────────────
# 2행 헤더 + 15행 데이터 = 17행 / 7열
# 헤더 Row1: 품목명 | 현재지급기준(2칸) | 변경기준안(2칸) | 예산부서 | 단가
# 헤더 Row2: -     | 최초지급 | 재지급 | 최초지급 | 재지급 | -      | -
rows_data = [
    # (품목, 현재-최초, 현재-재지급, 변경-최초, 변경-재지급, 예산부서, 단가, 변경여부)
    ("사원증",
     "입사 시 1매 지급",
     "분실·파손 시\n횟수 제한 없이 무상 재지급",
     "좌동",
     "분실·파손 시\n급여 100% 공제 후 재지급",
     "총무팀", "8,250원\n(재발급 46,750원)", True),
    ("사원증 케이스",
     "입사 시 1개 지급",
     "분실·파손 시\n횟수 제한 없이 무상 재지급",
     "좌동",
     "분실·파손 시\n급여 100% 공제 후 재지급",
     "총무팀", "30,800원", True),
    ("사원증 목걸이",
     "입사 시 1개 지급",
     "분실·파손 시\n횟수 제한 없이 무상 재지급",
     "좌동",
     "분실·파손 시\n급여 100% 공제 후 재지급",
     "총무팀", "7,700원", True),
    ("검사화",
     "입사 시 1켤레 지급",
     "구매요청서 작성 시 발주 및 지급",
     "좌동", "좌동",
     "각 부서", "49,500~151,800원", False),
    ("검사가운",
     "입사 시 2벌 지급",
     "구매요청서 작성 시 발주 및 지급",
     "좌동", "좌동",
     "각 부서", "27,500원", False),
    ("검사복",
     "입사 시 1벌 지급",
     "구매요청서 작성 시 발주 및 지급",
     "좌동", "좌동",
     "각 부서", "74,800원", False),
    ("전문의가운",
     "입사 시 2벌 지급",
     "구매요청서 작성 시 발주 및 지급",
     "좌동", "좌동",
     "각 부서", "77,000원", False),
    ("동계피복\n(점퍼)",
     "입사 년도 최초 지급",
     "없음",
     "좌동", "좌동",
     "총무팀", "100,000원", False),
    ("동계피복\n(조끼)",
     "입사 년도 최초 지급",
     "없음",
     "좌동", "좌동",
     "총무팀", "90,300원", False),
    ("명함",
     "오피스디포\n개별 신청 후 지급",
     "필요 시\n오피스디포 개별 신청",
     "좌동", "좌동",
     "—", "10,309원", False),
    ("다이어리",
     "입사 시 1권 지급",
     "없음",
     "좌동", "좌동",
     "학술홍보팀", "—", False),
    ("탁상달력",
     "입사 시 1부 지급",
     "없음",
     "좌동", "좌동",
     "학술홍보팀", "—", False),
    ("법인폰\n(지점 해당)",
     "입사 시 1개 지급",
     "없음",
     "좌동", "좌동",
     "총무팀", "30,000원", False),
    ("법인차량\n(임원·전문의·지점)",
     "기준에 따른 지급",
     "—",
     "좌동", "—",
     "—", "직급별 상이", False),
    ("태블릿\n(임원·전문의)",
     "입사 시 1개 지급",
     "없음",
     "좌동", "좌동",
     "—", "1,200,000원", False),
]

# 열 너비: 품목명 | 현재-최초 | 현재-재지급 | 변경-최초 | 변경-재지급 | 예산부서 | 단가
col_ws = [Inches(1.35), Inches(1.55), Inches(2.25), Inches(1.25), Inches(2.25), Inches(1.05), Inches(1.6)]
n_rows = 2 + len(rows_data)   # 헤더2 + 데이터15

tbl = slide.shapes.add_table(
    n_rows, 7,
    Inches(0.3), Inches(1.02),
    sum(col_ws), Inches(6.1)
).table

for ci, cw in enumerate(col_ws):
    tbl.columns[ci].width = cw

# ── 헤더 Row 0 ──────────────────────────────
header1 = [
    "품목명",
    "현재 지급 기준",   # 2칸 span (직접 텍스트만, 병합은 아래서)
    "",
    "변경 기준(안)",
    "",
    "예산부서",
    "단가(원)\n*1개당 기준",
]
for ci, v in enumerate(header1):
    c = tbl.cell(0, ci)
    sc(c, v, size=Pt(10), bold=True, color=WHITE, bg=CI_DARK,
       border_c=WHITE)

# ── 헤더 Row 1 ──────────────────────────────
header2 = ["", "최초 지급", "재 지급", "최초 지급", "재 지급", "", ""]
for ci, v in enumerate(header2):
    c = tbl.cell(1, ci)
    bg = RGBColor(0x44, 0x29, 0x26) if v else CI_DARK
    sc(c, v, size=Pt(9.5), bold=True, color=WHITE, bg=bg, border_c=WHITE)

# ── 데이터 행 ───────────────────────────────
for ri, (name, c1, c2, c3, c4, dept, price, changed) in enumerate(rows_data):
    row_idx = ri + 2
    base_bg = WHITE if ri % 2 == 0 else CI_GRAY

    for ci, val in enumerate([name, c1, c2, c3, c4, dept, price]):
        cell = tbl.cell(row_idx, ci)
        if changed:
            # 변경 대상 3종
            if ci == 0:
                sc(cell, val, size=Pt(9), bold=True, color=CI_RED,
                   bg=RED_LIGHT, border_c=CI_LINE)
            elif ci == 4:
                # 변경안 재지급 열 강조
                sc(cell, val, size=Pt(8.5), bold=True, color=CI_RED,
                   bg=RED_LIGHT, border_c=CI_RED, wrap=True)
            else:
                sc(cell, val, size=Pt(8.5), color=TEXT_GRAY,
                   bg=RED_LIGHT, border_c=CI_LINE, wrap=True)
        else:
            sc(cell, val,
               size=Pt(9) if ci == 0 else Pt(8.5),
               bold=(ci == 0),
               color=CI_DARK if ci == 0 else TEXT_GRAY,
               bg=base_bg, wrap=True)

# 범례
txt(slide,
    "※  붉은색 강조 = 이번 개정 대상 3종 (사원증·사원증 케이스·사원증 목걸이)  |  나머지 항목은 현행 기준 유지",
    Inches(0.3), Inches(7.14), Inches(12), Inches(0.3),
    size=Pt(8.5), color=TEXT_GRAY, italic=True)


# ════════════════════════════════════════
# 저장
# ════════════════════════════════════════
out = r"C:\Users\최다빈\Desktop\PPT\신규입사자_지원물품_CEO보고서_v3.pptx"
prs.save(out)
print("완료:", out)
