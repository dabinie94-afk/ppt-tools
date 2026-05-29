# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree
import copy

# ── 색상 팔레트 ──────────────────────────────────────────
NAVY       = RGBColor(0x0D, 0x2B, 0x55)   # 헤더/강조
BLUE_MID   = RGBColor(0x1A, 0x5F, 0xA8)   # 서브 강조
BLUE_LIGHT = RGBColor(0xD6, 0xE8, 0xF7)   # 표 홀수행
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
GRAY_DARK  = RGBColor(0x3A, 0x3A, 0x3A)
GRAY_LIGHT = RGBColor(0xF4, 0xF6, 0xF9)
GRAY_LINE  = RGBColor(0xCC, 0xCC, 0xCC)
ACCENT_RED = RGBColor(0xC0, 0x39, 0x2B)   # 변경 포인트 강조

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

blank = prs.slide_layouts[6]   # 완전 빈 레이아웃

# ════════════════════════════════════════════════════════
# 헬퍼 함수들
# ════════════════════════════════════════════════════════
def add_rect(slide, l, t, w, h, fill_rgb=None, line_rgb=None, line_w=Pt(0)):
    shp = slide.shapes.add_shape(1, l, t, w, h)   # MSO_SHAPE_TYPE.RECTANGLE = 1
    shp.line.width = line_w
    if fill_rgb:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill_rgb
    else:
        shp.fill.background()
    if line_rgb:
        shp.line.color.rgb = line_rgb
    else:
        shp.line.fill.background()
    return shp

def add_text(slide, text, l, t, w, h,
             font_size=Pt(12), bold=False, color=GRAY_DARK,
             align=PP_ALIGN.LEFT, v_anchor=None, wrap=True):
    from pptx.enum.text import MSO_ANCHOR
    txb = slide.shapes.add_textbox(l, t, w, h)
    tf  = txb.text_frame
    tf.word_wrap = wrap
    if v_anchor:
        tf.vertical_anchor = v_anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size  = font_size
    run.font.bold  = bold
    run.font.color.rgb = color
    run.font.name  = "맑은 고딕"
    return txb

def set_cell_bg(cell, rgb):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    solidFill = etree.SubElement(tcPr, qn('a:solidFill'))
    srgbClr   = etree.SubElement(solidFill, qn('a:srgbClr'))
    srgbClr.set('val', '{:02X}{:02X}{:02X}'.format(rgb[0], rgb[1], rgb[2]))

def style_cell(cell, text, font_size=Pt(10), bold=False,
               color=GRAY_DARK, align=PP_ALIGN.CENTER,
               bg=None, wrap=True):
    cell.text = text
    tf = cell.text_frame
    tf.word_wrap = wrap
    for para in tf.paragraphs:
        para.alignment = align
        for run in para.runs:
            run.font.name  = "맑은 고딕"
            run.font.size  = font_size
            run.font.bold  = bold
            run.font.color.rgb = color
    if bg:
        set_cell_bg(cell, bg)

def add_slide_header(slide, title_text, slide_num_text=""):
    """모든 슬라이드 공통 상단 헤더"""
    # 짙은 네이비 배너
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.75), fill_rgb=NAVY)
    # 제목
    add_text(slide, title_text,
             Inches(0.35), Inches(0.08), Inches(10), Inches(0.6),
             font_size=Pt(20), bold=True, color=WHITE)
    # 슬라이드 번호
    if slide_num_text:
        add_text(slide, slide_num_text,
                 Inches(12.3), Inches(0.18), Inches(0.8), Inches(0.4),
                 font_size=Pt(11), color=RGBColor(0xAA, 0xCC, 0xEE),
                 align=PP_ALIGN.RIGHT)
    # 하단 푸터 라인
    add_rect(slide, 0, SLIDE_H - Inches(0.3), SLIDE_W, Inches(0.3), fill_rgb=NAVY)
    add_text(slide, "GA Intelligence  |  총무팀  |  대외비",
             Inches(0.3), SLIDE_H - Inches(0.28), Inches(8), Inches(0.26),
             font_size=Pt(8), color=RGBColor(0xAA, 0xBB, 0xCC))

def section_chip(slide, text, l, t):
    """작은 섹션 칩(뱃지)"""
    add_rect(slide, l, t, Inches(0.18), Inches(0.32), fill_rgb=BLUE_MID)
    add_text(slide, text,
             l + Inches(0.22), t, Inches(3), Inches(0.32),
             font_size=Pt(13), bold=True, color=BLUE_MID)


# ════════════════════════════════════════════════════════
# SLIDE 1 — 표지
# ════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank)

# 배경 그라디언트 대체 (네이비 전체)
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_rgb=NAVY)

# 왼쪽 강조 세로 라인
add_rect(slide, Inches(0.7), Inches(1.8), Inches(0.06), Inches(3.5),
         fill_rgb=RGBColor(0x4A, 0xA8, 0xE8))

# 메인 제목
add_text(slide,
         "신규입사자 지원물품\n지급기준 및 개선방안",
         Inches(1.0), Inches(1.9), Inches(9), Inches(2.4),
         font_size=Pt(38), bold=True, color=WHITE)

# 부제목 라인
add_rect(slide, Inches(1.0), Inches(4.5), Inches(5.5), Inches(0.04),
         fill_rgb=RGBColor(0x4A, 0xA8, 0xE8))

add_text(slide, "사원증 지급·재발급 기준 현실화를 통한 비용 절감 및 보안 강화",
         Inches(1.0), Inches(4.6), Inches(9), Inches(0.5),
         font_size=Pt(14), color=RGBColor(0xAA, 0xCC, 0xEE))

# 메타 정보
add_text(slide, "보고부서   GA Intelligence  |  총무팀",
         Inches(1.0), Inches(5.4), Inches(6), Inches(0.38),
         font_size=Pt(12), color=RGBColor(0xCC, 0xDD, 0xEE))
add_text(slide, "보고일자   2026년 5월",
         Inches(1.0), Inches(5.8), Inches(6), Inches(0.38),
         font_size=Pt(12), color=RGBColor(0xCC, 0xDD, 0xEE))
add_text(slide, "문서등급   대외비",
         Inches(1.0), Inches(6.2), Inches(6), Inches(0.38),
         font_size=Pt(12), color=RGBColor(0xCC, 0xDD, 0xEE))


# ════════════════════════════════════════════════════════
# SLIDE 2 — 목차
# ════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank)
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_rgb=GRAY_LIGHT)
add_slide_header(slide, "목  차", "2 / 6")

toc_items = [
    ("01", "추진 배경 및 목적",     "사원증 무상 재발급 제도의 문제점 및 개선 필요성"),
    ("02", "현행 및 변경(안) 비교", "조정 대상 3종(사원증·케이스·목걸이) 기준 변경 내용"),
    ("03", "품목별 지급기준 전체표", "신규입사자 지원물품 14개 품목 기준 총괄"),
    ("04", "기대 효과",             "비용 절감 · 보안 강화 · 직원 불만 최소화"),
    ("05", "향후 일정",             "규정 개정 → 공지 → 2026.06.01 시행"),
]

for i, (num, title, desc) in enumerate(toc_items):
    top = Inches(1.1) + i * Inches(1.1)
    # 번호 박스
    add_rect(slide, Inches(0.8), top, Inches(0.65), Inches(0.65), fill_rgb=NAVY)
    add_text(slide, num,
             Inches(0.8), top, Inches(0.65), Inches(0.65),
             font_size=Pt(16), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # 제목
    add_text(slide, title,
             Inches(1.65), top, Inches(5), Inches(0.38),
             font_size=Pt(15), bold=True, color=NAVY)
    # 설명
    add_text(slide, desc,
             Inches(1.65), top + Inches(0.36), Inches(9), Inches(0.3),
             font_size=Pt(10.5), color=GRAY_DARK)
    # 구분선
    if i < len(toc_items) - 1:
        add_rect(slide, Inches(0.8), top + Inches(0.75),
                 Inches(11.8), Inches(0.01), fill_rgb=GRAY_LINE)


# ════════════════════════════════════════════════════════
# SLIDE 3 — 추진 배경 및 목적
# ════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank)
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_rgb=GRAY_LIGHT)
add_slide_header(slide, "01  추진 배경 및 목적", "3 / 6")

reasons = [
    ("자산 관리\n책임감 부여",
     "사원증 및 부속품(케이스·목걸이)의 횟수 제한 없는 무상 재지급 제도를 악용한\n부주의 구제 사례 지속 증가 → 직원의 자산 관리 의식 제고 필요"),
    ("비용 절감 및\n행정 효율화",
     "무분별한 재발급에 따른 총무팀 행정 소요를 줄이고,\n소모성 재경 비용의 누수를 방지하고자 지급 기준을 현실화"),
]

for i, (title, body) in enumerate(reasons):
    left = Inches(0.5) + i * Inches(6.4)
    # 카드 배경
    add_rect(slide, left, Inches(1.0), Inches(6.0), Inches(4.5),
             fill_rgb=WHITE, line_rgb=GRAY_LINE, line_w=Pt(1))
    # 상단 색 바
    add_rect(slide, left, Inches(1.0), Inches(6.0), Inches(0.55), fill_rgb=BLUE_MID)
    # 제목
    add_text(slide, title,
             left + Inches(0.25), Inches(1.05), Inches(5.5), Inches(0.5),
             font_size=Pt(14), bold=True, color=WHITE)
    # 본문
    add_text(slide, body,
             left + Inches(0.25), Inches(1.75), Inches(5.5), Inches(3.3),
             font_size=Pt(12), color=GRAY_DARK)

# 조정 대상 강조 박스
add_rect(slide, Inches(0.5), Inches(5.65), Inches(12.4), Inches(0.65),
         fill_rgb=RGBColor(0xFF, 0xF3, 0xCD), line_rgb=RGBColor(0xF0, 0xC0, 0x40), line_w=Pt(1))
add_text(slide,
         "▶  조정 대상 : 사원증 / 사원증 케이스 / 사원증 목걸이   (총 3종)  |  기타 11개 품목은 기존 기준 유지",
         Inches(0.7), Inches(5.7), Inches(12), Inches(0.5),
         font_size=Pt(12), bold=True, color=RGBColor(0x7D, 0x4E, 0x00))


# ════════════════════════════════════════════════════════
# SLIDE 4 — 현행 및 변경(안) 비교
# ════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank)
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_rgb=GRAY_LIGHT)
add_slide_header(slide, "02  현행 및 변경(안) 비교", "4 / 6")

# 테이블: 품목 / 현행 / 변경(안) / 비고
tbl_data = [
    ["구  분",          "현  행",                              "변경(안)",                          "비  고"],
    ["사원증",          "분실·파손 시\n횟수 제한 없이 무상 재지급", "분실·파손 시\n급여 100% 공제 후 재지급", "8,250원/매\n(재발급 46,750원)"],
    ["사원증 케이스",   "좌동",                                "좌동 → 급여 공제",                 "30,800원"],
    ["사원증 목걸이",   "좌동",                                "좌동 → 급여 공제",                 "7,700원"],
    ["예외 조항",       "(없음)",                              "근속 5년 이상, 자연 노후화 시\n1회 무상 교체 허용",   "인사팀 확인 후 처리"],
]

col_w = [Inches(1.8), Inches(3.6), Inches(4.2), Inches(2.8)]
tbl_l = Inches(0.6)
tbl_t = Inches(1.05)
tbl_h = Inches(5.1)
tbl_total_w = sum(col_w)

tbl = slide.shapes.add_table(len(tbl_data), 4,
                              tbl_l, tbl_t, tbl_total_w, tbl_h).table

for ci, w in enumerate(col_w):
    tbl.columns[ci].width = w

for ri, row in enumerate(tbl_data):
    for ci, val in enumerate(row):
        cell = tbl.cell(ri, ci)
        if ri == 0:
            style_cell(cell, val, font_size=Pt(12), bold=True,
                       color=WHITE, bg=NAVY, align=PP_ALIGN.CENTER)
        else:
            bg = WHITE if ri % 2 == 1 else BLUE_LIGHT
            # 변경(안) 열은 붉은 강조
            if ci == 2 and ri in (1, 2, 3):
                style_cell(cell, val, font_size=Pt(11), color=ACCENT_RED,
                           bold=True, bg=bg, align=PP_ALIGN.CENTER)
            elif ci == 0:
                style_cell(cell, val, font_size=Pt(11), bold=True,
                           color=NAVY, bg=bg, align=PP_ALIGN.CENTER)
            else:
                style_cell(cell, val, font_size=Pt(10.5),
                           color=GRAY_DARK, bg=bg, align=PP_ALIGN.CENTER)

# 범례
add_text(slide,
         "※  최초 지급 기준(입사 시 1매)은 변경 없이 유지  |  붉은 항목 = 이번 개정 대상",
         Inches(0.6), Inches(6.3), Inches(12), Inches(0.4),
         font_size=Pt(9.5), color=RGBColor(0x77, 0x77, 0x77))


# ════════════════════════════════════════════════════════
# SLIDE 5 — 품목별 지급기준 전체표
# ════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank)
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_rgb=GRAY_LIGHT)
add_slide_header(slide, "03  품목별 지급기준 전체표", "5 / 6")

full_data = [
    ["품목명",         "최초 지급",               "현행 재지급",                     "변경 재지급(안)",                     "예산부서",  "단가(원)"],
    ["사원증",         "입사 시 1매",             "횟수 제한 없이 무상",            "급여 100% 공제 후 재지급",           "총무팀",    "8,250"],
    ["사원증 케이스",  "입사 시 1개",             "횟수 제한 없이 무상",            "급여 공제 후 재지급",                "총무팀",    "30,800"],
    ["사원증 목걸이",  "입사 시 1개",             "횟수 제한 없이 무상",            "급여 공제 후 재지급",                "총무팀",    "7,700"],
    ["검사화",         "입사 시 1켤레",           "구매요청서 작성 시 발주·지급",    "좌동",                               "각 부서",   "49,500~\n151,800"],
    ["검사가운",       "입사 시 2벌",             "구매요청서 작성 시 발주·지급",    "좌동",                               "각 부서",   "27,500"],
    ["검사복",         "입사 시 1벌",             "구매요청서 작성 시 발주·지급",    "좌동",                               "각 부서",   "74,800"],
    ["전문의가운",     "입사 시 2벌",             "구매요청서 작성 시 발주·지급",    "좌동",                               "각 부서",   "77,000"],
    ["동계피복(점퍼)", "입사 년도 최초",          "없음",                            "좌동",                               "총무팀",    "100,000"],
    ["동계피복(조끼)", "입사 년도 최초",          "없음",                            "좌동",                               "총무팀",    "90,300"],
    ["명함",           "오피스디포 개별 신청",    "필요 시 개별 신청",               "좌동",                               "-",         "10,309"],
    ["다이어리·달력",  "입사 시 1권·1부",         "없음",                            "좌동",                               "학술홍보팀","—"],
    ["법인폰",         "입사 시 1개 (지점)",       "없음",                           "좌동",                               "총무팀",    "30,000"],
    ["법인차량",       "직급별 기준",             "-",                               "좌동",                               "-",         "직급별 상이"],
    ["태블릿",         "입사 시 1개 (임원·전문의)","없음",                           "좌동",                               "-",         "1,200,000"],
]

col_w2 = [Inches(1.55), Inches(1.5), Inches(2.2), Inches(2.4), Inches(1.3), Inches(1.4)]
tbl2 = slide.shapes.add_table(len(full_data), 6,
                               Inches(0.35), Inches(1.0),
                               sum(col_w2), Inches(6.1)).table

for ci, w in enumerate(col_w2):
    tbl2.columns[ci].width = w

for ri, row in enumerate(full_data):
    for ci, val in enumerate(row):
        cell = tbl2.cell(ri, ci)
        if ri == 0:
            style_cell(cell, val, font_size=Pt(10), bold=True,
                       color=WHITE, bg=NAVY, align=PP_ALIGN.CENTER)
        else:
            bg = WHITE if ri % 2 == 1 else BLUE_LIGHT
            is_changed = ri in (1, 2, 3) and ci == 3
            style_cell(cell, val,
                       font_size=Pt(9),
                       bold=(ci == 0 or is_changed),
                       color=ACCENT_RED if is_changed else (NAVY if ci == 0 else GRAY_DARK),
                       bg=bg,
                       align=PP_ALIGN.CENTER)

add_text(slide,
         "※ 음영 없음=홀수행(흰색)  /  파란 음영=짝수행  /  붉은 텍스트=이번 변경 대상 항목",
         Inches(0.35), Inches(7.15), Inches(12), Inches(0.3),
         font_size=Pt(8.5), color=RGBColor(0x88, 0x88, 0x88))


# ════════════════════════════════════════════════════════
# SLIDE 6 — 기대 효과 + 향후 일정
# ════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank)
add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill_rgb=GRAY_LIGHT)
add_slide_header(slide, "04·05  기대 효과 및 향후 일정", "6 / 6")

# ── 기대 효과 ───────────────────────────────────────────
section_chip(slide, "04  기대 효과", Inches(0.4), Inches(0.95))

effects = [
    ("💰 재경 비용 감소",
     "연간 무분별하게 지출되던\n사원증 제작비 및 부속품\n구입비 절감"),
    ("🔒 보안 강화",
     "사원증 분실에 대한 경각심 고취로\n병원·재단 내 출입 보안 사고 예방"),
    ("✅ 불만 최소화",
     "근속 5년 이상 자연 노후화 예외 조항으로\n제도 변경에 따른 반발 심리 차단"),
]

for i, (title, body) in enumerate(effects):
    lft = Inches(0.4) + i * Inches(4.2)
    add_rect(slide, lft, Inches(1.45), Inches(3.95), Inches(2.4),
             fill_rgb=WHITE, line_rgb=BLUE_MID, line_w=Pt(1.2))
    add_rect(slide, lft, Inches(1.45), Inches(3.95), Inches(0.5), fill_rgb=BLUE_MID)
    add_text(slide, title,
             lft + Inches(0.15), Inches(1.5), Inches(3.7), Inches(0.42),
             font_size=Pt(12), bold=True, color=WHITE)
    add_text(slide, body,
             lft + Inches(0.15), Inches(2.05), Inches(3.7), Inches(1.7),
             font_size=Pt(11), color=GRAY_DARK)

# ── 향후 일정 ───────────────────────────────────────────
section_chip(slide, "05  향후 일정", Inches(0.4), Inches(4.05))

timeline = [
    ("2026. 05",  "규정 개정",     "관련 사내 규정(복리후생 지급 기준) 개정\n재경·인사팀 프로세스 연동"),
    ("2026. 05",  "전사 공지",     "전사 공지 및 전파\n조직문화 유예기간 1주일 부여"),
    ("2026. 06. 01", "변경 기준 시행", "변경된 지급 기준 전격 시행"),
]

arrow_color = BLUE_MID
for i, (date, step, desc) in enumerate(timeline):
    lft = Inches(0.4) + i * Inches(4.2)
    # 날짜 박스
    add_rect(slide, lft, Inches(4.5), Inches(3.95), Inches(0.55), fill_rgb=NAVY)
    add_text(slide, date,
             lft, Inches(4.52), Inches(3.95), Inches(0.5),
             font_size=Pt(12), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # 단계명
    add_rect(slide, lft, Inches(5.08), Inches(3.95), Inches(0.42), fill_rgb=BLUE_LIGHT)
    add_text(slide, step,
             lft + Inches(0.1), Inches(5.1), Inches(3.8), Inches(0.38),
             font_size=Pt(12), bold=True, color=NAVY)
    # 설명
    add_rect(slide, lft, Inches(5.53), Inches(3.95), Inches(1.2),
             fill_rgb=WHITE, line_rgb=GRAY_LINE, line_w=Pt(0.8))
    add_text(slide, desc,
             lft + Inches(0.1), Inches(5.58), Inches(3.7), Inches(1.1),
             font_size=Pt(10.5), color=GRAY_DARK)
    # 화살표 (마지막 제외)
    if i < len(timeline) - 1:
        ax = lft + Inches(4.05)
        add_text(slide, "▶",
                 ax, Inches(5.15), Inches(0.3), Inches(0.35),
                 font_size=Pt(16), color=BLUE_MID, align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════
# 저장
# ════════════════════════════════════════════════════════
out = r"C:\Users\최다빈\Desktop\PPT\신규입사자_지원물품_지급기준_보고서.pptx"
prs.save(out)
print("저장 완료:", out)
