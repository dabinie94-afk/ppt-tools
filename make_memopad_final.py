"""
메모패드 제작업체비교 CEO 보고서
씨젠의료재단 CI 가이드라인 준수 버전
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from lxml import etree
import os

# ─────────── 색상 ───────────
RED        = RGBColor(0xCC, 0x00, 0x00)   # 씨젠 시그니처 레드
BLUE       = RGBColor(0x15, 0x4E, 0x9B)   # 비교 강조 블루
BLACK      = RGBColor(0x1A, 0x1A, 0x1A)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY      = RGBColor(0xF5, 0xF5, 0xF5)   # 연한 회색 배경
MGRAY      = RGBColor(0xD8, 0xD8, 0xD8)   # 구분선/테두리
DGRAY      = RGBColor(0x66, 0x66, 0x66)   # 보조 텍스트
LRED       = RGBColor(0xFF, 0xF0, 0xF0)   # 연한 빨강 배경
LBLUE      = RGBColor(0xEB, 0xF2, 0xFF)   # 연한 파랑 배경
TABLE_EVEN = RGBColor(0xFA, 0xFA, 0xFA)   # 테이블 짝수행

FONT = 'Malgun Gothic'

BASE = r'c:\Users\최다빈\Desktop\PPT'
LOGO = os.path.join(BASE, 'seegene_ci.png')
OUT  = r'c:\Users\최다빈\Desktop\메모패드_제작업체비교_최종.pptx'

W, H = Inches(13.333), Inches(7.5)


# ─────────── 기본 헬퍼 ───────────

def rect(slide, x, y, w, h, fill=None, line=None, lw=0.5):
    s = slide.shapes.add_shape(1, x, y, w, h)
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = line; s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    return s


def txt(slide, text, x, y, w, h, size=11, bold=False, color=BLACK,
        align=PP_ALIGN.LEFT, wrap=True, italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = FONT; r.font.size = Pt(size)
    r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    return tb


def multiline(slide, x, y, w, h, lines, wrap=True):
    """lines: list of dicts with keys text,size,bold,color,align,space_before,space_after"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ln.get('align', PP_ALIGN.LEFT)
        if ln.get('space_before'): p.space_before = Pt(ln['space_before'])
        if ln.get('space_after'):  p.space_after  = Pt(ln['space_after'])
        if ln.get('line_space'):   p.line_spacing  = Pt(ln['line_space'])
        runs = ln.get('runs')
        if runs:
            for rn in runs:
                r = p.add_run()
                r.text = rn['text']
                r.font.name = FONT
                r.font.size = Pt(rn.get('size', 11))
                r.font.bold = rn.get('bold', False)
                r.font.color.rgb = rn.get('color', BLACK)
        else:
            r = p.add_run()
            r.text = ln.get('text', '')
            r.font.name = FONT; r.font.size = Pt(ln.get('size', 11))
            r.font.bold = ln.get('bold', False)
            r.font.color.rgb = ln.get('color', BLACK)
    return tb


def tc_borders(cell, hex_color='CCCCCC', w_pt=0.5, sides=None):
    if sides is None: sides = ['lnT','lnB','lnL','lnR']
    tc   = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None: tcPr = etree.SubElement(tc, qn('a:tcPr'))
    w = str(int(w_pt * 12700))
    for side in sides:
        el = tcPr.find(qn(f'a:{side}'))
        if el is None: el = etree.SubElement(tcPr, qn(f'a:{side}'))
        el.set('w', w); el.set('cap','flat'); el.set('cmpd','sng')
        for ch in list(el): el.remove(ch)
        sf = etree.SubElement(el, qn('a:solidFill'))
        sr = etree.SubElement(sf, qn('a:srgbClr'))
        sr.set('val', hex_color)


def cell_fmt(cell, text='', size=10, bold=False, color=BLACK,
             fill=None, align=PP_ALIGN.CENTER,
             bcolor='BBBBBB', bw=0.5, mt=2, ml=5):
    if fill:
        cell.fill.solid(); cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
    tf = cell.text_frame
    tf.margin_top = Pt(mt); tf.margin_bottom = Pt(mt)
    tf.margin_left = Pt(ml); tf.margin_right = Pt(ml)
    tf.word_wrap = True
    while len(tf.paragraphs) > 1:
        tf.paragraphs[-1]._p.getparent().remove(tf.paragraphs[-1]._p)
    p = tf.paragraphs[0]; p.alignment = align
    for r in list(p.runs): r._r.getparent().remove(r._r)
    r = p.add_run(); r.text = text
    r.font.name = FONT; r.font.size = Pt(size)
    r.font.bold = bold; r.font.color.rgb = color
    tc_borders(cell, hex_color=bcolor, w_pt=bw)


def cell_multirun(cell, runs, size=10, fill=None, align=PP_ALIGN.CENTER,
                  bcolor='BBBBBB', bw=0.5, mt=2, ml=5):
    """Cell with multiple colored runs"""
    if fill:
        cell.fill.solid(); cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
    tf = cell.text_frame
    tf.margin_top = Pt(mt); tf.margin_bottom = Pt(mt)
    tf.margin_left = Pt(ml); tf.margin_right = Pt(ml)
    tf.word_wrap = True
    while len(tf.paragraphs) > 1:
        tf.paragraphs[-1]._p.getparent().remove(tf.paragraphs[-1]._p)
    p = tf.paragraphs[0]; p.alignment = align
    for r_elem in list(p.runs): r_elem._r.getparent().remove(r_elem._r)
    for rn in runs:
        r = p.add_run(); r.text = rn['text']
        r.font.name = FONT; r.font.size = Pt(rn.get('size', size))
        r.font.bold = rn.get('bold', False)
        r.font.color.rgb = rn.get('color', BLACK)
    tc_borders(cell, hex_color=bcolor, w_pt=bw)


# ─────────── 공통 슬라이드 요소 ───────────

def slide_header(slide, title, page):
    """레드 헤더 바 + 페이지 번호"""
    rect(slide, Inches(0), Inches(0), W, Inches(0.9), fill=RED)
    rect(slide, Inches(0), Inches(0.9), W, Inches(0.04), fill=MGRAY)
    txt(slide, title, Inches(0.45), Inches(0.13), Inches(10), Inches(0.65),
        size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    txt(slide, page, Inches(11.5), Inches(0.13), Inches(1.6), Inches(0.65),
        size=13, bold=False, color=RGBColor(0xFF,0xCC,0xCC), align=PP_ALIGN.RIGHT)


def slide_footer(slide, page_text):
    """하단 얇은 레드 라인 + 페이지 번호"""
    rect(slide, Inches(0), Inches(7.28), W, Inches(0.03), fill=RED)
    txt(slide, page_text, Inches(10.5), Inches(7.3), Inches(2.6), Inches(0.2),
        size=8, color=DGRAY, align=PP_ALIGN.RIGHT)


def section_label(slide, label, x, y, w=Inches(2.8), h=Inches(0.4)):
    """레드 섹션 레이블 박스"""
    rect(slide, x, y, w, h, fill=RED)
    txt(slide, label, x, y + Inches(0.02), w, h - Inches(0.04),
        size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════
#  SLIDE 1 ─ 표지
# ══════════════════════════════════════════════
prs = Presentation()
prs.slide_width, prs.slide_height = W, H
blank = prs.slide_layouts[6]

slide1 = prs.slides.add_slide(blank)

# ① 배경 전체 흰색 (기본값)
# ② 좌측 레드 세로 바
rect(slide1, Inches(0), Inches(0), Inches(0.3), H, fill=RED)

# ③ 상단 라이트그레이 밴드
rect(slide1, Inches(0.3), Inches(0), W - Inches(0.3), Inches(1.55), fill=LGRAY)

# ④ 로고
try:
    slide1.shapes.add_picture(LOGO, Inches(0.45), Inches(0.25), Inches(2.9), Inches(0.9))
except:
    txt(slide1, '씨젠의료재단', Inches(0.45), Inches(0.35), Inches(3), Inches(0.6),
        size=18, bold=True, color=RED)

# ⑤ 결재란 테이블
APPR_LABELS = ['담당', '총무팀장', '경영관리본부', '경영관리부문', '기획조정실', '행정원장', '이사장']
tbl_appr = slide1.shapes.add_table(2, 7, Inches(6.6), Inches(0.22), Inches(6.4), Inches(1.1)).table
col_w = Inches(6.4) // 7
for col in tbl_appr.columns: col.width = col_w
tbl_appr.rows[0].height = int(Inches(1.1) * 0.37)
tbl_appr.rows[1].height = int(Inches(1.1) * 0.63)
for i, lbl in enumerate(APPR_LABELS):
    cell_fmt(tbl_appr.cell(0,i), lbl, size=7.5, bold=True, color=WHITE, fill=RED,
             bcolor='FFFFFF', bw=0.5)
for i in range(7):
    cell_fmt(tbl_appr.cell(1,i), '', fill=WHITE, bcolor='AAAAAA', bw=0.5)

# ⑥ 헤더 하단 구분선
rect(slide1, Inches(0.3), Inches(1.55), W - Inches(0.3), Inches(0.05), fill=MGRAY)

# ⑦ '검토보고' 배지
rect(slide1, Inches(0.5), Inches(1.95), Inches(1.7), Inches(0.42), fill=RED)
txt(slide1, '검 토 보 고', Inches(0.5), Inches(1.96), Inches(1.7), Inches(0.4),
    size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# ⑧ 메인 타이틀
multiline(slide1, Inches(0.5), Inches(2.6), Inches(11), Inches(3.0), [
    {'text': '메모패드', 'size': 13, 'bold': False, 'color': DGRAY,
     'align': PP_ALIGN.LEFT, 'space_after': 6},
    {'text': '공급 단가 검증 및 업체별', 'size': 36, 'bold': True, 'color': BLACK,
     'align': PP_ALIGN.LEFT, 'space_after': 2},
    {'text': '견적 비교 검토', 'size': 36, 'bold': True, 'color': BLACK,
     'align': PP_ALIGN.LEFT},
])

# ⑨ 타이틀 하단 레드 라인
rect(slide1, Inches(0.5), Inches(5.25), Inches(4.0), Inches(0.06), fill=RED)

# ⑩ 하단 정보 배경
rect(slide1, Inches(0.3), Inches(6.3), W - Inches(0.3), Inches(1.2), fill=LGRAY)
rect(slide1, Inches(0.3), Inches(6.3), W - Inches(0.3), Inches(0.05), fill=RED)

# 부서/담당자/날짜
multiline(slide1, Inches(0.5), Inches(6.42), Inches(7), Inches(0.85), [
    {'text': '경영관리본부  총무팀', 'size': 10, 'bold': False, 'color': DGRAY,
     'align': PP_ALIGN.LEFT, 'space_after': 3},
    {'text': '최다빈 대리', 'size': 12, 'bold': True, 'color': BLACK, 'align': PP_ALIGN.LEFT},
])
txt(slide1, '2026. 05. 27.', Inches(9.5), Inches(6.55), Inches(3.6), Inches(0.5),
    size=11, color=DGRAY, align=PP_ALIGN.RIGHT)
txt(slide1, '01 / 04', Inches(10.5), Inches(6.95), Inches(2.6), Inches(0.35),
    size=9, color=DGRAY, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════
#  SLIDE 2 ─ 검토 배경 및 목적
# ══════════════════════════════════════════════
slide2 = prs.slides.add_slide(blank)

slide_header(slide2, '검토 배경 및 목적', '02 / 04')

# ── 좌측 섹션: 검토 목적 ──
LX = Inches(0.35)
RX = Inches(6.95)
TY = Inches(1.05)
SW = Inches(6.3)

section_label(slide2, '검토 목적', LX, TY, w=Inches(2.5))

# 3개 목적 카드
objectives = [
    {
        'num': '01',
        'title': '장기 거래처 단가의 적정성 정밀 검증',
        'body': '장기 거래처(에버컴퍼니)의 공급 단가가 시장 가격 대비 적정 수준인지 정밀 검증할 필요성 제기'
    },
    {
        'num': '02',
        'title': '신규 업체 견적의 객관적 대조 분석',
        'body': '신규 유사 업체 견적·공급 조건을 동일 기준으로 대조, 재단의 최적 예산 구조를 객관적으로 파악'
    },
    {
        'num': '03',
        'title': '관할 병·의원 영업 활용',
        'body': '관할 병·의원 네트워크 유지 활동 시 현장에서 상시 배포되는 핵심 판촉물'
    },
]

OBJ_Y = TY + Inches(0.55)
OBJ_H = Inches(1.2)
OBJ_GAP = Inches(0.1)

for i, obj in enumerate(objectives):
    oy = OBJ_Y + i * (OBJ_H + OBJ_GAP)
    # 카드 배경
    rect(slide2, LX, oy, SW, OBJ_H, fill=LGRAY, line=MGRAY, lw=0.5)
    # 번호 배지
    rect(slide2, LX, oy, Inches(0.55), OBJ_H, fill=RED)
    txt(slide2, obj['num'], LX, oy + Inches(0.35), Inches(0.55), Inches(0.5),
        size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # 제목
    txt(slide2, obj['title'], LX + Inches(0.65), oy + Inches(0.1), SW - Inches(0.75), Inches(0.4),
        size=11, bold=True, color=BLACK)
    # 내용
    txt(slide2, obj['body'], LX + Inches(0.65), oy + Inches(0.52), SW - Inches(0.75), Inches(0.65),
        size=9.5, color=DGRAY, wrap=True)

# ── 우측 섹션: 주요 소요처 및 용도 ──
section_label(slide2, '메모패드 주요 소요처 및 용도', RX, TY, w=Inches(6.0))

uses = [
    {
        'title': '교육 프로그램 및 방문객 지급',
        'body': '센터 주관 교육 프로그램, 내방객 방문 및\n학생 초청 행사 지급용'
    },
    {
        'title': '재단 홍보 및 외빈 응대 기념품',
        'body': '사내 직무 교육·대학생 견학·주요 외빈 방문 시\n재단 홍보 기념품으로 비치.'
    },
]

USE_Y = TY + Inches(0.55)
USE_H = Inches(1.78)
USE_W = Inches(6.0)

for i, use in enumerate(uses):
    uy = USE_Y + i * (USE_H + OBJ_GAP)
    rect(slide2, RX, uy, USE_W, USE_H, fill=LGRAY, line=MGRAY, lw=0.5)
    # 아이콘 영역 (레드 좌측 바)
    rect(slide2, RX, uy, Inches(0.12), USE_H, fill=RED)
    txt(slide2, use['title'], RX + Inches(0.22), uy + Inches(0.15), USE_W - Inches(0.3), Inches(0.4),
        size=11, bold=True, color=BLACK)
    txt(slide2, use['body'], RX + Inches(0.22), uy + Inches(0.6), USE_W - Inches(0.3), Inches(1.1),
        size=10, color=DGRAY, wrap=True)

# ── 의견 박스 ──
OPN_Y = Inches(5.42)
OPN_H = Inches(1.82)
rect(slide2, Inches(0.35), OPN_Y, W - Inches(0.45), OPN_H, fill=LRED, line=RED, lw=0.8)
rect(slide2, Inches(0.35), OPN_Y, Inches(0.1), OPN_H, fill=RED)

txt(slide2, '의견', Inches(0.55), OPN_Y + Inches(0.1), Inches(1.0), Inches(0.38),
    size=11, bold=True, color=RED)

multiline(slide2, Inches(0.55), OPN_Y + Inches(0.45), W - Inches(0.75), Inches(1.2), [
    {
        'runs': [
            {'text': '견적이 제출된 전 수량 구간에서 ', 'size': 10.5, 'color': BLACK},
            {'text': '아이디컴이 최저가를 형성', 'size': 10.5, 'bold': True, 'color': RED},
            {'text': '하였으며, 제작 기간 단축 및 샘플 대응 가능 측면에서도 차별점이 확인됨.', 'size': 10.5, 'color': BLACK},
        ],
        'align': PP_ALIGN.LEFT,
    },
    {
        'runs': [
            {'text': '기존 대비 약 ', 'size': 10.5, 'color': BLACK},
            {'text': '40.3% ~ 47.3%', 'size': 11.5, 'bold': True, 'color': RED},
            {'text': ' 예산이 절감되는 효과가 있으나, 실제 도입 전 사전 샘플을 통한 품질 검증 과정이 필요할 것으로 보임.', 'size': 10.5, 'color': BLACK},
        ],
        'align': PP_ALIGN.LEFT,
        'space_before': 5,
    },
])

slide_footer(slide2, '02 / 04')


# ══════════════════════════════════════════════
#  SLIDE 3 ─ 수량별 업체 비교표
# ══════════════════════════════════════════════
slide3 = prs.slides.add_slide(blank)
slide_header(slide3, '수량별 업체 비교표', '04 / 04')

# 테이블 설정
# 7 rows (header + 6 data), 4 cols
COLS = 4; ROWS = 7
TBL_X = Inches(0.4)
TBL_Y = Inches(1.1)
TBL_W = W - Inches(0.8)
TBL_H = Inches(6.1)

tbl3 = slide3.shapes.add_table(ROWS, COLS, TBL_X, TBL_Y, TBL_W, TBL_H).table

# 열 너비 설정
tbl3.columns[0].width = Inches(2.4)   # 구분
tbl3.columns[1].width = Inches(3.1)   # 에버컴퍼니
tbl3.columns[2].width = Inches(3.1)   # 이야기
tbl3.columns[3].width = Inches(3.7)   # 아이디컴

# 행 높이
header_h = Inches(0.75)
data_h   = Inches(0.88)
tbl3.rows[0].height = header_h
for i in range(1, ROWS):
    tbl3.rows[i].height = data_h

HDRS = ['구분', '에버컴퍼니 (기존)', '이야기', '아이디컴']
# 헤더 행
for j, h in enumerate(HDRS):
    c = tbl3.cell(0, j)
    if j == 3:
        cell_fmt(c, h, size=12, bold=True, color=WHITE, fill=RED, bcolor='FFFFFF', bw=0.5)
    elif j == 0:
        cell_fmt(c, h, size=12, bold=True, color=WHITE, fill=RGBColor(0x33,0x33,0x33), bcolor='FFFFFF', bw=0.5)
    else:
        cell_fmt(c, h, size=12, bold=True, color=WHITE, fill=RGBColor(0x55,0x55,0x55), bcolor='FFFFFF', bw=0.5)

# 데이터 행 정의
rows_data = [
    ('500개 구간',   '견적 없음',  '견적 없음',  ('3,300원', True)),
    ('1,000개 구간', '5,500원',    '5,200원',    ('2,900원', True)),
    ('1,500개 구간', '4,590원',    '견적 없음',  ('2,600원', True)),
    ('2,000개 구간', '3,850원',    '4,700원',    ('2,300원', True)),
    ('제작 기간',    '4~5주 소요', '미확인',      ('10일 소요', False)),
    ('샘플 대응',    '재고있음',   '불가능',      ('가능(요청시)', False)),
]

NO_QUOTE_C = DGRAY  # 견적없음/미확인 등 회색

for i, (label, ever, story, (idcom_txt, is_price)) in enumerate(rows_data):
    row = i + 1
    bg = WHITE if i % 2 == 0 else TABLE_EVEN

    # 구분 열 (레이블)
    cell_fmt(tbl3.cell(row, 0), label, size=11, bold=True, color=BLACK,
             fill=RGBColor(0xF0,0xF0,0xF0), bcolor='CCCCCC', bw=0.5, align=PP_ALIGN.CENTER)

    # 에버컴퍼니 열
    ever_color = NO_QUOTE_C if ever == '견적 없음' else BLACK
    cell_fmt(tbl3.cell(row, 1), ever, size=12, bold=False, color=ever_color,
             fill=bg, bcolor='CCCCCC', bw=0.5, align=PP_ALIGN.CENTER)

    # 이야기 열
    story_color = NO_QUOTE_C if story in ('견적 없음', '미확인', '불가능') else BLACK
    cell_fmt(tbl3.cell(row, 2), story, size=12, bold=False, color=story_color,
             fill=bg, bcolor='CCCCCC', bw=0.5, align=PP_ALIGN.CENTER)

    # 아이디컴 열 (최저가/우위)
    if is_price:
        cell_multirun(
            tbl3.cell(row, 3),
            [
                {'text': idcom_txt, 'size': 13, 'bold': True, 'color': RED},
                {'text': '  ▼ 최저', 'size': 9,  'bold': True, 'color': RED},
            ],
            fill=LRED, bcolor='CC0000', bw=0.8, align=PP_ALIGN.CENTER,
        )
    else:
        cell_fmt(tbl3.cell(row, 3), idcom_txt, size=12, bold=True, color=BLUE,
                 fill=LBLUE, bcolor='CC0000', bw=0.8, align=PP_ALIGN.CENTER)

slide_footer(slide3, '04 / 04')


# ══════════════════════════════════════════════
#  SLIDE 4 ─ 수량별 업체 단가 순위표
# ══════════════════════════════════════════════
slide4 = prs.slides.add_slide(blank)
slide_header(slide4, '수량별 업체 단가 순위표', '04 / 04')

# 3개 시나리오 패널
scenarios = [
    {
        'qty':   '1,000개 발주 시',
        'ever':  '5,500,000원',
        'story': '5,200,000원',
        'story_note': None,
        'idcom': '2,900,000원',
        'save':  '2,600,000원',
    },
    {
        'qty':   '1,500개 발주 시',
        'ever':  '6,885,000원',
        'story': '진행 불가',
        'story_note': None,
        'idcom': '3,900,000원',
        'save':  '2,985,000원',
    },
    {
        'qty':   '2,000개 발주 시',
        'ever':  '7,700,000원',
        'story': '9,400,000원',
        'story_note': None,
        'idcom': '4,600,000원',
        'save':  '3,100,000원',
    },
]

PANEL_Y  = Inches(1.1)
PANEL_H  = Inches(6.15)
PANEL_W  = Inches(4.1)
PANEL_GAP = Inches(0.17)

for pi, sc in enumerate(scenarios):
    PX = Inches(0.37) + pi * (PANEL_W + PANEL_GAP)

    # 패널 외곽 테두리 (미세한 그레이)
    rect(slide4, PX, PANEL_Y, PANEL_W, PANEL_H, fill=WHITE, line=MGRAY, lw=0.7)

    # 패널 헤더 (블루)
    HDR_H = Inches(0.68)
    rect(slide4, PX, PANEL_Y, PANEL_W, HDR_H, fill=BLUE)
    txt(slide4, sc['qty'], PX, PANEL_Y + Inches(0.08), PANEL_W, HDR_H - Inches(0.1),
        size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # 각 업체 행
    ITEM_Y0 = PANEL_Y + HDR_H + Inches(0.2)
    ITEM_H  = Inches(1.2)
    ITEM_GAP = Inches(0.1)

    def panel_item(slide, px, iy, label, amount, highlight=None, strikethrough=False, note=None):
        BG = LRED if highlight == 'red' else (LBLUE if highlight == 'blue' else LGRAY)
        BORDER = RED if highlight == 'red' else (BLUE if highlight == 'blue' else MGRAY)
        rect(slide, px + Inches(0.12), iy, PANEL_W - Inches(0.24), ITEM_H,
             fill=BG, line=BORDER, lw=0.7)
        # 레이블
        txt(slide, label,
            px + Inches(0.22), iy + Inches(0.1), PANEL_W - Inches(0.44), Inches(0.35),
            size=9, bold=False, color=DGRAY if highlight != 'red' else RED, align=PP_ALIGN.LEFT)
        # 금액
        amt_color = RED if highlight == 'red' else (BLUE if highlight == 'blue' else BLACK)
        amt_size  = 18 if highlight == 'red' else 16
        amt_bold  = True
        if strikethrough:
            amt_color = DGRAY; amt_size = 14; amt_bold = False
        txt(slide, amount,
            px + Inches(0.22), iy + Inches(0.46), PANEL_W - Inches(0.44), Inches(0.65),
            size=amt_size, bold=amt_bold, color=amt_color, align=PP_ALIGN.CENTER)
        if note:
            txt(slide, note, px + Inches(0.22), iy + Inches(0.85), PANEL_W - Inches(0.44), Inches(0.28),
                size=8, color=DGRAY, align=PP_ALIGN.CENTER)

    # 에버컴퍼니 (기존 / 비교 기준)
    panel_item(slide4, PX, ITEM_Y0, '에버컴퍼니 (기존)', sc['ever'])

    # 이야기
    is_unavail = sc['story'] == '진행 불가'
    panel_item(slide4, PX, ITEM_Y0 + ITEM_H + ITEM_GAP,
               '이야기', sc['story'],
               strikethrough=is_unavail, note=sc['story_note'])

    # 아이디컴 (최저가 - 레드)
    panel_item(slide4, PX, ITEM_Y0 + 2*(ITEM_H + ITEM_GAP),
               '아이디컴  ▼ 최저', sc['idcom'], highlight='red')

    # 구분선
    SEP_Y = ITEM_Y0 + 3*(ITEM_H + ITEM_GAP) + Inches(0.05)
    rect(slide4, PX + Inches(0.12), SEP_Y, PANEL_W - Inches(0.24), Inches(0.03), fill=MGRAY)

    # 기존 대비 절감 예상 (블루)
    SAVE_Y = SEP_Y + Inches(0.1)
    rect(slide4, PX + Inches(0.12), SAVE_Y, PANEL_W - Inches(0.24), Inches(0.85),
         fill=LBLUE, line=BLUE, lw=0.7)
    txt(slide4, '기존 대비 절감 예상', PX + Inches(0.22), SAVE_Y + Inches(0.06),
        PANEL_W - Inches(0.44), Inches(0.3), size=9, color=BLUE, align=PP_ALIGN.CENTER)
    txt(slide4, sc['save'], PX + Inches(0.22), SAVE_Y + Inches(0.35),
        PANEL_W - Inches(0.44), Inches(0.45), size=17, bold=True, color=BLUE, align=PP_ALIGN.CENTER)

slide_footer(slide4, '04 / 04')


# ─────────── 저장 ───────────
prs.save(OUT)
print(f'저장 완료: {OUT}')
