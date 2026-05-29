"""
메모패드 제작업체비교 CEO 보고서 — Kanban 스타일 v3
씨젠의료재단 CI: #CC0000, 맑은 고딕, 화이트 배경
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import os

# ── 팔레트 (CI 준수) ─────────────────────────────────────
RED    = RGBColor(0xCC, 0x00, 0x00)
BLUE   = RGBColor(0x1A, 0x56, 0xAB)
BLACK  = RGBColor(0x12, 0x12, 0x12)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
COL_BG = RGBColor(0xF4, 0xF4, 0xF4)   # 칸반 컬럼 배경
CARD_BORDER = RGBColor(0xD8, 0xD8, 0xD8)  # 카드 테두리
HDR_DARK = RGBColor(0x1A, 0x1A, 0x1A)    # 컬럼 헤더 배경
GRAY2  = RGBColor(0xC0, 0xC0, 0xC0)      # 구분선
GRAY3  = RGBColor(0x88, 0x88, 0x88)      # 보조 텍스트
GRAY4  = RGBColor(0x44, 0x44, 0x44)      # 서브 헤딩
LRED   = RGBColor(0xFF, 0xF4, 0xF4)      # 아이디컴 카드 배경
LBLUE  = RGBColor(0xEE, 0xF4, 0xFF)      # 절감 카드 배경

FONT = 'Malgun Gothic'
LOGO = r'c:\Users\최다빈\Desktop\PPT\seegene_ci.png'
OUT  = r'c:\Users\최다빈\Desktop\메모패드_제작업체비교_최종.pptx'

W, H = Inches(13.333), Inches(7.5)


# ── 기본 도형 ─────────────────────────────────────────────

def rect(slide, x, y, w, h, fill=None, border=None, bw=0.6):
    s = slide.shapes.add_shape(1, x, y, w, h)
    if fill:   s.fill.solid(); s.fill.fore_color.rgb = fill
    else:      s.fill.background()
    if border: s.line.color.rgb = border; s.line.width = Pt(bw)
    else:      s.line.fill.background()
    return s


def rnd_rect(slide, x, y, w, h, fill=None, border=None, bw=0.6):
    """둥근 모서리 카드 (shape type 5)"""
    s = slide.shapes.add_shape(5, x, y, w, h)
    # 둥근 정도 최소화 (adj 값을 5000으로 줄임 → 약 5%)
    try:
        sp = s._element
        ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        pg = sp.find(f'.//{{{ns}}}prstGeom')
        if pg is not None:
            al = pg.find(f'{{{ns}}}avLst')
            if al is None:
                al = etree.SubElement(pg, f'{{{ns}}}avLst')
            for ch in list(al): al.remove(ch)
            gd = etree.SubElement(al, f'{{{ns}}}gd')
            gd.set('name', 'adj'); gd.set('fmla', 'val 3000')
    except Exception:
        pass
    if fill:   s.fill.solid(); s.fill.fore_color.rgb = fill
    else:      s.fill.background()
    if border: s.line.color.rgb = border; s.line.width = Pt(bw)
    else:      s.line.fill.background()
    return s


# ── 텍스트 헬퍼 ───────────────────────────────────────────

def tb(slide, x, y, w, h, lines, wrap=True):
    """lines = list of dict(text, size, bold, color, align, space_before, space_after, runs)"""
    obj = slide.shapes.add_textbox(x, y, w, h)
    tf  = obj.text_frame
    tf.word_wrap = wrap
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ln.get('align', PP_ALIGN.LEFT)
        if ln.get('space_before'): p.space_before = Pt(ln['space_before'])
        if ln.get('space_after'):  p.space_after  = Pt(ln['space_after'])
        if ln.get('line_space'):   p.line_spacing  = Pt(ln['line_space'])
        if ln.get('runs'):
            for rn in ln['runs']:
                r = p.add_run(); r.text = rn['text']
                r.font.name = FONT; r.font.size = Pt(rn.get('size', 11))
                r.font.bold = rn.get('bold', False)
                r.font.color.rgb = rn.get('color', BLACK)
        else:
            r = p.add_run(); r.text = ln.get('text', '')
            r.font.name  = FONT;  r.font.size  = Pt(ln.get('size', 11))
            r.font.bold  = ln.get('bold', False)
            r.font.color.rgb = ln.get('color', BLACK)
    return obj


def t(slide, text, x, y, w, h, size=11, bold=False, color=BLACK,
      align=PP_ALIGN.LEFT, italic=False):
    return tb(slide, x, y, w, h, [
        {'text': text, 'size': size, 'bold': bold, 'color': color, 'align': align}])


# ── 테이블 셀 포맷 ────────────────────────────────────────

def cell_border(cell, sides, hex_c, w_pt):
    tc   = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None: tcPr = etree.SubElement(tc, qn('a:tcPr'))
    w_str = str(int(w_pt * 12700))
    side_tag = {'T':'lnT','B':'lnB','L':'lnL','R':'lnR'}
    for s in sides:
        tag = side_tag[s]
        el  = tcPr.find(qn(f'a:{tag}'))
        if el is None: el = etree.SubElement(tcPr, qn(f'a:{tag}'))
        el.set('w', w_str); el.set('cap','flat'); el.set('cmpd','sng')
        for ch in list(el): el.remove(ch)
        sf = etree.SubElement(el, qn('a:solidFill'))
        sr = etree.SubElement(sf, qn('a:srgbClr')); sr.set('val', hex_c)


def no_line(cell, sides):
    tc   = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None: tcPr = etree.SubElement(tc, qn('a:tcPr'))
    side_tag = {'T':'lnT','B':'lnB','L':'lnL','R':'lnR'}
    for s in sides:
        el = tcPr.find(qn(f'a:{side_tag[s]}'))
        if el is None: el = etree.SubElement(tcPr, qn(f'a:{side_tag[s]}'))
        for ch in list(el): el.remove(ch)
        etree.SubElement(el, qn('a:noFill'))


def fmt_cell(cell, text='', size=11, bold=False, color=BLACK, fill=None,
             align=PP_ALIGN.CENTER, border_hex='D0D0D0', bw=0.5,
             mt=Pt(4), ml=Pt(8)):
    if fill:  cell.fill.solid(); cell.fill.fore_color.rgb = fill
    else:     cell.fill.background()
    tf = cell.text_frame
    tf.margin_top = mt; tf.margin_bottom = mt
    tf.margin_left = ml; tf.margin_right = ml
    tf.word_wrap = True
    while len(tf.paragraphs) > 1:
        tf.paragraphs[-1]._p.getparent().remove(tf.paragraphs[-1]._p)
    p = tf.paragraphs[0]; p.alignment = align
    for r in list(p.runs): r._r.getparent().remove(r._r)
    r = p.add_run(); r.text = text
    r.font.name = FONT; r.font.size = Pt(size)
    r.font.bold = bold; r.font.color.rgb = color
    cell_border(cell, ['T','B','L','R'], border_hex, bw)


def fmt_cell_runs(cell, runs, fill=None, align=PP_ALIGN.CENTER,
                  border_hex='D0D0D0', bw=0.5, mt=Pt(4), ml=Pt(8)):
    if fill:  cell.fill.solid(); cell.fill.fore_color.rgb = fill
    else:     cell.fill.background()
    tf = cell.text_frame
    tf.margin_top = mt; tf.margin_bottom = mt
    tf.margin_left = ml; tf.margin_right = ml
    tf.word_wrap = True
    while len(tf.paragraphs) > 1:
        tf.paragraphs[-1]._p.getparent().remove(tf.paragraphs[-1]._p)
    p = tf.paragraphs[0]; p.alignment = align
    for r in list(p.runs): r._r.getparent().remove(r._r)
    for rn in runs:
        r = p.add_run(); r.text = rn['text']
        r.font.name = FONT; r.font.size = Pt(rn.get('size', 11))
        r.font.bold = rn.get('bold', False)
        r.font.color.rgb = rn.get('color', BLACK)
    cell_border(cell, ['T','B','L','R'], border_hex, bw)


# ── 공통 슬라이드 헤더 ────────────────────────────────────

def slide_header(slide, title, page):
    rect(slide, Inches(0), Inches(0), W, Inches(0.05), fill=RED)
    t(slide, title, Inches(0.45), Inches(0.12), Inches(10), Inches(0.62),
      size=22, bold=True, color=BLACK)
    t(slide, page, Inches(11.5), Inches(0.15), Inches(1.65), Inches(0.5),
      size=11, color=GRAY3, align=PP_ALIGN.RIGHT)
    rect(slide, Inches(0.45), Inches(0.75), W - Inches(0.55), Inches(0.02), fill=GRAY2)


# ── 결재란 ────────────────────────────────────────────────

def approval(slide, x, y, w, h):
    labels = ['담당', '총무팀장', '경영관리본부', '경영관리부문', '기획조정실', '행정원장', '이사장']
    tbl = slide.shapes.add_table(2, 7, x, y, w, h).table
    cw  = w // 7
    for c in tbl.columns: c.width = cw
    tbl.rows[0].height = int(h * 0.35)
    tbl.rows[1].height = int(h * 0.65)
    for i, lbl in enumerate(labels):
        fmt_cell(tbl.cell(0, i), lbl, size=7.5, bold=True, color=GRAY4,
                 fill=RGBColor(0xEE,0xEE,0xEE), border_hex='CCCCCC', bw=0.5, mt=Pt(2))
    for i in range(7):
        fmt_cell(tbl.cell(1, i), '', fill=WHITE, border_hex='BBBBBB', bw=0.5, mt=Pt(2))


# ── Kanban 카드 빌더 ──────────────────────────────────────

def kcard(slide, x, y, w, h, accent_c=None, bg=WHITE, border_c=CARD_BORDER):
    """카드 배경 + 좌측 컬러 액센트 바"""
    ACCENT_W = Inches(0.07)
    rnd_rect(slide, x, y, w, h, fill=bg, border=border_c, bw=0.7)
    if accent_c:
        rect(slide, x, y, ACCENT_W, h, fill=accent_c)
    return x + (ACCENT_W + Inches(0.1)) if accent_c else x + Inches(0.12)


def badge(slide, text, x, y, color, bg):
    """작은 뱃지"""
    BW, BH = Inches(0.9), Inches(0.28)
    rnd_rect(slide, x, y, BW, BH, fill=bg, border=color, bw=0.5)
    t(slide, text, x, y, BW, BH, size=8, bold=True, color=color, align=PP_ALIGN.CENTER)


def col_header(slide, x, y, w, h, title, subtitle='', bg=HDR_DARK, accent=RED):
    """칸반 컬럼 헤더"""
    rnd_rect(slide, x, y, w, h, fill=bg)
    rect(slide, x, y, w, Inches(0.045), fill=accent)
    t(slide, title, x + Inches(0.18), y + Inches(0.08),
      w - Inches(0.25), Inches(0.4), size=13, bold=True, color=WHITE)
    if subtitle:
        t(slide, subtitle, x + Inches(0.18), y + Inches(0.48),
          w - Inches(0.25), Inches(0.25), size=9, color=RGBColor(0xAA,0xAA,0xAA))


# ══════════════════════════════════════════════════════════
#  SLIDE 1  표지
# ══════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width, prs.slide_height = W, H
blank = prs.slide_layouts[6]

s1 = prs.slides.add_slide(blank)

# 상단 레드 라인
rect(s1, Inches(0), Inches(0), W, Inches(0.045), fill=RED)

# 로고
try:
    s1.shapes.add_picture(LOGO, Inches(0.48), Inches(0.22), Inches(2.8), Inches(0.9))
except:
    t(s1, '씨젠의료재단', Inches(0.48), Inches(0.22), Inches(3), Inches(0.7),
      size=18, bold=True, color=RED)

# 결재란
approval(s1, Inches(7.5), Inches(0.2), Inches(5.52), Inches(1.08))

# 상단 회색 구분선
rect(s1, Inches(0.45), Inches(1.4), W - Inches(0.55), Inches(0.018), fill=GRAY2)

# '검토 보고' 뱃지
rnd_rect(s1, Inches(0.48), Inches(1.7), Inches(1.6), Inches(0.38), fill=RED)
t(s1, '검 토 보 고', Inches(0.48), Inches(1.72), Inches(1.6), Inches(0.35),
  size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# 메인 타이틀 카드
kcard(s1, Inches(0.38), Inches(2.3), Inches(11.5), Inches(2.85),
      accent_c=RED, bg=RGBColor(0xF9,0xF9,0xF9), border_c=CARD_BORDER)

tb(s1, Inches(0.7), Inches(2.5), Inches(11), Inches(2.6), [
    {'text': '메모패드',
     'size': 12, 'bold': False, 'color': GRAY3, 'space_after': 10},
    {'text': '공급 단가 검증 및 업체별',
     'size': 36, 'bold': True, 'color': BLACK, 'space_after': 2},
    {'text': '견적 비교 검토',
     'size': 36, 'bold': True, 'color': BLACK},
])

# 하단 정보 영역
rect(s1, Inches(0), Inches(6.22), W, Inches(0.025), fill=GRAY2)

tb(s1, Inches(0.48), Inches(6.38), Inches(7), Inches(0.95), [
    {'text': '경영관리본부  총무팀',
     'size': 10, 'color': GRAY3, 'space_after': 5},
    {'text': '최다빈 대리',
     'size': 12, 'bold': True, 'color': GRAY4},
])

t(s1, '2026. 05. 27.', Inches(9.8), Inches(6.5), Inches(3.3), Inches(0.45),
  size=11, color=GRAY3, align=PP_ALIGN.RIGHT)
t(s1, '01 / 04', Inches(10.9), Inches(6.95), Inches(2.2), Inches(0.38),
  size=9, color=GRAY3, align=PP_ALIGN.RIGHT)

rect(s1, Inches(0), H - Inches(0.04), W, Inches(0.04), fill=RED)


# ══════════════════════════════════════════════════════════
#  SLIDE 2  검토 배경 및 목적  — 2컬럼 Kanban
# ══════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
slide_header(s2, '검토 배경 및 목적', '02 / 04')

BOARD_Y = Inches(0.88)
BOARD_H = Inches(4.38)
LW  = Inches(6.2)
RW  = Inches(5.85)
LX  = Inches(0.38)
RX  = LX + LW + Inches(0.22)

# ── 왼쪽 컬럼: 검토 목적 ──
rect(s2, LX, BOARD_Y, LW, BOARD_H, fill=COL_BG)
col_header(s2, LX, BOARD_Y, LW, Inches(0.72), '검토 목적')

objs = [
    ('01', '장기 거래처 단가의 적정성 정밀 검증',
     '장기 거래처(에버컴퍼니)의 공급 단가가 시장 가격 대비 적정 수준인지\n정밀 검증할 필요성 제기'),
    ('02', '신규 업체 견적의 객관적 대조 분석',
     '신규 유사 업체 견적·공급 조건을 동일 기준으로 대조,\n재단의 최적 예산 구조를 객관적으로 파악'),
    ('03', '관할 병·의원 영업 활용',
     '관할 병·의원 네트워크 유지 활동 시 현장에서 상시 배포되는 핵심 판촉물'),
]
OBJ_CH = Inches(1.14)
OBJ_GAP = Inches(0.1)

for i, (num, title, body) in enumerate(objs):
    cy = BOARD_Y + Inches(0.82) + i * (OBJ_CH + OBJ_GAP)
    cx = LX + Inches(0.14)
    cw = LW - Inches(0.28)

    tx = kcard(s2, cx, cy, cw, OBJ_CH, accent_c=RED, bg=WHITE)

    # 번호 뱃지
    rnd_rect(s2, tx, cy + Inches(0.12), Inches(0.38), Inches(0.38),
             fill=RED)
    t(s2, num, tx, cy + Inches(0.12), Inches(0.38), Inches(0.38),
      size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    t(s2, title, tx + Inches(0.48), cy + Inches(0.1),
      cw - Inches(0.62), Inches(0.38), size=11, bold=True, color=BLACK)
    t(s2, body, tx + Inches(0.48), cy + Inches(0.5),
      cw - Inches(0.62), Inches(0.62), size=9.5, color=GRAY3)

# ── 오른쪽 컬럼: 소요처 및 용도 ──
rect(s2, RX, BOARD_Y, RW, BOARD_H, fill=COL_BG)
col_header(s2, RX, BOARD_Y, RW, Inches(0.72), '메모패드 주요 소요처 및 용도')

uses = [
    ('교육 프로그램 및 방문객 지급',
     '센터 주관 교육 프로그램, 내방객 방문 및\n학생 초청 행사 지급용'),
    ('재단 홍보 및 외빈 응대 기념품',
     '사내 직무 교육·대학생 견학·주요 외빈 방문 시\n재단 홍보 기념품으로 비치.'),
]
USE_CH = Inches(1.73)

for i, (title, body) in enumerate(uses):
    cy = BOARD_Y + Inches(0.82) + i * (USE_CH + OBJ_GAP)
    cx = RX + Inches(0.14)
    cw = RW - Inches(0.28)
    tx = kcard(s2, cx, cy, cw, USE_CH, accent_c=RED, bg=WHITE)
    t(s2, title, tx, cy + Inches(0.14), cw - Inches(0.2), Inches(0.38),
      size=11, bold=True, color=BLACK)
    t(s2, body, tx, cy + Inches(0.58), cw - Inches(0.2), Inches(1.08),
      size=10, color=GRAY3)

# ── 의견 카드 ──
OPN_Y  = BOARD_Y + BOARD_H + Inches(0.2)
OPN_H  = Inches(1.72)

tx2 = kcard(s2, Inches(0.38), OPN_Y, W - Inches(0.5), OPN_H,
            accent_c=RED, bg=WHITE, border_c=CARD_BORDER)

rnd_rect(s2, tx2, OPN_Y + Inches(0.14), Inches(0.8), Inches(0.32), fill=RED)
t(s2, '의견', tx2, OPN_Y + Inches(0.14), Inches(0.8), Inches(0.32),
  size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

tb(s2, tx2 + Inches(0.92), OPN_Y + Inches(0.12), W - tx2 - Inches(1.1), Inches(0.55), [
    {'runs': [
        {'text': '견적이 제출된 전 수량 구간에서 ', 'size': 10.5, 'color': BLACK},
        {'text': '아이디컴이 최저가를 형성', 'size': 10.5, 'bold': True, 'color': RED},
        {'text': '하였으며, 제작 기간 단축 및 샘플 대응 가능 측면에서도 차별점이 확인됨.', 'size': 10.5, 'color': BLACK},
    ], 'align': PP_ALIGN.LEFT}
])
tb(s2, tx2 + Inches(0.92), OPN_Y + Inches(0.72), W - tx2 - Inches(1.1), Inches(0.9), [
    {'runs': [
        {'text': '기존 대비 약 ', 'size': 10.5, 'color': BLACK},
        {'text': '40.3% ~ 47.3%', 'size': 12, 'bold': True, 'color': RED},
        {'text': ' 예산이 절감되는 효과가 있으나, 실제 도입 전 사전 샘플을 통한 품질 검증 과정이 필요할 것으로 보임.', 'size': 10.5, 'color': BLACK},
    ], 'align': PP_ALIGN.LEFT}
])

rect(s2, Inches(0), H - Inches(0.04), W, Inches(0.04), fill=RED)
t(s2, '02 / 04', Inches(11.2), H - Inches(0.25), Inches(2.0), Inches(0.22),
  size=8, color=GRAY2, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════
#  SLIDE 3  수량별 업체 비교표 — Kanban Grid
# ══════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
slide_header(s3, '수량별 업체 비교표', '04 / 04')

ROWS = 7; COLS = 4
TBL_X = Inches(0.35); TBL_Y = Inches(0.92)
TBL_W = W - Inches(0.45); TBL_H = H - TBL_Y - Inches(0.18)

tbl3 = s3.shapes.add_table(ROWS, COLS, TBL_X, TBL_Y, TBL_W, TBL_H).table

tbl3.columns[0].width = Inches(2.1)
tbl3.columns[1].width = Inches(3.1)
tbl3.columns[2].width = Inches(3.1)
tbl3.columns[3].width = Inches(4.47)

tbl3.rows[0].height = Inches(0.78)
for i in range(1, ROWS):
    tbl3.rows[i].height = Inches(0.98)

# 헤더
HDR_DATA = [
    ('구분',              RGBColor(0x2A,0x2A,0x2A)),
    ('에버컴퍼니 (기존)', RGBColor(0x3A,0x3A,0x3A)),
    ('이야기',            RGBColor(0x3A,0x3A,0x3A)),
    ('아이디컴',          RED),
]
for j, (htext, hfill) in enumerate(HDR_DATA):
    c = tbl3.cell(0, j)
    fmt_cell(c, htext, size=13, bold=True, color=WHITE, fill=hfill,
             border_hex='111111', bw=0.5, mt=Pt(5))

# 데이터 행
NA = {'견적 없음', '미확인', '불가능'}
rows_data = [
    ('500개 구간',   '견적 없음',  '견적 없음',  '3,300원',  True,  False),
    ('1,000개 구간', '5,500원',    '5,200원',    '2,900원',  True,  False),
    ('1,500개 구간', '4,590원',    '견적 없음',  '2,600원',  True,  False),
    ('2,000개 구간', '3,850원',    '4,700원',    '2,300원',  True,  False),
    ('제작 기간',    '4~5주 소요', '미확인',      '10일 소요', False, True),
    ('샘플 대응',    '재고있음',   '불가능',      '가능(요청시)', False, True),
]

EVEN_BG = RGBColor(0xFA,0xFA,0xFA)

for ri, (label, ever, story, idcom, is_price, is_adv) in enumerate(rows_data):
    row  = ri + 1
    bg   = WHITE if ri % 2 == 0 else EVEN_BG

    # 구분 열
    fmt_cell(tbl3.cell(row, 0), label, size=11, bold=True, color=GRAY4,
             fill=RGBColor(0xF0,0xF0,0xF0), border_hex='C8C8C8', bw=0.5, mt=Pt(5))

    # 에버컴퍼니
    fmt_cell(tbl3.cell(row, 1), ever, size=13, bold=False,
             color=GRAY3 if ever in NA else BLACK,
             fill=bg, border_hex='C8C8C8', bw=0.5, mt=Pt(5))

    # 이야기
    fmt_cell(tbl3.cell(row, 2), story, size=13, bold=False,
             color=GRAY3 if story in NA else BLACK,
             fill=bg, border_hex='C8C8C8', bw=0.5, mt=Pt(5))

    # 아이디컴
    if is_price:
        fmt_cell_runs(tbl3.cell(row, 3), [
            {'text': idcom, 'size': 15, 'bold': True, 'color': RED},
            {'text': '   ▼ 최저', 'size': 9, 'bold': True, 'color': RED},
        ], fill=LRED, border_hex='CC0000', bw=0.9, mt=Pt(5))
    else:
        fmt_cell(tbl3.cell(row, 3), idcom, size=13, bold=True, color=BLUE,
                 fill=LBLUE, border_hex='1A56AB', bw=0.9, mt=Pt(5))

rect(s3, Inches(0), H - Inches(0.04), W, Inches(0.04), fill=RED)
t(s3, '04 / 04', Inches(11.2), H - Inches(0.25), Inches(2.0), Inches(0.22),
  size=8, color=GRAY2, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════
#  SLIDE 4  수량별 업체 단가 순위표 — 순수 3컬럼 Kanban
# ══════════════════════════════════════════════════════════
s4 = prs.slides.add_slide(blank)
slide_header(s4, '수량별 업체 단가 순위표', '04 / 04')

BOARD_Y4 = Inches(0.9)
BOARD_H4 = H - BOARD_Y4 - Inches(0.2)

N_COL = 3
GAP   = Inches(0.2)
AVAILABLE_W = W - Inches(0.38) * 2 - GAP * (N_COL - 1)
COL_W = AVAILABLE_W / N_COL

scenarios = [
    {'qty':  '1,000개 발주 시',
     'ever': '5,500,000원', 'story': '5,200,000원', 'story_na': False,
     'idcom':'2,900,000원', 'save':  '2,600,000원'},
    {'qty':  '1,500개 발주 시',
     'ever': '6,885,000원', 'story': '진행 불가',   'story_na': True,
     'idcom':'3,900,000원', 'save':  '2,985,000원'},
    {'qty':  '2,000개 발주 시',
     'ever': '7,700,000원', 'story': '9,400,000원', 'story_na': False,
     'idcom':'4,600,000원', 'save':  '3,100,000원'},
]

CARD_H   = Inches(1.3)
CARD_GAP = Inches(0.13)
HDR_H    = Inches(0.8)

for pi, sc in enumerate(scenarios):
    PX = Inches(0.38) + pi * (COL_W + GAP)

    # 컬럼 배경
    rect(s4, PX, BOARD_Y4, COL_W, BOARD_H4, fill=COL_BG)

    # 컬럼 헤더
    col_header(s4, PX, BOARD_Y4, COL_W, HDR_H, sc['qty'])

    CY = BOARD_Y4 + HDR_H + CARD_GAP

    # ─ 에버컴퍼니 카드 ─
    cx = PX + Inches(0.14)
    cw = COL_W - Inches(0.28)
    kcard(s4, cx, CY, cw, CARD_H, accent_c=GRAY2, bg=WHITE, border_c=CARD_BORDER)
    t(s4, '에버컴퍼니 (기존)', cx + Inches(0.17), CY + Inches(0.1),
      cw - Inches(0.22), Inches(0.3), size=9.5, color=GRAY3)
    t(s4, sc['ever'], cx + Inches(0.17), CY + Inches(0.42),
      cw - Inches(0.22), Inches(0.78), size=18, bold=True, color=BLACK,
      align=PP_ALIGN.CENTER)

    CY += CARD_H + CARD_GAP

    # ─ 이야기 카드 ─
    kcard(s4, cx, CY, cw, CARD_H,
          accent_c=GRAY2, bg=WHITE, border_c=CARD_BORDER)
    t(s4, '이야기', cx + Inches(0.17), CY + Inches(0.1),
      cw - Inches(0.22), Inches(0.3), size=9.5, color=GRAY3)
    story_color = GRAY3 if sc['story_na'] else BLACK
    t(s4, sc['story'], cx + Inches(0.17), CY + Inches(0.42),
      cw - Inches(0.22), Inches(0.78),
      size=14 if sc['story_na'] else 18,
      bold=not sc['story_na'], color=story_color, align=PP_ALIGN.CENTER)

    CY += CARD_H + CARD_GAP

    # ─ 아이디컴 카드 (레드 강조) ─
    kcard(s4, cx, CY, cw, CARD_H, accent_c=RED, bg=LRED,
          border_c=RED)

    # '▼ 최저' 뱃지
    BW, BH = Inches(0.85), Inches(0.28)
    rnd_rect(s4, cx + cw - BW - Inches(0.1), CY + Inches(0.08), BW, BH,
             fill=RED)
    t(s4, '▼ 최저', cx + cw - BW - Inches(0.1), CY + Inches(0.08),
      BW, BH, size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    t(s4, '아이디컴', cx + Inches(0.17), CY + Inches(0.1),
      cw * 0.55, Inches(0.3), size=9.5, color=RED)
    t(s4, sc['idcom'], cx + Inches(0.17), CY + Inches(0.42),
      cw - Inches(0.22), Inches(0.78), size=20, bold=True, color=RED,
      align=PP_ALIGN.CENTER)

    CY += CARD_H + CARD_GAP + Inches(0.06)

    # 구분선
    rect(s4, cx, CY - Inches(0.04), cw, Inches(0.018), fill=GRAY2)

    # ─ 절감 예상 카드 (블루 강조) ─
    SAVE_H = BOARD_Y4 + BOARD_H4 - CY - Inches(0.14)
    kcard(s4, cx, CY, cw, SAVE_H, accent_c=BLUE, bg=LBLUE,
          border_c=BLUE)
    t(s4, '기존 대비 절감 예상', cx + Inches(0.17), CY + Inches(0.1),
      cw - Inches(0.22), Inches(0.3), size=9, color=BLUE)
    t(s4, sc['save'], cx + Inches(0.17), CY + Inches(0.38),
      cw - Inches(0.22), SAVE_H - Inches(0.5),
      size=18, bold=True, color=BLUE, align=PP_ALIGN.CENTER)

rect(s4, Inches(0), H - Inches(0.04), W, Inches(0.04), fill=RED)
t(s4, '04 / 04', Inches(11.2), H - Inches(0.25), Inches(2.0), Inches(0.22),
  size=8, color=GRAY2, align=PP_ALIGN.RIGHT)


# ── 저장 ─────────────────────────────────────────────────
prs.save(OUT)
print(f'완료: {OUT}')
