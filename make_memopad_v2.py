"""
메모패드 제작업체비교 CEO 보고서 — 미니멀 라인형 (v2)
씨젠의료재단 CI: #CC0000 레드, 맑은 고딕, 화이트 배경
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import os

# ── 팔레트 ──────────────────────────────────────────────
RED    = RGBColor(0xCC, 0x00, 0x00)
BLUE   = RGBColor(0x1A, 0x56, 0xAB)
BLACK  = RGBColor(0x12, 0x12, 0x12)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GRAY1  = RGBColor(0xF3, 0xF3, 0xF3)   # 아주 연한 회색 (테이블 줄무늬)
GRAY2  = RGBColor(0xC8, 0xC8, 0xC8)   # 구분선
GRAY3  = RGBColor(0x88, 0x88, 0x88)   # 보조 텍스트
GRAY4  = RGBColor(0x44, 0x44, 0x44)   # 서브 헤딩
DKGRAY = RGBColor(0x22, 0x22, 0x22)   # 테이블 헤더 배경

FONT = 'Malgun Gothic'
LOGO = r'c:\Users\최다빈\Desktop\PPT\seegene_ci.png'
OUT  = r'c:\Users\최다빈\Desktop\메모패드_제작업체비교_최종.pptx'

W, H = Inches(13.333), Inches(7.5)


# ── 핵심 헬퍼 ────────────────────────────────────────────

def line(slide, x, y, w, h, color):
    """얇은 구분선 (채움 사각형)"""
    s = slide.shapes.add_shape(1, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def box(slide, x, y, w, h, fill=None, line_c=None, lw=0.5):
    """일반 사각형"""
    s = slide.shapes.add_shape(1, x, y, w, h)
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line_c:
        s.line.color.rgb = line_c; s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    return s


def tb(slide, x, y, w, h, lines, wrap=True):
    """
    lines: list of dicts:
      text, size, bold, color, align, italic,
      space_before, space_after, runs (list of {text,size,bold,color})
    """
    tb_obj = slide.shapes.add_textbox(x, y, w, h)
    tf = tb_obj.text_frame
    tf.word_wrap = wrap
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ln.get('align', PP_ALIGN.LEFT)
        sb = ln.get('space_before', 0)
        sa = ln.get('space_after', 0)
        ls = ln.get('line_space', 0)
        if sb: p.space_before = Pt(sb)
        if sa: p.space_after  = Pt(sa)
        if ls: p.line_spacing  = Pt(ls)
        if ln.get('runs'):
            for rn in ln['runs']:
                r = p.add_run()
                r.text = rn['text']
                r.font.name  = FONT
                r.font.size  = Pt(rn.get('size', 11))
                r.font.bold  = rn.get('bold', False)
                r.font.color.rgb = rn.get('color', BLACK)
        else:
            r = p.add_run()
            r.text = ln.get('text', '')
            r.font.name   = FONT
            r.font.size   = Pt(ln.get('size', 11))
            r.font.bold   = ln.get('bold', False)
            r.font.italic = ln.get('italic', False)
            r.font.color.rgb = ln.get('color', BLACK)
    return tb_obj


def simple_tb(slide, text, x, y, w, h, size=11, bold=False, color=BLACK,
              align=PP_ALIGN.LEFT, italic=False):
    return tb(slide, x, y, w, h, [
        {'text': text, 'size': size, 'bold': bold, 'color': color,
         'align': align, 'italic': italic}
    ])


# ── 테이블 셀 포맷 ────────────────────────────────────────

def tc_border(cell, sides, hex_c, w_pt):
    tc   = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None: tcPr = etree.SubElement(tc, qn('a:tcPr'))
    w = str(int(w_pt * 12700))
    side_map = {'T':'lnT','B':'lnB','L':'lnL','R':'lnR'}
    for s in sides:
        tag = side_map[s]
        el = tcPr.find(qn(f'a:{tag}'))
        if el is None: el = etree.SubElement(tcPr, qn(f'a:{tag}'))
        el.set('w', w); el.set('cap','flat'); el.set('cmpd','sng')
        for ch in list(el): el.remove(ch)
        sf = etree.SubElement(el, qn('a:solidFill'))
        sr = etree.SubElement(sf, qn('a:srgbClr'))
        sr.set('val', hex_c)


def no_border(cell):
    """셀 테두리 완전 제거"""
    tc   = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None: tcPr = etree.SubElement(tc, qn('a:tcPr'))
    for tag in ['lnT','lnB','lnL','lnR']:
        el = tcPr.find(qn(f'a:{tag}'))
        if el is None: el = etree.SubElement(tcPr, qn(f'a:{tag}'))
        for ch in list(el): el.remove(ch)
        etree.SubElement(el, qn('a:noFill'))


def cf(cell, text='', size=10.5, bold=False, color=BLACK, fill=None,
       align=PP_ALIGN.CENTER, bc='C8C8C8', bw=0.5,
       top_b=True, bot_b=True, left_b=True, right_b=True,
       mt=Pt(3), ml=Pt(6)):
    """셀 포맷 + 테두리 세부 제어"""
    if fill:
        cell.fill.solid(); cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
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
    sides = []
    if top_b:   sides.append('T')
    if bot_b:   sides.append('B')
    if left_b:  sides.append('L')
    if right_b: sides.append('R')
    if sides: tc_border(cell, sides, bc, bw)
    # Remove borders not requested
    off_sides = [s for s in ['T','B','L','R'] if s not in sides]
    for s in off_sides:
        tc = cell._tc
        tcPr = tc.find(qn('a:tcPr'))
        if tcPr is None: tcPr = etree.SubElement(tc, qn('a:tcPr'))
        tag = {'T':'lnT','B':'lnB','L':'lnL','R':'lnR'}[s]
        el = tcPr.find(qn(f'a:{tag}'))
        if el is None: el = etree.SubElement(tcPr, qn(f'a:{tag}'))
        for ch in list(el): el.remove(ch)
        etree.SubElement(el, qn('a:noFill'))


def cf_runs(cell, runs, fill=None, align=PP_ALIGN.CENTER,
            bc='C8C8C8', bw=0.5, mt=Pt(3), ml=Pt(6)):
    """멀티런 셀 (색상이 다른 텍스트 조합)"""
    if fill:
        cell.fill.solid(); cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
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
        r.font.name = FONT; r.font.size = Pt(rn.get('size', 10.5))
        r.font.bold = rn.get('bold', False)
        r.font.color.rgb = rn.get('color', BLACK)
    tc_border(cell, ['T','B','L','R'], bc, bw)


# ── 결재란 ─────────────────────────────────────────────────

def approval_table(slide, x, y, w, h):
    labels = ['담당', '총무팀장', '경영관리본부', '경영관리부문', '기획조정실', '행정원장', '이사장']
    tbl = slide.shapes.add_table(2, 7, x, y, w, h).table
    cw = w // 7
    for c in tbl.columns: c.width = cw
    tbl.rows[0].height = int(h * 0.36)
    tbl.rows[1].height = int(h * 0.64)
    for i, lbl in enumerate(labels):
        cf(tbl.cell(0, i), lbl, size=7.5, bold=True, color=GRAY4,
           fill=GRAY1, bc='C8C8C8', bw=0.5, align=PP_ALIGN.CENTER, mt=Pt(2))
    for i in range(7):
        cf(tbl.cell(1, i), '', fill=WHITE, bc='C8C8C8', bw=0.5, mt=Pt(2))


# ══════════════════════════════════════════════════════════
#  SLIDE 1  표지
# ══════════════════════════════════════════════════════════
prs = Presentation()
prs.slide_width, prs.slide_height = W, H
blank = prs.slide_layouts[6]

s1 = prs.slides.add_slide(blank)

# 상단 로고
try:
    s1.shapes.add_picture(LOGO, Inches(0.5), Inches(0.32), Inches(2.7), Inches(0.82))
except:
    simple_tb(s1, '씨젠의료재단', Inches(0.5), Inches(0.32), Inches(3), Inches(0.7),
              size=17, bold=True, color=RED)

# 결재란 (우측 상단, 얇은 테두리)
approval_table(s1, Inches(7.5), Inches(0.25), Inches(5.5), Inches(1.05))

# 상단 레드 라인 (전체 폭)
line(s1, Inches(0), Inches(1.45), W, Inches(0.035), RED)

# '검토 보고' 레이블 (작은 레드 텍스트)
simple_tb(s1, '검 토 보 고', Inches(0.52), Inches(1.65), Inches(2.5), Inches(0.4),
          size=10.5, bold=True, color=RED)

# 메인 타이틀
tb(s1, Inches(0.52), Inches(2.2), Inches(11.5), Inches(2.8), [
    {'text': '메모패드',
     'size': 13, 'bold': False, 'color': GRAY3,
     'align': PP_ALIGN.LEFT, 'space_after': 10},
    {'text': '공급 단가 검증 및 업체별',
     'size': 38, 'bold': True, 'color': BLACK,
     'align': PP_ALIGN.LEFT, 'space_after': 0},
    {'text': '견적 비교 검토',
     'size': 38, 'bold': True, 'color': BLACK,
     'align': PP_ALIGN.LEFT},
])

# 타이틀 하단 레드 짧은 라인
line(s1, Inches(0.52), Inches(5.15), Inches(2.2), Inches(0.045), RED)

# 하단 구분선 (회색)
line(s1, Inches(0), Inches(6.25), W, Inches(0.02), GRAY2)

# 부서/담당자 (좌)
tb(s1, Inches(0.52), Inches(6.38), Inches(7), Inches(0.9), [
    {'text': '경영관리본부  총무팀',
     'size': 10, 'bold': False, 'color': GRAY3, 'space_after': 5},
    {'text': '최다빈 대리',
     'size': 12, 'bold': True, 'color': GRAY4},
])

# 날짜 (우)
simple_tb(s1, '2026. 05. 27.',
          Inches(9.8), Inches(6.45), Inches(3.3), Inches(0.5),
          size=11, color=GRAY3, align=PP_ALIGN.RIGHT)

# 페이지 번호
simple_tb(s1, '01 / 04',
          Inches(10.9), Inches(6.95), Inches(2.2), Inches(0.4),
          size=9, color=GRAY3, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════
#  SLIDE 2  검토 배경 및 목적
# ══════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)

# 상단 레드 라인
line(s2, Inches(0), Inches(0), W, Inches(0.04), RED)

# 슬라이드 제목 영역
simple_tb(s2, '검토 배경 및 목적',
          Inches(0.5), Inches(0.15), Inches(9), Inches(0.7),
          size=22, bold=True, color=BLACK)
simple_tb(s2, '02 / 04',
          Inches(11.5), Inches(0.18), Inches(1.6), Inches(0.55),
          size=11, color=GRAY3, align=PP_ALIGN.RIGHT)

# 제목 하단 회색 구분선
line(s2, Inches(0.5), Inches(0.88), W - Inches(0.6), Inches(0.018), GRAY2)

# ─ 좌측: 검토 목적 ──────────────────────────────
LX, RX = Inches(0.5), Inches(6.95)
TY = Inches(1.05)
SW = Inches(6.15)

simple_tb(s2, '검토 목적', LX, TY, Inches(3), Inches(0.42),
          size=12, bold=True, color=GRAY4)
line(s2, LX, TY + Inches(0.42), Inches(2.6), Inches(0.022), GRAY2)

objectives = [
    ('장기 거래처 단가의 적정성 정밀 검증',
     '장기 거래처(에버컴퍼니)의 공급 단가가 시장 가격 대비 적정 수준인지\n정밀 검증할 필요성 제기'),
    ('신규 업체 견적의 객관적 대조 분석',
     '신규 유사 업체 견적·공급 조건을 동일 기준으로 대조,\n재단의 최적 예산 구조를 객관적으로 파악'),
    ('관할 병·의원 영업 활용',
     '관할 병·의원 네트워크 유지 활동 시 현장에서 상시 배포되는 핵심 판촉물'),
]

OBJ_Y   = TY + Inches(0.55)
OBJ_GAP = Inches(0.15)
OBJ_H   = Inches(1.1)

for i, (title, body) in enumerate(objectives):
    oy = OBJ_Y + i * (OBJ_H + OBJ_GAP)
    # 좌측 레드 세로 강조선
    line(s2, LX, oy, Inches(0.05), OBJ_H, RED)
    # 번호
    simple_tb(s2, f'0{i+1}', LX + Inches(0.15), oy + Inches(0.05),
              Inches(0.4), Inches(0.38), size=11, bold=True, color=RED)
    # 제목
    simple_tb(s2, title, LX + Inches(0.55), oy + Inches(0.05),
              SW - Inches(0.6), Inches(0.38), size=11, bold=True, color=BLACK)
    # 내용
    simple_tb(s2, body, LX + Inches(0.55), oy + Inches(0.45),
              SW - Inches(0.6), Inches(0.6), size=9.5, color=GRAY3)

# ─ 우측: 주요 소요처 및 용도 ────────────────────
simple_tb(s2, '메모패드 주요 소요처 및 용도',
          RX, TY, Inches(6), Inches(0.42), size=12, bold=True, color=GRAY4)
line(s2, RX, TY + Inches(0.42), Inches(5.9), Inches(0.022), GRAY2)

uses = [
    ('교육 프로그램 및 방문객 지급',
     '센터 주관 교육 프로그램, 내방객 방문 및 학생 초청 행사 지급용'),
    ('재단 홍보 및 외빈 응대 기념품',
     '사내 직무 교육·대학생 견학·주요 외빈 방문 시 재단 홍보 기념품으로 비치.'),
]

USE_H   = Inches(1.6)
USE_W   = Inches(5.9)

for i, (title, body) in enumerate(uses):
    uy = OBJ_Y + i * (USE_H + OBJ_GAP)
    # 레드 세로 강조선
    line(s2, RX, uy, Inches(0.05), USE_H, RED)
    # 배경 (극히 연한 회색)
    box(s2, RX + Inches(0.05), uy, USE_W - Inches(0.05), USE_H, fill=GRAY1)
    simple_tb(s2, title, RX + Inches(0.2), uy + Inches(0.15),
              USE_W - Inches(0.3), Inches(0.38), size=11, bold=True, color=BLACK)
    simple_tb(s2, body, RX + Inches(0.2), uy + Inches(0.6),
              USE_W - Inches(0.3), Inches(0.9), size=10, color=GRAY3)

# ─ 의견 박스 ─────────────────────────────────────
OPN_Y = Inches(5.42)
OPN_H = Inches(1.85)
# 좌측 레드 강조선
line(s2, Inches(0.5), OPN_Y, Inches(0.055), OPN_H, RED)
simple_tb(s2, '의견', Inches(0.68), OPN_Y + Inches(0.06),
          Inches(1.2), Inches(0.35), size=10.5, bold=True, color=RED)
tb(s2, Inches(0.68), OPN_Y + Inches(0.44), W - Inches(0.9), Inches(1.35), [
    {
        'runs': [
            {'text': '견적이 제출된 전 수량 구간에서 ', 'size': 10.5, 'color': BLACK},
            {'text': '아이디컴이 최저가를 형성', 'size': 10.5, 'bold': True, 'color': RED},
            {'text': '하였으며, 제작 기간 단축 및 샘플 대응 가능 측면에서도 차별점이 확인됨.', 'size': 10.5, 'color': BLACK},
        ],
        'align': PP_ALIGN.LEFT, 'space_after': 6,
    },
    {
        'runs': [
            {'text': '기존 대비 약 ', 'size': 10.5, 'color': BLACK},
            {'text': '40.3% ~ 47.3%', 'size': 11.5, 'bold': True, 'color': RED},
            {'text': ' 예산이 절감되는 효과가 있으나, 실제 도입 전 사전 샘플을 통한 품질 검증 과정이 필요할 것으로 보임.', 'size': 10.5, 'color': BLACK},
        ],
        'align': PP_ALIGN.LEFT,
    },
])

# 하단 레드 라인
line(s2, Inches(0), H - Inches(0.03), W, Inches(0.03), RED)
simple_tb(s2, '02 / 04', Inches(11.2), H - Inches(0.25), Inches(1.9), Inches(0.22),
          size=8, color=GRAY2, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════
#  SLIDE 3  수량별 업체 비교표
# ══════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)

line(s3, Inches(0), Inches(0), W, Inches(0.04), RED)
simple_tb(s3, '수량별 업체 비교표',
          Inches(0.5), Inches(0.15), Inches(9), Inches(0.7),
          size=22, bold=True, color=BLACK)
simple_tb(s3, '04 / 04',
          Inches(11.5), Inches(0.18), Inches(1.6), Inches(0.55),
          size=11, color=GRAY3, align=PP_ALIGN.RIGHT)
line(s3, Inches(0.5), Inches(0.88), W - Inches(0.6), Inches(0.018), GRAY2)

# 테이블
ROWS, COLS = 7, 4
TBL_X = Inches(0.42)
TBL_Y = Inches(1.08)
TBL_W = W - Inches(0.55)
TBL_H = Inches(6.18)

tbl3 = s3.shapes.add_table(ROWS, COLS, TBL_X, TBL_Y, TBL_W, TBL_H).table

# 열 너비
tbl3.columns[0].width = Inches(2.3)
tbl3.columns[1].width = Inches(2.95)
tbl3.columns[2].width = Inches(2.95)
tbl3.columns[3].width = Inches(4.45)

# 행 높이
tbl3.rows[0].height = Inches(0.72)
for i in range(1, ROWS):
    tbl3.rows[i].height = Inches(0.91)

# 헤더 행
HDR_FILLS = [DKGRAY, RGBColor(0x3A,0x3A,0x3A), RGBColor(0x3A,0x3A,0x3A), RED]
for j, (htext, hfill) in enumerate(zip(
        ['구분', '에버컴퍼니 (기존)', '이야기', '아이디컴'], HDR_FILLS)):
    cf(tbl3.cell(0, j), htext,
       size=12, bold=True, color=WHITE, fill=hfill,
       bc='1A1A1A', bw=0.5, align=PP_ALIGN.CENTER, mt=Pt(4))

# 데이터
rows_data = [
    ('500개 구간',   '견적 없음',  '견적 없음',  '3,300원',  True,  False),
    ('1,000개 구간', '5,500원',    '5,200원',    '2,900원',  True,  False),
    ('1,500개 구간', '4,590원',    '견적 없음',  '2,600원',  True,  False),
    ('2,000개 구간', '3,850원',    '4,700원',    '2,300원',  True,  False),
    ('제작 기간',    '4~5주 소요', '미확인',      '10일 소요', False, True),
    ('샘플 대응',    '재고있음',   '불가능',      '가능(요청시)', False, True),
]
NA_VALS = {'견적 없음', '미확인', '불가능'}

for i, (label, ever, story, idcom, is_price, is_adv) in enumerate(rows_data):
    row = i + 1
    bg = WHITE if i % 2 == 0 else GRAY1

    # 구분 열
    cf(tbl3.cell(row, 0), label, size=11, bold=True, color=GRAY4,
       fill=GRAY1, bc='C8C8C8', bw=0.5, align=PP_ALIGN.CENTER, mt=Pt(4))

    # 에버컴퍼니
    c_ever = GRAY3 if ever in NA_VALS else BLACK
    cf(tbl3.cell(row, 1), ever, size=12, bold=False, color=c_ever,
       fill=bg, bc='C8C8C8', bw=0.5, align=PP_ALIGN.CENTER, mt=Pt(4))

    # 이야기
    c_story = GRAY3 if story in NA_VALS else BLACK
    cf(tbl3.cell(row, 2), story, size=12, bold=False, color=c_story,
       fill=bg, bc='C8C8C8', bw=0.5, align=PP_ALIGN.CENTER, mt=Pt(4))

    # 아이디컴
    if is_price:
        cf_runs(tbl3.cell(row, 3), [
            {'text': idcom, 'size': 13.5, 'bold': True, 'color': RED},
            {'text': '   ▼ 최저', 'size': 8.5, 'bold': True, 'color': RED},
        ], fill=bg, bc='CC0000', bw=0.8, mt=Pt(4))
    elif is_adv:
        cf(tbl3.cell(row, 3), idcom, size=12, bold=True, color=BLUE,
           fill=bg, bc='C8C8C8', bw=0.5, align=PP_ALIGN.CENTER, mt=Pt(4))
    else:
        cf(tbl3.cell(row, 3), idcom, size=12, bold=False, color=BLACK,
           fill=bg, bc='C8C8C8', bw=0.5, align=PP_ALIGN.CENTER, mt=Pt(4))

# 하단 라인
line(s3, Inches(0), H - Inches(0.03), W, Inches(0.03), RED)
simple_tb(s3, '04 / 04', Inches(11.2), H - Inches(0.25), Inches(1.9), Inches(0.22),
          size=8, color=GRAY2, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════
#  SLIDE 4  수량별 업체 단가 순위표
# ══════════════════════════════════════════════════════════
s4 = prs.slides.add_slide(blank)

line(s4, Inches(0), Inches(0), W, Inches(0.04), RED)
simple_tb(s4, '수량별 업체 단가 순위표',
          Inches(0.5), Inches(0.15), Inches(9), Inches(0.7),
          size=22, bold=True, color=BLACK)
simple_tb(s4, '04 / 04',
          Inches(11.5), Inches(0.18), Inches(1.6), Inches(0.55),
          size=11, color=GRAY3, align=PP_ALIGN.RIGHT)
line(s4, Inches(0.5), Inches(0.88), W - Inches(0.6), Inches(0.018), GRAY2)

scenarios = [
    {'qty':   '1,000개 발주 시',
     'ever':  '5,500,000원',
     'story': '5,200,000원',
     'story_na': False,
     'idcom': '2,900,000원',
     'save':  '2,600,000원'},
    {'qty':   '1,500개 발주 시',
     'ever':  '6,885,000원',
     'story': '진행 불가',
     'story_na': True,
     'idcom': '3,900,000원',
     'save':  '2,985,000원'},
    {'qty':   '2,000개 발주 시',
     'ever':  '7,700,000원',
     'story': '9,400,000원',
     'story_na': False,
     'idcom': '4,600,000원',
     'save':  '3,100,000원'},
]

PANEL_Y   = Inches(1.08)
PANEL_H   = Inches(6.18)
PANEL_W   = Inches(4.05)
PANEL_GAP = Inches(0.19)
ITEM_H    = Inches(1.22)
ITEM_GAP  = Inches(0.1)

for pi, sc in enumerate(scenarios):
    PX = Inches(0.42) + pi * (PANEL_W + PANEL_GAP)

    # 패널 구분선 (우측, 첫 패널 제외)
    if pi > 0:
        line(s4, PX - PANEL_GAP / 2, PANEL_Y,
             Inches(0.018), PANEL_H, GRAY2)

    # 패널 헤더 (수량 레이블) — 블루 텍스트, 하단 레드 얇은 선
    simple_tb(s4, sc['qty'], PX, PANEL_Y + Inches(0.08),
              PANEL_W, Inches(0.5),
              size=15, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    line(s4, PX, PANEL_Y + Inches(0.62), PANEL_W, Inches(0.025), GRAY2)

    ITEM_Y0 = PANEL_Y + Inches(0.72)

    def panel_row(slide, px, iy, label, amount, amount_color=BLACK,
                  label_color=GRAY3, amount_size=16, label_size=9):
        """패널 내 업체 1행"""
        simple_tb(slide, label, px + Inches(0.12), iy + Inches(0.05),
                  PANEL_W - Inches(0.24), Inches(0.32),
                  size=label_size, color=label_color)
        simple_tb(slide, amount, px + Inches(0.12), iy + Inches(0.4),
                  PANEL_W - Inches(0.24), Inches(0.72),
                  size=amount_size, bold=True, color=amount_color,
                  align=PP_ALIGN.CENTER)

    # 에버컴퍼니
    panel_row(s4, PX, ITEM_Y0, '에버컴퍼니 (기존)', sc['ever'])
    line(s4, PX + Inches(0.1), ITEM_Y0 + ITEM_H - Inches(0.02),
         PANEL_W - Inches(0.2), Inches(0.015), GRAY2)

    # 이야기
    iy2 = ITEM_Y0 + ITEM_H + ITEM_GAP
    story_color = GRAY3 if sc['story_na'] else BLACK
    story_size  = 13 if sc['story_na'] else 16
    panel_row(s4, PX, iy2, '이야기', sc['story'],
              amount_color=story_color, amount_size=story_size)
    line(s4, PX + Inches(0.1), iy2 + ITEM_H - Inches(0.02),
         PANEL_W - Inches(0.2), Inches(0.015), GRAY2)

    # 아이디컴 (최저가, 레드)
    iy3 = iy2 + ITEM_H + ITEM_GAP
    simple_tb(s4, '아이디컴', PX + Inches(0.12), iy3 + Inches(0.05),
              PANEL_W * 0.6, Inches(0.32), size=9, color=GRAY3)
    simple_tb(s4, '▼ 최저',
              PX + PANEL_W * 0.6, iy3 + Inches(0.05),
              PANEL_W * 0.35, Inches(0.32),
              size=8, bold=True, color=RED, align=PP_ALIGN.RIGHT)
    simple_tb(s4, sc['idcom'],
              PX + Inches(0.12), iy3 + Inches(0.4),
              PANEL_W - Inches(0.24), Inches(0.72),
              size=20, bold=True, color=RED, align=PP_ALIGN.CENTER)

    # 절감 예상 (블루, 하단 강조)
    SAVE_Y = iy3 + ITEM_H + ITEM_GAP + Inches(0.1)
    line(s4, PX, SAVE_Y - Inches(0.06), PANEL_W, Inches(0.025), GRAY2)
    simple_tb(s4, '기존 대비 절감 예상',
              PX + Inches(0.12), SAVE_Y + Inches(0.08),
              PANEL_W - Inches(0.24), Inches(0.32),
              size=9, color=BLUE, align=PP_ALIGN.CENTER)
    simple_tb(s4, sc['save'],
              PX + Inches(0.12), SAVE_Y + Inches(0.42),
              PANEL_W - Inches(0.24), Inches(0.72),
              size=18, bold=True, color=BLUE, align=PP_ALIGN.CENTER)

# 하단 라인
line(s4, Inches(0), H - Inches(0.03), W, Inches(0.03), RED)
simple_tb(s4, '04 / 04', Inches(11.2), H - Inches(0.25), Inches(1.9), Inches(0.22),
          size=8, color=GRAY2, align=PP_ALIGN.RIGHT)


# ── 저장 ─────────────────────────────────────────────────
prs.save(OUT)
print(f'저장 완료: {OUT}')
