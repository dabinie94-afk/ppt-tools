# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData

OUT_XLSX = r'C:\Users\최다빈\Desktop\PPT\수량별_업체비교_보고서.xlsx'
OUT_PPTX = r'C:\Users\최다빈\Desktop\PPT\수량별_업체비교_보고서.pptx'

# ── COLORS ────────────────────────────────────────────────────────────
NAVY  = RGBColor(0x0F,0x34,0x60); RED   = RGBColor(0xE9,0x45,0x60)
GREEN = RGBColor(0x1E,0x8E,0x3E); ORANGE= RGBColor(0xF5,0xA6,0x23)
BG    = RGBColor(0xF0,0xF2,0xF5); WHITE = RGBColor(0xFF,0xFF,0xFF)
DARK  = RGBColor(0x1A,0x1A,0x2E); GRAY  = RGBColor(0x7A,0x85,0x99)
FONT  = "맑은 고딕"

# ── DATA ──────────────────────────────────────────────────────────────
quantities = ['500개','1,000개','1,500개','2,000개']
ever  = [[None,None,None],[5500,6050,6050000],[4590,5049,7573500],[3850,4235,8470000]]
iyagi = [[None,None,None],[5200,5720,5720000],[None,None,None],   [4700,4970,10340000]]
idcom = [[3300,3630,1815000],[2900,3190,3190000],[2600,2860,4290000],[2300,2530,5060000]]
iyagi_diff_total = [None, -330000, None, 1870000]
idcom_diff_total = [None, -2860000, -3283500, -3410000]

# ═══════════════════════════════════════════════════════════════════════
# EXCEL
# ═══════════════════════════════════════════════════════════════════════
def make_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "업체비교표"

    def F(h): return PatternFill(start_color=h,end_color=h,fill_type="solid")
    def fn(bold=False,color="1A1A2E",size=10,name="맑은 고딕"):
        return Font(bold=bold,color=color,size=size,name=name)
    def al(h="center",v="center",wrap=False):
        return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
    def bd():
        s=Side(style="thin"); return Border(left=s,right=s,top=s,bottom=s)
    def bd_dash():
        return Border(left=Side(style="thin"),right=Side(style="thin"),
                      top=Side(style="dashed"),bottom=Side(style="thin"))

    # Column widths
    for i,w in enumerate([10,9,12,14,9,12,14,14,9,12,14,14]):
        ws.column_dimensions[get_column_letter(i+1)].width = w

    # Row 1: Title
    ws.merge_cells('A1:L1')
    ws['A1'] = '수량별 업체 비교표'; ws['A1'].font=fn(True,"FFFFFF",15)
    ws['A1'].fill=F("0F3460"); ws['A1'].alignment=al(); ws.row_dimensions[1].height=38

    # Row 2: Subtitle
    ws.merge_cells('A2:L2')
    ws['A2'] = '비교 업체: (기존) 에버컴퍼니 / 이야기 / 아이디컴   ·   단위: 원(KRW), VAT 포함'
    ws['A2'].font=fn(False,"5A6578",10); ws['A2'].fill=F("EEF1F6")
    ws['A2'].alignment=al(); ws.row_dimensions[2].height=20

    # Row 3: Company headers
    ws.row_dimensions[3].height=26
    ws.merge_cells('A3:A4'); ws['A3']='수량'; ws['A3'].font=fn(True,"FFFFFF",11)
    ws['A3'].fill=F("1A1A2E"); ws['A3'].alignment=al(); ws['A3'].border=bd()

    for cells,label,bg in [('B3:D3','에버컴퍼니 (기존 업체)','0F3460'),
                            ('E3:H3','이야기','C0392B'),
                            ('I3:L3','아이디컴','1B7A3C')]:
        ws.merge_cells(cells)
        c=ws[cells.split(':')[0]]; c.value=label
        c.font=fn(True,"FFFFFF",11); c.fill=F(bg); c.alignment=al(); c.border=bd()

    # Row 4: Sub-headers
    ws.row_dimensions[4].height=32
    sub=[('단가','D9E4F5'),('개당\n(VAT포함)','D9E4F5'),('합계\n(VAT포함)','BBCFE8'),
         ('단가','FADDDD'),('개당\n(VAT포함)','FADDDD'),('합계\n(VAT포함)','F5C6C6'),('기존대비\n차액','F5C6C6'),
         ('단가','DBEEDD'),('개당\n(VAT포함)','DBEEDD'),('합계\n(VAT포함)','B8DDCA'),('기존대비\n차액','B8DDCA')]
    ws['A4'].border=bd()
    for i,(hdr,bg) in enumerate(sub):
        c=ws.cell(4,i+2,hdr); c.font=fn(True,"2D3748",9)
        c.fill=F(bg); c.alignment=al(wrap=True); c.border=bd()

    # Data rows
    row=5
    for qi,qty in enumerate(quantities):
        e=ever[qi]; iy=iyagi[qi]; id_=idcom[qi]
        ws.row_dimensions[row].height=22

        has_diff = (iyagi_diff_total[qi] is not None or idcom_diff_total[qi] is not None)
        if has_diff:
            ws.row_dimensions[row+1].height=16
            for cols in [f'A{row}:A{row+1}',f'B{row}:B{row+1}',
                         f'C{row}:C{row+1}',f'D{row}:D{row+1}']:
                ws.merge_cells(cols)

        def nc(r,col,val,bg,bold=False,color="1A1A2E",best=False,worst=False):
            c=ws.cell(r,col)
            if val is None: c.value='-'; c.font=fn(False,"C0C4CC",9)
            else:
                c.value=val; c.number_format='#,##0' if isinstance(val,int) else '@'
                clr="1B7A3C" if best else ("C0392B" if worst else color)
                c.font=fn(bold,clr,10)
            c.fill=F(bg); c.alignment=al(); c.border=bd()

        # Qty
        c=ws.cell(row,1,qty); c.font=fn(True,"0F3460",11)
        c.fill=F("E8EDF5"); c.alignment=al(); c.border=bd()

        # Ever
        nc(row,2,e[0],"EDF2FA"); nc(row,3,e[1],"EDF2FA"); nc(row,4,e[2],"D9E4F5",True)

        # Iyagi
        iy_worse = iy[2] is not None and e[2] is not None and iy[2]>e[2]
        iy_better= iy[2] is not None and e[2] is not None and iy[2]<e[2]
        nc(row,5,iy[0],"FFF5F5"); nc(row,6,iy[1],"FFF5F5")
        nc(row,7,iy[2],"FADDDD",True,best=iy_better,worst=iy_worse)

        # Iyagi 차액
        iy_d=iyagi_diff_total[qi]
        c8=ws.cell(row,8)
        if iy_d is not None:
            c8.value=iy_d; c8.number_format='+#,##0;-#,##0;"-"'
            c8.font=fn(True,"1B7A3C" if iy_d<0 else "C0392B",11)
            c8.fill=F("E8F7EC" if iy_d<0 else "FDECEA")
        else:
            c8.value='-'; c8.font=fn(False,"C0C4CC",9); c8.fill=F("F8F8F8")
        c8.alignment=al(); c8.border=bd()

        # IDcom
        all_t=[v[2] for v in [e,iy] if v[2] is not None]
        id_best=id_[2] is not None and (not all_t or id_[2]<min(all_t))
        nc(row,9,id_[0],"F0FAF3"); nc(row,10,id_[1],"F0FAF3")
        nc(row,11,id_[2],"DBEEDD",True,best=id_best)

        # IDcom 차액
        id_d=idcom_diff_total[qi]
        c12=ws.cell(row,12)
        if id_d is not None:
            c12.value=id_d; c12.number_format='+#,##0;-#,##0;"-"'
            c12.font=fn(True,"1B7A3C" if id_d<0 else "C0392B",11)
            c12.fill=F("E8F7EC" if id_d<0 else "FDECEA")
        else:
            c12.value='-'; c12.font=fn(False,"C0C4CC",9); c12.fill=F("F8F8F8")
        c12.alignment=al(); c12.border=bd()

        # Diff row fill
        if has_diff:
            r2=row+1
            for col_,bg_ in zip(range(5,13),['FFF5F5','FFF5F5','FADDDD','FDE8EC',
                                              'F0FAF3','F0FAF3','DBEEDD','F0F7EC']):
                c=ws.cell(r2,col_); c.fill=F(bg_); c.border=bd_dash()
            ws.merge_cells(f'H{row}:H{r2}'); ws.merge_cells(f'L{row}:L{r2}')
            row+=2
        else:
            row+=1

    ws.freeze_panes='B5'

    # ── Charts data sheet ──────────────────────────────────────────
    cd=wb.create_sheet("차트데이터"); cd.sheet_state='hidden'
    cd['A1']='수량'; cd['B1']='에버컴퍼니'; cd['C1']='이야기'; cd['D1']='아이디컴'
    for i,(q,e_,iy_,id__) in enumerate(zip(['1,000개','1,500개','2,000개'],
                                            [6050000,7573500,8470000],
                                            [5720000,None,10340000],
                                            [3190000,4290000,5060000])):
        cd.cell(i+2,1,q); cd.cell(i+2,2,e_)
        if iy_ is not None: cd.cell(i+2,3,iy_)
        cd.cell(i+2,4,id__)

    cd['F1']='수량'; cd['G1']='에버컴퍼니'; cd['H1']='이야기'; cd['I1']='아이디컴'
    for i,(q,e_,iy_,id__) in enumerate(zip(['500개','1,000개','1,500개','2,000개'],
                                            [None,6050,5049,4235],
                                            [None,5720,None,4970],
                                            [3630,3190,2860,2530])):
        cd.cell(i+2,6,q)
        if e_ is not None: cd.cell(i+2,7,e_)
        if iy_ is not None: cd.cell(i+2,8,iy_)
        cd.cell(i+2,9,id__)

    # Bar chart
    bar=BarChart(); bar.type="col"; bar.grouping="clustered"
    bar.title="총 발주 금액 비교 (VAT 포함)"; bar.style=10; bar.width=18; bar.height=12
    bar.y_axis.title="금액 (원)"; bar.x_axis.title="발주 수량"
    cats=Reference(cd,min_col=1,min_row=2,max_row=4)
    for col,name in [(2,'에버컴퍼니'),(3,'이야기'),(4,'아이디컴')]:
        d=Reference(cd,min_col=col,min_row=1,max_row=4)
        bar.add_data(d,titles_from_data=True)
    bar.set_categories(cats)
    ws.add_chart(bar,f'A{row+1}')

    # Line chart
    line=LineChart(); line.title="수량별 개당 단가 추이 (VAT 포함)"
    line.style=10; line.width=18; line.height=12
    line.y_axis.title="단가 (원/개)"; line.x_axis.title="발주 수량"
    cats2=Reference(cd,min_col=6,min_row=2,max_row=5)
    for col,name in [(7,'에버컴퍼니'),(8,'이야기'),(9,'아이디컴')]:
        d=Reference(cd,min_col=col,min_row=1,max_row=5)
        line.add_data(d,titles_from_data=True)
    line.set_categories(cats2)
    ws.add_chart(line,f'G{row+1}')

    wb.save(OUT_XLSX)
    print(f"Excel saved: {OUT_XLSX}")

# ═══════════════════════════════════════════════════════════════════════
# POWERPOINT
# ═══════════════════════════════════════════════════════════════════════
def make_pptx():
    prs=Presentation()
    prs.slide_width=Inches(13.33); prs.slide_height=Inches(7.5)
    BLANK=prs.slide_layouts[6]

    def rect(slide,x,y,w,h,fill=None):
        s=slide.shapes.add_shape(1,Inches(x),Inches(y),Inches(w),Inches(h))
        if fill: s.fill.solid(); s.fill.fore_color.rgb=fill
        else: s.fill.background()
        s.line.fill.background()
        return s

    def txt(slide,x,y,w,h,text,size,bold=False,color=DARK,align=PP_ALIGN.LEFT,italic=False):
        tb=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
        tf=tb.text_frame; tf.word_wrap=True
        p=tf.paragraphs[0]; p.alignment=align
        r=p.add_run(); r.text=text
        r.font.size=Pt(size); r.font.bold=bold
        r.font.color.rgb=color; r.font.name=FONT; r.font.italic=italic
        return tb

    def header(slide,title,sub):
        rect(slide,0,0,13.33,7.5,BG)
        rect(slide,0,0,13.33,1.25,NAVY)
        txt(slide,0.4,0.15,12,0.7,title,22,True,WHITE)
        txt(slide,0.4,0.75,12,0.45,sub,10.5,color=RGBColor(0xA0,0xB4,0xC8))

    # ── Slide 1: Title ────────────────────────────────────────────────
    s1=prs.slides.add_slide(BLANK)
    rect(s1,0,0,13.33,7.5,DARK)
    rect(s1,0,0,0.12,7.5,RED)
    txt(s1,1.0,1.1,11.5,0.5,'EXECUTIVE BRIEFING  ·  구매전략 검토',
        12,color=RGBColor(0x70,0x90,0xB0),italic=True)
    txt(s1,1.0,1.7,11.5,1.5,'수량별 업체 비교 분석',44,True,WHITE)
    txt(s1,1.0,3.3,11.5,0.6,'에버컴퍼니(기존) · 이야기 · 아이디컴  |  VAT 포함 기준',
        14,color=RGBColor(0xA0,0xB4,0xC8))
    for i,(lbl,col) in enumerate([('에버컴퍼니',NAVY),('이야기',RED),('아이디컴',GREEN)]):
        rect(s1,1.0+i*2.6,4.2,2.4,0.5,col)
        txt(s1,1.0+i*2.6,4.27,2.4,0.36,lbl,11,True,WHITE,PP_ALIGN.CENTER)
    txt(s1,10.0,7.0,3.0,0.4,'2026년 5월  ·  총무팀',10,
        color=RGBColor(0x50,0x68,0x88),align=PP_ALIGN.RIGHT)

    # ── Slide 2: KPI ──────────────────────────────────────────────────
    s2=prs.slides.add_slide(BLANK)
    header(s2,'핵심 요약','Key Findings | 수량별 업체 비교 분석 결과')
    kpis=[('최대 절감 가능','△ 3,410,000원','2,000개 발주 시 아이디컴 채택\n기존 대비 40.3% 절감',RED),
          ('최저 단가 업체','아이디컴','전 수량 일관된 최저가 달성\n1,000개 개당 3,190원(VAT포함)',GREEN),
          ('이야기 주의','▲ 1,870,000원','2,000개 발주 시 기존 대비 초과\n1,000개만 소폭 절감(△5.5%)',ORANGE)]
    for i,(title,val,desc,col) in enumerate(kpis):
        cx=0.4+i*4.3
        rect(s2,cx,1.4,4.0,5.7,WHITE)
        rect(s2,cx,1.4,4.0,0.09,col)
        txt(s2,cx+0.2,1.57,3.6,0.45,title,11,color=GRAY)
        txt(s2,cx+0.1,2.1,3.8,0.85,val,23,True,col,PP_ALIGN.CENTER)
        txt(s2,cx+0.2,3.15,3.6,1.2,desc,11,color=RGBColor(0x4A,0x55,0x68))
    rect(s2,0.4,6.85,12.5,0.5,RGBColor(0xE8,0xF0,0xFE))
    txt(s2,0.6,6.9,12.1,0.4,
        '→ 권고: 아이디컴을 신규 발주처로 채택 검토  |  품질 · 납기 · 거래안정성 검증 전제',
        11,True,NAVY)

    # ── Slide 3: Bar Chart ────────────────────────────────────────────
    s3=prs.slides.add_slide(BLANK)
    header(s3,'총 발주 금액 비교 (VAT 포함)','수량별 합계 금액 비교 | 단위: 원(KRW)')
    cd3=CategoryChartData()
    cd3.categories=['1,000개','1,500개','2,000개']
    cd3.add_series('에버컴퍼니(기존)',(6050000,7573500,8470000))
    cd3.add_series('이야기',(5720000,None,10340000))
    cd3.add_series('아이디컴',(3190000,4290000,5060000))
    ch3=s3.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.5),Inches(1.35),Inches(12.3),Inches(5.9),cd3).chart
    ch3.has_legend=True
    ch3.legend.position=XL_LEGEND_POSITION.BOTTOM
    ch3.legend.include_in_layout=False
    try:
        ch3.series[0].format.fill.solid(); ch3.series[0].format.fill.fore_color.rgb=NAVY
        ch3.series[1].format.fill.solid(); ch3.series[1].format.fill.fore_color.rgb=RED
        ch3.series[2].format.fill.solid(); ch3.series[2].format.fill.fore_color.rgb=GREEN
    except: pass

    # ── Slide 4: Line Chart ───────────────────────────────────────────
    s4=prs.slides.add_slide(BLANK)
    header(s4,'수량별 개당 단가 추이 (VAT 포함)','발주량 증가에 따른 단가 인하 효과 | 단위: 원/개')
    cd4=CategoryChartData()
    cd4.categories=['500개','1,000개','1,500개','2,000개']
    cd4.add_series('에버컴퍼니(기존)',(None,6050,5049,4235))
    cd4.add_series('이야기',(None,5720,None,4970))
    cd4.add_series('아이디컴',(3630,3190,2860,2530))
    ch4=s4.shapes.add_chart(XL_CHART_TYPE.LINE,
        Inches(0.5),Inches(1.35),Inches(12.3),Inches(5.9),cd4).chart
    ch4.has_legend=True
    ch4.legend.position=XL_LEGEND_POSITION.BOTTOM
    ch4.legend.include_in_layout=False
    try:
        for i,col in enumerate([NAVY,RED,GREEN]):
            ch4.series[i].format.line.color.rgb=col
            ch4.series[i].format.line.width=Pt(2.5)
    except: pass

    # ── Slide 5: Table ────────────────────────────────────────────────
    s5=prs.slides.add_slide(BLANK)
    header(s5,'수량별 상세 비교표',
           '에버컴퍼니(기존) vs 이야기 vs 아이디컴  |  △ 절감  ▲ 비용증가  |  단위: 원, VAT포함')

    tbl_rows=[
        ['수량','에버컴퍼니(기존)','','이야기','','','아이디컴','',''],
        ['','단가','합계(VAT포함)','단가','합계(VAT포함)','기존대비 차액','단가','합계(VAT포함)','기존대비 차액'],
        ['500개',  '-',     '-',          '-',     '-',           '-',          '3,300', '1,815,000','비교불가'],
        ['1,000개','6,050', '6,050,000',  '5,200', '5,720,000',   '△330,000',  '2,900', '3,190,000','△2,860,000'],
        ['1,500개','4,590', '7,573,500',  '-',     '-',           '-',          '2,600', '4,290,000','△3,283,500'],
        ['2,000개','3,850', '8,470,000',  '4,700', '10,340,000',  '▲1,870,000','2,300', '5,060,000','△3,410,000'],
    ]

    tbl=s5.shapes.add_table(6,9,Inches(0.2),Inches(1.33),
                             Inches(12.93),Inches(6.0)).table

    pcts=[0.09,0.10,0.14,0.10,0.14,0.13,0.10,0.14,0.12]
    TW=Inches(12.93)
    for i,p in enumerate(pcts): tbl.columns[i].width=int(TW*p)
    for i,h in enumerate([Inches(0.52),Inches(0.58),
                           Inches(1.22),Inches(1.22),Inches(1.22),Inches(1.22)]):
        tbl.rows[i].height=h

    BG_EVER =RGBColor(0xED,0xF2,0xFA); BG_IYAGI=RGBColor(0xFF,0xF5,0xF5)
    BG_IDCOM=RGBColor(0xF0,0xFA,0xF3); BG_QTY  =RGBColor(0xF4,0xF6,0xFA)
    HN=RGBColor(0x0F,0x34,0x60); HR=RGBColor(0xC0,0x39,0x2B); HG=RGBColor(0x1B,0x7A,0x3C)
    SN=RGBColor(0xD9,0xE4,0xF5); SR=RGBColor(0xFA,0xDD,0xDD); SG=RGBColor(0xDB,0xEE,0xDD)

    def sc(row,col,text,bg=None,fcolor=DARK,fsize=9.5,bold=False,align=PP_ALIGN.CENTER):
        cell=tbl.cell(row,col); cell.text=text
        tf=cell.text_frame; p=tf.paragraphs[0]; p.alignment=align
        r=p.runs[0] if p.runs else p.add_run(); r.text=text
        r.font.size=Pt(fsize); r.font.bold=bold
        r.font.color.rgb=fcolor; r.font.name=FONT
        if bg: cell.fill.solid(); cell.fill.fore_color.rgb=bg

    # Row 0: company headers
    sc(0,0,'수량',HN,WHITE,11,True)
    sc(0,1,'에버컴퍼니 (기존 업체)',HN,WHITE,11,True)
    sc(0,2,'',HN,WHITE)
    sc(0,3,'이야기',HR,WHITE,11,True)
    sc(0,4,'',HR,WHITE); sc(0,5,'',HR,WHITE)
    sc(0,6,'아이디컴',HG,WHITE,11,True)
    sc(0,7,'',HG,WHITE); sc(0,8,'',HG,WHITE)
    try:
        tbl.cell(0,1).merge(tbl.cell(0,2))
        tbl.cell(0,3).merge(tbl.cell(0,5))
        tbl.cell(0,6).merge(tbl.cell(0,8))
    except: pass

    # Row 1: sub-headers
    sub_bgs=[BG_QTY,SN,SN,SR,SR,SR,SG,SG,SG]
    for j,(t,bg) in enumerate(zip(tbl_rows[1],sub_bgs)):
        sc(1,j,t,bg,RGBColor(0x2D,0x37,0x48),9,True)

    # Data rows
    col_bgs=[BG_QTY,BG_EVER,BG_EVER,BG_IYAGI,BG_IYAGI,BG_IYAGI,BG_IDCOM,BG_IDCOM,BG_IDCOM]
    for di,row_data in enumerate(tbl_rows[2:]):
        ri=di+2
        for j,(val,bg) in enumerate(zip(row_data,col_bgs)):
            fc=DARK
            if j==0: fc=RGBColor(0x0F,0x34,0x60)
            elif val.startswith('△'): fc=RGBColor(0x1B,0x5E,0x20)
            elif val.startswith('▲'): fc=RGBColor(0xB7,0x1C,0x1C)
            elif val=='-' or val=='비교불가': fc=RGBColor(0xC0,0xC4,0xCC)
            elif j in [2,4,7] and val!='-':
                # IDcom total highlight
                if j==7: fc=RGBColor(0x1B,0x5E,0x20)
            is_bold=(j in [0,2,4,5,7,8])
            fsize=10 if j in [2,4,5,7,8] else (10.5 if j==0 else 9.5)
            sc(ri,j,val,bg,fc,fsize,is_bold)

    prs.save(OUT_PPTX)
    print(f"PPT saved: {OUT_PPTX}")

make_excel()
make_pptx()
print("완료!")
