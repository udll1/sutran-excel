from flask import Flask, request, send_file, jsonify
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
import io, json

app = Flask(__name__)

COLORES = {
    "A":"C6EFCE","T":"FFEB9C","F":"FFC7CE","DS":"D9D9D9",
    "CP":"FFE6CC","CPSS":"FFE6CC","FE":"E1D5E7","FR":"E1D5E7",
    "C":"D5E8D4","CO":"FFF2CC","DM":"DAE8FC","LM":"DAE8FC",
    "LP":"DAE8FC","LO":"DAE8FC","LUB":"DAE8FC","LS":"DAE8FC",
    "V":"DAE8FC","REVISION":"FCE4D6"
}

def fill(color): return PatternFill('solid', start_color=color)
def border():
    t = Side(style='thin', color='AAAAAA')
    return Border(left=t, right=t, top=t, bottom=t)
def center(): return Alignment(horizontal='center', vertical='center', wrap_text=True)
def left_al(): return Alignment(horizontal='left', vertical='center', wrap_text=True)

def setup_page(ws):
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.4, bottom=0.4)

def get_semanas(dias_mes):
    semanas = []
    d = 1
    while d <= dias_mes:
        semanas.append((d, min(d+6, dias_mes)))
        d = min(d+6, dias_mes) + 1
    return semanas

SEM_COLORS = ["1F4E79","2E75B6","1F4E79","2E75B6","1F4E79"]

def generar_mensual(trabajadores, dias_mes, mes_nom, anio, leyenda):
    wb = Workbook()
    ws = wb.active
    ws.title = f"ASISTENCIA {mes_nom} {anio}"
    semanas = get_semanas(dias_mes)

    COL_N=1; COL_DNI=2; COL_NOMBRE=3; COL_CARGO=4
    COL_DIA1=5
    COL_TF = COL_DIA1+dias_mes
    COL_TT = COL_TF+1
    COL_TA = COL_TT+1
    COL_TDS = COL_TA+1
    TC = COL_TDS

    hdr = fill("1F3864")
    azul = fill("2E75B6")

    # Fila 1 - Titulo completo
    ws.merge_cells(1,1,1,TC)
    c = ws.cell(1,1,f"CONTROL DE ASISTENCIA - SUTRAN UDLL - {mes_nom} {anio}")
    c.font=Font(name='Arial',bold=True,size=10,color='FFFFFF'); c.fill=hdr; c.alignment=center()
    ws.row_dimensions[1].height=20

    # Fila 2 - Semanas
    for col in range(1,5): ws.cell(2,col).fill=hdr
    for si,(s1,s2) in enumerate(semanas):
        c1=COL_DIA1+s1-1; c2=COL_DIA1+s2-1
        if c1<c2: ws.merge_cells(2,c1,2,c2)
        c=ws.cell(2,c1,f"SEMANA {si+1}")
        c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(SEM_COLORS[si]); c.alignment=center()
    for col in range(COL_TF,TC+1): ws.cell(2,col).fill=hdr
    ws.row_dimensions[2].height=12

    # Fila 3 - Encabezados
    for col,h,w in [(COL_N,'N',4),(COL_DNI,'DNI',10),(COL_NOMBRE,'APELLIDOS Y NOMBRES',35),(COL_CARGO,'CARGO',18)]:
        c=ws.cell(3,col,h); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
        c.fill=azul; c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=w
    for dd in range(1,dias_mes+1):
        si=(dd-1)//7; col=COL_DIA1+dd-1
        c=ws.cell(3,col,dd); c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(SEM_COLORS[si]); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=3.0
    for col,h,bg,w in [(COL_TF,'T.F','C0504D',4.5),(COL_TT,'T.T','E36C09',4.5),(COL_TA,'T.A','4F81BD',4.5),(COL_TDS,'T.DS','2E75B6',4.5)]:
        c=ws.cell(3,col,h); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
        c.fill=fill(bg); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.row_dimensions[3].height=24

    # Datos
    for ti,t in enumerate(trabajadores):
        r=ti+4; ws.row_dimensions[r].height=14
        for col,val,al in [(COL_N,t['item'],center()),(COL_DNI,t.get('dni',''),center()),(COL_NOMBRE,t['nombre_completo'],left_al()),(COL_CARGO,t.get('cargo',''),center())]:
            c=ws.cell(r,col,val); c.font=Font(name='Arial',size=7); c.alignment=al; c.border=border()
        tF=tT=tA=tDS=0
        for dd in range(1,dias_mes+1):
            est=str(t['dias'].get(str(dd),t['dias'].get(dd,'')))
            c=ws.cell(r,COL_DIA1+dd-1,est if est else '')
            c.font=Font(name='Arial',size=6,bold=bool(est)); c.alignment=center(); c.border=border()
            if est in COLORES: c.fill=fill(COLORES[est])
            if est=='F': tF+=1
            elif est=='T': tT+=1
            elif est=='A': tA+=1
            elif est=='DS': tDS+=1
        for col,val,bg in [(COL_TF,tF,'FFC7CE'),(COL_TT,tT,'FFEB9C'),(COL_TA,tA,'C6EFCE'),(COL_TDS,tDS,'D9D9D9')]:
            c=ws.cell(r,col,val); c.font=Font(name='Arial',size=7,bold=True)
            c.alignment=center(); c.border=border(); c.fill=fill(bg)

    # Leyenda
    lr=len(trabajadores)+5
    ws.merge_cells(lr,1,lr,2)
    c=ws.cell(lr,1,'LEYENDA'); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
    c.fill=hdr; c.alignment=center(); ws.row_dimensions[lr].height=13
    for li,ley in enumerate(leyenda):
        r=lr+1+li; ws.row_dimensions[r].height=11
        c1=ws.cell(r,1,ley['cod']); c1.font=Font(name='Arial',bold=True,size=6)
        if ley['cod'] in COLORES: c1.fill=fill(COLORES[ley['cod']])
        c1.alignment=center(); c1.border=border()
        c2=ws.cell(r,2,ley['desc']); c2.font=Font(name='Arial',size=6)
        if ley['cod'] in COLORES: c2.fill=fill(COLORES[ley['cod']])
        c2.alignment=left_al(); c2.border=border()

    setup_page(ws)
    ws.print_title_rows='1:3'
    return wb

def generar_firmas(trabajadores, dias_mes, mes_nom, anio, leyenda):
    wb = Workbook()
    ws = wb.active
    ws.title = f"FIRMAS {mes_nom} {anio}"
    semanas = get_semanas(dias_mes)

    COL_N=1; COL_NOMBRE=2; COL_DIA1=3
    COL_TF=COL_DIA1+dias_mes
    COL_TT=COL_TF+1; COL_TA=COL_TT+1; COL_TD=COL_TA+1; COL_FIRMA=COL_TD+1
    TC=COL_FIRMA

    hdr=fill("1F3864"); azul=fill("2E75B6")

    # Fila 1
    ws.merge_cells(1,1,1,TC)
    c=ws.cell(1,1,f"CONTROL DE ASISTENCIA DEL PERSONAL SUTRAN REGION LA LIBERTAD - {mes_nom} {anio}")
    c.font=Font(name='Arial',bold=True,size=9,color='FFFFFF'); c.fill=hdr; c.alignment=center()
    ws.row_dimensions[1].height=20

    # Fila 2
    ws.merge_cells(2,1,2,TC)
    c=ws.cell(2,1,"GERENCIA DE ARTICULACION TERRITORIAL")
    c.font=Font(name='Arial',bold=True,size=8,color='FFFFFF'); c.fill=azul; c.alignment=center()
    ws.row_dimensions[2].height=14

    # Fila 3 - Semanas
    for col in [COL_N,COL_NOMBRE]: ws.cell(3,col).fill=hdr
    for si,(s1,s2) in enumerate(semanas):
        c1=COL_DIA1+s1-1; c2=COL_DIA1+s2-1
        if c1<c2: ws.merge_cells(3,c1,3,c2)
        c=ws.cell(3,c1,f"SEM {si+1}")
        c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(SEM_COLORS[si]); c.alignment=center()
    for col in range(COL_TF,TC+1): ws.cell(3,col).fill=hdr
    ws.row_dimensions[3].height=12

    # Fila 4 - Encabezados
    for col,h,w in [(COL_N,'N',4),(COL_NOMBRE,'APELLIDOS Y NOMBRES',35)]:
        c=ws.cell(4,col,h); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
        c.fill=azul; c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=w
    for dd in range(1,dias_mes+1):
        si=(dd-1)//7; col=COL_DIA1+dd-1
        c=ws.cell(4,col,dd); c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(SEM_COLORS[si]); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=3.0
    for col,h,bg,w in [(COL_TF,'T.FALTAS','C0504D',5),(COL_TT,'T.TARD','E36C09',5),(COL_TA,'T.ASIST','4F81BD',5),(COL_TD,'T.DIAS','2E75B6',5),(COL_FIRMA,'FIRMA','2E75B6',20)]:
        c=ws.cell(4,col,h); c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(bg); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.row_dimensions[4].height=24

    # Datos
    for ti,t in enumerate(trabajadores):
        r=ti+5; ws.row_dimensions[r].height=16
        for col,val,al in [(COL_N,t['item'],center()),(COL_NOMBRE,t['nombre_completo'],left_al())]:
            c=ws.cell(r,col,val); c.font=Font(name='Arial',size=7); c.alignment=al; c.border=border()
        tF=tT=tA=tD=0
        for dd in range(1,dias_mes+1):
            est=str(t['dias'].get(str(dd),t['dias'].get(dd,'')))
            c=ws.cell(r,COL_DIA1+dd-1,est if est else '')
            c.font=Font(name='Arial',size=6,bold=bool(est)); c.alignment=center(); c.border=border()
            if est in COLORES: c.fill=fill(COLORES[est])
            if est=='F': tF+=1
            elif est=='T': tT+=1
            elif est=='A': tA+=1; tD+=1
        for col,val,bg in [(COL_TF,tF,'FFC7CE'),(COL_TT,tT,'FFEB9C'),(COL_TA,tA,'C6EFCE'),(COL_TD,tD,None)]:
            c=ws.cell(r,col,val); c.font=Font(name='Arial',size=7,bold=True)
            c.alignment=center(); c.border=border()
            if bg: c.fill=fill(bg)
        ws.cell(r,COL_FIRMA).border=border()

    # Leyenda
    lr=len(trabajadores)+6
    ws.merge_cells(lr,1,lr,2)
    c=ws.cell(lr,1,'LEYENDA'); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
    c.fill=hdr; c.alignment=center(); ws.row_dimensions[lr].height=13
    for li,ley in enumerate(leyenda):
        r=lr+1+li; ws.row_dimensions[r].height=11
        c1=ws.cell(r,1,ley['cod']); c1.font=Font(name='Arial',bold=True,size=6)
        if ley['cod'] in COLORES: c1.fill=fill(COLORES[ley['cod']])
        c1.alignment=center(); c1.border=border()
        c2=ws.cell(r,2,ley['desc']); c2.font=Font(name='Arial',size=6)
        if ley['cod'] in COLORES: c2.fill=fill(COLORES[ley['cod']])
        c2.alignment=left_al(); c2.border=border()

    setup_page(ws)
    ws.print_title_rows='1:4'
    return wb

def abreviar_lugar(lugar):
    """Abrevia el nombre del lugar a max 6 caracteres"""
    if not lugar: return ""
    palabras = lugar.split()
    if len(palabras) == 1: return lugar[:6].upper()
    return "".join(p[0] for p in palabras[:4]).upper()

def generar_tardanzas(trabajadores, dias_mes, mes_nom, anio, leyenda):
    wb = Workbook()
    ws = wb.active
    ws.title = f"TARDANZAS {mes_nom} {anio}"
    semanas = get_semanas(dias_mes)
    sem_colors = ["1F4E79","2E75B6","1F4E79","2E75B6","1F4E79"]

    COL_N=1; COL_DNI=2; COL_NOMBRE=3; COL_CARGO=4
    COL_DIA1=5; COL_TF=COL_DIA1+dias_mes; COL_TT=COL_TF+1
    COL_TMIN=COL_TT+1; COL_TA=COL_TMIN+1; TC=COL_TA

    hdr=fill("1F3864"); azul=fill("2E75B6")

    ws.merge_cells(1,1,1,TC)
    c=ws.cell(1,1,f"REPORTE DE TARDANZAS - SUTRAN UDLL - {mes_nom} {anio}")
    c.font=Font(name='Arial',bold=True,size=10,color='FFFFFF'); c.fill=hdr; c.alignment=center()
    ws.row_dimensions[1].height=20

    for col in range(1,5): ws.cell(2,col).fill=hdr
    for si,(s1,s2) in enumerate(semanas):
        c1=COL_DIA1+s1-1; c2=COL_DIA1+s2-1
        if c1<c2: ws.merge_cells(2,c1,2,c2)
        c=ws.cell(2,c1,f"SEMANA {si+1}")
        c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(sem_colors[si]); c.alignment=center()
    for col in range(COL_TF,TC+1): ws.cell(2,col).fill=hdr
    ws.row_dimensions[2].height=12

    for col,h,w,bg in [(COL_N,'N',4,'2E75B6'),(COL_DNI,'DNI',10,'2E75B6'),(COL_NOMBRE,'APELLIDOS Y NOMBRES',35,'2E75B6'),(COL_CARGO,'CARGO',18,'2E75B6')]:
        c=ws.cell(3,col,h); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
        c.fill=fill(bg); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=w
    for dd in range(1,dias_mes+1):
        si=(dd-1)//7; col=COL_DIA1+dd-1
        c=ws.cell(3,col,dd); c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(sem_colors[si]); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=3.0
    for col,h,bg,w in [(COL_TF,'T.F','C0504D',4.5),(COL_TT,'T.TARD','E36C09',4.5),(COL_TMIN,'MIN.T','FFB830',6),(COL_TA,'T.A','4F81BD',4.5)]:
        c=ws.cell(3,col,h); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
        c.fill=fill(bg); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.row_dimensions[3].height=24

    for ti,t in enumerate(trabajadores):
        r=ti+4; ws.row_dimensions[r].height=14
        for col,val,al in [(COL_N,t['item'],center()),(COL_DNI,t.get('dni',''),center()),(COL_NOMBRE,t['nombre_completo'],left_al()),(COL_CARGO,t.get('cargo',''),center())]:
            c=ws.cell(r,col,val); c.font=Font(name='Arial',size=7); c.alignment=al; c.border=border()
        tF=tT=tA=tMin=0
        for dd in range(1,dias_mes+1):
            dia_data = t['dias'].get(str(dd), t['dias'].get(dd, {}))
            if isinstance(dia_data, dict):
                est = str(dia_data.get('estado',''))
                entrada = str(dia_data.get('entrada',''))
                h_inicio = str(dia_data.get('h_inicio',''))
            else:
                est = str(dia_data)
                entrada = ''; h_inicio = ''
            
            # Calcular minutos tardanza
            min_tard = 0
            if est == 'T' and entrada and h_inicio:
                try:
                    pe = entrada.split(':'); pi = h_inicio.split(':')
                    min_e = int(pe[0])*60 + int(pe[1])
                    min_i = int(pi[0])*60 + int(pi[1])
                    min_tard = max(0, min_e - min_i)
                except: pass
            
            col = COL_DIA1+dd-1
            if est == 'T' and min_tard > 0:
                display = f"+{min_tard}"
                c=ws.cell(r,col,display)
                c.font=Font(name='Arial',size=6,bold=True,color='9C5700')
                c.fill=fill('FFEB9C'); c.alignment=center(); c.border=border()
                tT+=1; tMin+=min_tard
            elif est == 'T':
                c=ws.cell(r,col,'T')
                c.font=Font(name='Arial',size=6,bold=True); c.fill=fill('FFEB9C')
                c.alignment=center(); c.border=border(); tT+=1
            else:
                c=ws.cell(r,col,est if est else '')
                c.font=Font(name='Arial',size=6,bold=bool(est))
                c.alignment=center(); c.border=border()
                if est in COLORES: c.fill=fill(COLORES[est])
                if est=='F': tF+=1
                elif est=='A': tA+=1

        for col,val,bg in [(COL_TF,tF,'FFC7CE'),(COL_TT,tT,'FFEB9C'),(COL_TMIN,tMin,'FFB830'),(COL_TA,tA,'C6EFCE')]:
            c=ws.cell(r,col,val); c.font=Font(name='Arial',size=7,bold=True)
            c.alignment=center(); c.border=border(); c.fill=fill(bg)

    # Leyenda
    lr=len(trabajadores)+5
    ws.merge_cells(lr,1,lr,2)
    c=ws.cell(lr,1,'LEYENDA'); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
    c.fill=hdr; c.alignment=center(); ws.row_dimensions[lr].height=13
    for li,ley in enumerate(leyenda):
        r2=lr+1+li; ws.row_dimensions[r2].height=11
        c1=ws.cell(r2,1,ley['cod']); c1.font=Font(name='Arial',bold=True,size=6)
        if ley['cod'] in COLORES: c1.fill=fill(COLORES[ley['cod']])
        c1.alignment=center(); c1.border=border()
        c2=ws.cell(r2,2,ley['desc']); c2.font=Font(name='Arial',size=6)
        if ley['cod'] in COLORES: c2.fill=fill(COLORES[ley['cod']])
        c2.alignment=left_al(); c2.border=border()

    setup_page(ws); ws.print_title_rows='1:3'
    return wb


def generar_detalle(trabajadores, dias_mes, mes_nom, anio, leyenda):
    wb = Workbook()
    ws = wb.active
    ws.title = f"DETALLE {mes_nom} {anio}"
    semanas = get_semanas(dias_mes)
    sem_colors = ["1F4E79","2E75B6","1F4E79","2E75B6","1F4E79"]

    COL_N=1; COL_NOMBRE=2; COL_DIA1=3
    COL_TF=COL_DIA1+dias_mes; COL_TA=COL_TF+1; TC=COL_TA

    hdr=fill("1F3864"); azul=fill("2E75B6")
    gris=fill("F2F2F2")

    ws.merge_cells(1,1,1,TC)
    c=ws.cell(1,1,f"DETALLE DE ASISTENCIA - SUTRAN UDLL - {mes_nom} {anio}")
    c.font=Font(name='Arial',bold=True,size=10,color='FFFFFF'); c.fill=hdr; c.alignment=center()
    ws.row_dimensions[1].height=20

    for col in [COL_N,COL_NOMBRE]: ws.cell(2,col).fill=hdr
    for si,(s1,s2) in enumerate(semanas):
        c1=COL_DIA1+s1-1; c2=COL_DIA1+s2-1
        if c1<c2: ws.merge_cells(2,c1,2,c2)
        c=ws.cell(2,c1,f"SEMANA {si+1}")
        c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(sem_colors[si]); c.alignment=center()
    for col in range(COL_TF,TC+1): ws.cell(2,col).fill=hdr
    ws.row_dimensions[2].height=12

    for col,h,w in [(COL_N,'N',4),(COL_NOMBRE,'APELLIDOS Y NOMBRES',35)]:
        c=ws.cell(3,col,h); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
        c.fill=azul; c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=w
    for dd in range(1,dias_mes+1):
        si=(dd-1)//7; col=COL_DIA1+dd-1
        c=ws.cell(3,col,dd); c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(sem_colors[si]); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=4.5
    for col,h,bg,w in [(COL_TF,'T.F','C0504D',4.5),(COL_TA,'T.A','4F81BD',4.5)]:
        c=ws.cell(3,col,h); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
        c.fill=fill(bg); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=w
    ws.row_dimensions[3].height=24

    fila_excel = 4
    for t in trabajadores:
        # Fila 1: estado
        ws.row_dimensions[fila_excel].height=14
        c=ws.cell(fila_excel,COL_N,t['item']); c.font=Font(name='Arial',size=7,bold=True); c.alignment=center(); c.border=border()
        c=ws.cell(fila_excel,COL_NOMBRE,t['nombre_completo']); c.font=Font(name='Arial',size=7,bold=True); c.alignment=left_al(); c.border=border()
        tF=tA=0
        for dd in range(1,dias_mes+1):
            dia_data=t['dias'].get(str(dd),t['dias'].get(dd,{}))
            est=str(dia_data.get('estado','')) if isinstance(dia_data,dict) else str(dia_data)
            col=COL_DIA1+dd-1
            c=ws.cell(fila_excel,col,est if est else '')
            c.font=Font(name='Arial',size=6,bold=bool(est)); c.alignment=center(); c.border=border()
            if est in COLORES: c.fill=fill(COLORES[est])
            if est=='F': tF+=1
            elif est in ['A','T']: tA+=1
        ws.cell(fila_excel,COL_TF,tF).font=Font(name='Arial',size=7,bold=True); ws.cell(fila_excel,COL_TF).alignment=center(); ws.cell(fila_excel,COL_TF).border=border(); ws.cell(fila_excel,COL_TF).fill=fill('FFC7CE')
        ws.cell(fila_excel,COL_TA,tA).font=Font(name='Arial',size=7,bold=True); ws.cell(fila_excel,COL_TA).alignment=center(); ws.cell(fila_excel,COL_TA).border=border(); ws.cell(fila_excel,COL_TA).fill=fill('C6EFCE')

        # Fila 2: detalle lugar/turno/entrada/salida
        fila_excel+=1
        ws.row_dimensions[fila_excel].height=22
        ws.cell(fila_excel,COL_N).fill=gris; ws.cell(fila_excel,COL_N).border=border()
        ws.cell(fila_excel,COL_NOMBRE).fill=gris; ws.cell(fila_excel,COL_NOMBRE).border=border()
        for dd in range(1,dias_mes+1):
            dia_data=t['dias'].get(str(dd),t['dias'].get(dd,{}))
            col=COL_DIA1+dd-1
            if isinstance(dia_data,dict) and dia_data.get('estado','') not in ['','F','DS','FE','FR','DM','LM','LP','LO','LUB','LS','V']:
                lugar=abreviar_lugar(dia_data.get('lugar',''))
                hi=dia_data.get('h_inicio',''); hf=dia_data.get('h_fin','')
                ent=dia_data.get('entrada',''); sal=dia_data.get('salida','')
                turno=f"{hi}-{hf}" if hi and hf else hi or hf
                detalle=f"{lugar}
{turno}
{ent}>{sal}" if lugar else f"{turno}
{ent}>{sal}"
                c=ws.cell(fila_excel,col,detalle)
                c.font=Font(name='Arial',size=5); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
                c.fill=gris; c.border=border()
            else:
                ws.cell(fila_excel,col).fill=gris; ws.cell(fila_excel,col).border=border()
        ws.cell(fila_excel,COL_TF).fill=gris; ws.cell(fila_excel,COL_TF).border=border()
        ws.cell(fila_excel,COL_TA).fill=gris; ws.cell(fila_excel,COL_TA).border=border()
        fila_excel+=1

    # Leyenda
    ws.merge_cells(fila_excel,1,fila_excel,2)
    c=ws.cell(fila_excel,1,'LEYENDA'); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
    c.fill=hdr; c.alignment=center(); ws.row_dimensions[fila_excel].height=13
    for li,ley in enumerate(leyenda):
        r2=fila_excel+1+li; ws.row_dimensions[r2].height=11
        c1=ws.cell(r2,1,ley['cod']); c1.font=Font(name='Arial',bold=True,size=6)
        if ley['cod'] in COLORES: c1.fill=fill(COLORES[ley['cod']])
        c1.alignment=center(); c1.border=border()
        c2=ws.cell(r2,2,ley['desc']); c2.font=Font(name='Arial',size=6)
        if ley['cod'] in COLORES: c2.fill=fill(COLORES[ley['cod']])
        c2.alignment=left_al(); c2.border=border()

    setup_page(ws); ws.print_title_rows='1:3'
    return wb


@app.route('/excel', methods=['POST','GET'])
def generar_excel():
    try:
        if request.method=='POST': data=request.json
        else: data=json.loads(request.args.get('data','{}'))
        tipo=data.get('tipo','mensual')
        trabajadores=data.get('trabajadores',[])
        dias_mes=data.get('dias_mes',31)
        mes_nom=data.get('mes_nom','MES')
        anio=data.get('anio',2026)
        leyenda=data.get('leyenda',[])
        if tipo=='firmas':
            wb=generar_firmas(trabajadores,dias_mes,mes_nom,anio,leyenda)
            fname=f"Reporte_Firmas_{mes_nom}_{anio}.xlsx"
        elif tipo=='tardanzas':
            wb=generar_tardanzas(trabajadores,dias_mes,mes_nom,anio,leyenda)
            fname=f"Reporte_Tardanzas_{mes_nom}_{anio}.xlsx"
        elif tipo=='detalle':
            wb=generar_detalle(trabajadores,dias_mes,mes_nom,anio,leyenda)
            fname=f"Reporte_Detalle_{mes_nom}_{anio}.xlsx"
        else:
            wb=generar_mensual(trabajadores,dias_mes,mes_nom,anio,leyenda)
            fname=f"Reporte_Asistencia_{mes_nom}_{anio}.xlsx"
        out=io.BytesIO(); wb.save(out); out.seek(0)
        return send_file(out,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',as_attachment=True,download_name=fname)
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/')
def home(): return jsonify({'status':'ok','message':'Servidor Excel SUTRAN UDLL'})

if __name__=='__main__': app.run(debug=False,host='0.0.0.0',port=10000)
