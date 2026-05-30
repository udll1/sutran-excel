from flask import Flask, request, send_file, jsonify
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
import io, json

app = Flask(__name__)

@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response

@app.route('/excel', methods=['OPTIONS'])
def options_excel():
    return '', 204

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
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TC)
    c = ws.cell(1,1,f"CONTROL DE ASISTENCIA - SUTRAN UDLL - {mes_nom} {anio}")
    c.font=Font(name='Arial',bold=True,size=10,color='FFFFFF'); c.fill=hdr; c.alignment=center()
    ws.row_dimensions[1].height=20

    # Fila 2 - Semanas
    for col in range(1,5): ws.cell(2,col).fill=hdr
    for si,(s1,s2) in enumerate(semanas):
        c1=COL_DIA1+s1-1; c2=COL_DIA1+s2-1
        if c1<c2: ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
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
    ws.merge_cells(start_row=lr, start_column=1, end_row=lr, end_column=2)
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
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TC)
    c=ws.cell(1,1,f"CONTROL DE ASISTENCIA DEL PERSONAL SUTRAN REGION LA LIBERTAD - {mes_nom} {anio}")
    c.font=Font(name='Arial',bold=True,size=9,color='FFFFFF'); c.fill=hdr; c.alignment=center()
    ws.row_dimensions[1].height=20

    # Fila 2
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=TC)
    c=ws.cell(2,1,"GERENCIA DE ARTICULACION TERRITORIAL")
    c.font=Font(name='Arial',bold=True,size=8,color='FFFFFF'); c.fill=azul; c.alignment=center()
    ws.row_dimensions[2].height=14

    # Fila 3 - Semanas
    for col in [COL_N,COL_NOMBRE]: ws.cell(3,col).fill=hdr
    for si,(s1,s2) in enumerate(semanas):
        c1=COL_DIA1+s1-1; c2=COL_DIA1+s2-1
        if c1<c2: ws.merge_cells(start_row=3, start_column=c1, end_row=3, end_column=c2)
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
    ws.merge_cells(start_row=lr, start_column=1, end_row=lr, end_column=2)
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

    # Calcular minutos tardanza por trabajador y dia
    trab_con_tard = []
    for t in trabajadores:
        tard_dias = {}
        total_min = 0
        for dd in range(1, dias_mes+1):
            dia_data = t['dias'].get(str(dd), t['dias'].get(dd, {}))
            if isinstance(dia_data, dict):
                est = str(dia_data.get('estado',''))
                entrada = str(dia_data.get('entrada',''))
                h_inicio = str(dia_data.get('h_inicio',''))
                if est == 'T' and entrada and h_inicio:
                    try:
                        pe = entrada.split(':'); pi = h_inicio.split(':')
                        min_e = int(pe[0])*60 + int(pe[1])
                        min_i = int(pi[0])*60 + int(pi[1])
                        min_tard = max(0, min_e - min_i)
                        if min_tard > 0:
                            tard_dias[dd] = min_tard
                            total_min += min_tard
                    except: pass
        if total_min > 0:
            trab_con_tard.append({
                'item': len(trab_con_tard)+1,
                'nombre_completo': t['nombre_completo'],
                'dni': t.get('dni',''),
                'cargo': t.get('cargo',''),
                'tard_dias': tard_dias,
                'total_min': total_min
            })

    COL_N=1; COL_DNI=2; COL_NOMBRE=3; COL_CARGO=4
    COL_DIA1=5; COL_TOTAL=COL_DIA1+dias_mes
    TC=COL_TOTAL

    hdr=fill("1F3864"); azul=fill("2E75B6")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TC)
    c=ws.cell(1,1,f"REPORTE DE TARDANZAS - SUTRAN UDLL - {mes_nom} {anio}")
    c.font=Font(name='Arial',bold=True,size=10,color='FFFFFF'); c.fill=hdr; c.alignment=center()
    ws.row_dimensions[1].height=20

    for col in range(1,5): ws.cell(2,col).fill=hdr
    for si,(s1,s2) in enumerate(semanas):
        c1=COL_DIA1+s1-1; c2=COL_DIA1+s2-1
        if c1<c2: ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        c=ws.cell(2,c1,f"SEMANA {si+1}")
        c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(sem_colors[si]); c.alignment=center()
    ws.cell(2,COL_TOTAL).fill=hdr
    ws.row_dimensions[2].height=12

    for col,h,w in [(COL_N,'N',4),(COL_DNI,'DNI',10),(COL_NOMBRE,'APELLIDOS Y NOMBRES',35),(COL_CARGO,'CARGO',18)]:
        c=ws.cell(3,col,h); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
        c.fill=azul; c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=w
    for dd in range(1,dias_mes+1):
        si=(dd-1)//7; col=COL_DIA1+dd-1
        c=ws.cell(3,col,dd); c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(sem_colors[si]); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=4.5
    c=ws.cell(3,COL_TOTAL,'TOTAL\nMIN'); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
    c.fill=fill('E36C09'); c.alignment=center(); c.border=border()
    ws.column_dimensions[get_column_letter(COL_TOTAL)].width=8
    ws.row_dimensions[3].height=24

    if not trab_con_tard:
        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=TC)
        c=ws.cell(4,1,'Sin tardanzas registradas en el mes')
        c.font=Font(name='Arial',italic=True,size=9,color='6B7280'); c.alignment=center()
        ws.row_dimensions[4].height=20
    else:
        for ti,t in enumerate(trab_con_tard):
            r=ti+4; ws.row_dimensions[r].height=14
            for col,val,al in [(COL_N,t['item'],center()),(COL_DNI,t['dni'],center()),(COL_NOMBRE,t['nombre_completo'],left_al()),(COL_CARGO,t['cargo'],center())]:
                c=ws.cell(r,col,val); c.font=Font(name='Arial',size=8); c.alignment=al; c.border=border()
            for dd in range(1,dias_mes+1):
                col=COL_DIA1+dd-1
                if dd in t['tard_dias']:
                    c=ws.cell(r,col,f"+{t['tard_dias'][dd]}")
                    c.font=Font(name='Arial',size=7,bold=True,color='9C5700')
                    c.fill=fill('FFEB9C'); c.alignment=center(); c.border=border()
                else:
                    ws.cell(r,col).border=border()
            c=ws.cell(r,COL_TOTAL,t['total_min'])
            c.font=Font(name='Arial',size=9,bold=True,color='9C5700')
            c.fill=fill('FFB830'); c.alignment=center(); c.border=border()

    setup_page(ws); ws.print_title_rows='1:3'
    return wb



def generar_detalle(trabajadores, dias_mes, mes_nom, anio, leyenda):
    wb = Workbook()
    ws = wb.active
    ws.title = f"DETALLE {mes_nom} {anio}"
    semanas = get_semanas(dias_mes)
    sem_colors = ["1F4E79","2E75B6","1F4E79","2E75B6","1F4E79"]

    COL_N=1; COL_NOMBRE=2; COL_DIA1=3
    TC = COL_DIA1 + dias_mes - 1

    hdr=fill("1F3864"); azul=fill("2E75B6"); gris=fill("F2F2F2")

    # Titulo
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=TC)
    c=ws.cell(1,1,f"DETALLE DE ASISTENCIA - SUTRAN UDLL - {mes_nom} {anio}")
    c.font=Font(name='Arial',bold=True,size=10,color='FFFFFF'); c.fill=hdr; c.alignment=center()
    ws.row_dimensions[1].height=20

    # Semanas
    for col in [COL_N,COL_NOMBRE]: ws.cell(2,col).fill=hdr
    for si,(s1,s2) in enumerate(semanas):
        c1=COL_DIA1+s1-1; c2=COL_DIA1+s2-1
        if c1<c2: ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        c=ws.cell(2,c1,f"SEMANA {si+1}")
        c.font=Font(name='Arial',bold=True,size=6,color='FFFFFF')
        c.fill=fill(sem_colors[si]); c.alignment=center()
    ws.row_dimensions[2].height=12

    # Encabezados
    for col,h,w in [(COL_N,'N',4),(COL_NOMBRE,'APELLIDOS Y NOMBRES',32)]:
        c=ws.cell(3,col,h); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
        c.fill=azul; c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=w
    for dd in range(1,dias_mes+1):
        si=(dd-1)//7; col=COL_DIA1+dd-1
        c=ws.cell(3,col,dd); c.font=Font(name='Arial',bold=True,size=7,color='FFFFFF')
        c.fill=fill(sem_colors[si]); c.alignment=center(); c.border=border()
        ws.column_dimensions[get_column_letter(col)].width=7
    ws.row_dimensions[3].height=18

    # Estados que NO requieren detalle (no asistio)
    sin_detalle = ['F','DS','FE','FR','DM','LM','LP','LO','LUB','LS','V','CP','CPSS','CO','REVISION']

    fila = 4
    for t in trabajadores:
        ws.row_dimensions[fila].height = 38
        c=ws.cell(fila,COL_N,t['item']); c.font=Font(name='Arial',size=8,bold=True); c.alignment=center(); c.border=border()
        c=ws.cell(fila,COL_NOMBRE,t['nombre_completo']); c.font=Font(name='Arial',size=8,bold=True); c.alignment=left_al(); c.border=border()

        for dd in range(1,dias_mes+1):
            col = COL_DIA1+dd-1
            dia_data = t['dias'].get(str(dd), t['dias'].get(dd, {}))
            
            if isinstance(dia_data, dict):
                est = str(dia_data.get('estado',''))
                lugar = abreviar_lugar(dia_data.get('lugar',''))
                hi = str(dia_data.get('h_inicio',''))
                hf = str(dia_data.get('h_fin',''))
                ent = str(dia_data.get('entrada',''))
                sal = str(dia_data.get('salida',''))
            else:
                est = str(dia_data); lugar=''; hi=''; hf=''; ent=''; sal=''

            if not est:
                ws.cell(fila,col).border=border()
                continue

            if est in sin_detalle:
                # No asistio - mostrar solo la nomenclatura centrada
                c = ws.cell(fila,col,est)
                c.font=Font(name='Arial',size=9,bold=True)
                c.alignment=center(); c.border=border()
                if est in COLORES: c.fill=fill(COLORES[est])
            else:
                # Asistio (A o T) - mostrar detalle sin la letra
                turno = f"{hi}-{hf}" if hi and hf else (hi or hf)
                ent_sal = f"{ent}>{sal}" if ent and sal else (ent or sal)
                partes = []
                if lugar: partes.append(lugar)
                if turno: partes.append(turno)
                if ent_sal: partes.append(ent_sal)
                detalle = "\n".join(partes)
                c = ws.cell(fila,col,detalle)
                color_font = '9C5700' if est == 'T' else '000000'
                c.font=Font(name='Arial',size=6,bold=False,color=color_font)
                c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
                if est == 'T': c.fill=fill('FFEB9C')
                elif est == 'A': c.fill=fill('E8F5E9')
                c.border=border()
        fila += 1

    # Leyenda
    lr = fila + 1
    ws.merge_cells(start_row=lr, start_column=1, end_row=lr, end_column=2)
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
        import traceback
        traceback.print_exc()
        return jsonify({'error':str(e), 'trace':traceback.format_exc()}),500

@app.route('/')
def home(): return jsonify({'status':'ok','message':'Servidor Excel SUTRAN UDLL'})

if __name__=='__main__': app.run(debug=False,host='0.0.0.0',port=10000)
