from flask import Flask, request, send_file, jsonify
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
import io
import json

app = Flask(__name__)

# Colores por estado
COLORES = {
    "A":    "C6EFCE",
    "T":    "FFEB9C",
    "F":    "FFC7CE",
    "DS":   "D9D9D9",
    "CP":   "FFE6CC",
    "CPSS": "FFE6CC",
    "FE":   "E1D5E7",
    "FR":   "E1D5E7",
    "C":    "D5E8D4",
    "CO":   "FFF2CC",
    "DM":   "DAE8FC",
    "LM":   "DAE8FC",
    "LP":   "DAE8FC",
    "LO":   "DAE8FC",
    "LUB":  "DAE8FC",
    "LS":   "DAE8FC",
    "V":    "DAE8FC",
    "REVISION": "FCE4D6"
}

def get_fill(estado):
    color = COLORES.get(estado)
    if color:
        return PatternFill('solid', start_color=color)
    return None

def thin_border():
    t = Side(style='thin', color='AAAAAA')
    return Border(left=t, right=t, top=t, bottom=t)

def thick_side(side='left'):
    thick = Side(style='medium', color='000000')
    thin = Side(style='thin', color='AAAAAA')
    if side == 'left':
        return Border(left=thick, right=thin, top=thin, bottom=thin)
    return Border(left=thin, right=thick, top=thin, bottom=thin)

def generar_excel_mensual(trabajadores, dias_mes, mes_nom, anio, leyenda):
    wb = Workbook()
    ws = wb.active
    ws.title = f"ASISTENCIA {mes_nom} {anio}"

    # Semanas
    semanas = []
    d = 1
    while d <= dias_mes:
        semanas.append((d, min(d+6, dias_mes)))
        d = min(d+6, dias_mes) + 1

    sem_colors = ["1F4E79","2E75B6","1F4E79","2E75B6","1F4E79"]

    COL_N = 1
    COL_DNI = 2
    COL_NOMBRE = 3
    COL_CARGO = 4
    COL_DIA1 = 5
    COL_TFALTAS = COL_DIA1 + dias_mes
    COL_TTARD = COL_TFALTAS + 1
    COL_TASIST = COL_TTARD + 1
    COL_TDS = COL_TASIST + 1
    total_cols = COL_TDS

    hdr_fill = PatternFill('solid', start_color='1F3864')
    hdr_font = Font(name='Arial', bold=True, size=7, color='FFFFFF')
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center')

    # Fila 1 - Titulo
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c = ws.cell(1, 1, f"CONTROL DE ASISTENCIA - SUTRAN UDLL - {mes_nom} {anio}")
    c.font = Font(name='Arial', bold=True, size=10, color='FFFFFF')
    c.fill = hdr_fill
    c.alignment = center
    ws.row_dimensions[1].height = 20

    # Fila 2 - Semanas
    for col in range(1, 5):
        ws.cell(2, col).fill = hdr_fill
    for si, (s_ini, s_fin) in enumerate(semanas):
        c1 = COL_DIA1 + s_ini - 1
        c2 = COL_DIA1 + s_fin - 1
        if c1 < c2:
            ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        c = ws.cell(2, c1, f"SEMANA {si+1}")
        c.font = Font(name='Arial', bold=True, size=6, color='FFFFFF')
        c.fill = PatternFill('solid', start_color=sem_colors[si])
        c.alignment = center
    for col in range(COL_TFALTAS, total_cols+1):
        ws.cell(2, col).fill = hdr_fill
    ws.row_dimensions[2].height = 13

    # Fila 3 - Encabezados
    for col, h in [(COL_N,'N'),(COL_DNI,'DNI'),(COL_NOMBRE,'APELLIDOS Y NOMBRES'),(COL_CARGO,'CARGO')]:
        c = ws.cell(3, col, h)
        c.font = hdr_font; c.fill = PatternFill('solid', start_color='2E75B6')
        c.alignment = center; c.border = thin_border()

    for dd in range(1, dias_mes+1):
        si = (dd-1)//7
        c = ws.cell(3, COL_DIA1+dd-1, dd)
        c.font = Font(name='Arial', bold=True, size=6, color='FFFFFF')
        c.fill = PatternFill('solid', start_color=sem_colors[si])
        c.alignment = center; c.border = thin_border()

    for col, h, bg in [(COL_TFALTAS,'T.F','C0504D'),(COL_TTARD,'T.T','E36C09'),
                        (COL_TASIST,'T.A','4F81BD'),(COL_TDS,'T.DS','2E75B6')]:
        c = ws.cell(3, col, h)
        c.font = Font(name='Arial', bold=True, size=7, color='FFFFFF')
        c.fill = PatternFill('solid', start_color=bg)
        c.alignment = center; c.border = thin_border()
    ws.row_dimensions[3].height = 25

    # Datos
    for ti, t in enumerate(trabajadores):
        r = ti + 4
        ws.row_dimensions[r].height = 14
        for col, val, al in [(COL_N, t['item'], center),(COL_DNI, t['dni'], center),
                              (COL_NOMBRE, t['nombre_completo'], left),(COL_CARGO, t['cargo'], center)]:
            c = ws.cell(r, col, val)
            c.font = Font(name='Arial', size=7)
            c.alignment = al; c.border = thin_border()

        tF = tT = tA = tDS = 0
        for dd in range(1, dias_mes+1):
            est = t['dias'].get(str(dd), t['dias'].get(dd, ''))
            c = ws.cell(r, COL_DIA1+dd-1, est)
            c.font = Font(name='Arial', size=6, bold=bool(est))
            c.alignment = center; c.border = thin_border()
            fill = get_fill(est)
            if fill: c.fill = fill
            if est == 'F': tF += 1
            elif est == 'T': tT += 1
            elif est == 'A': tA += 1
            elif est == 'DS': tDS += 1

        for col, val, bg in [(COL_TFALTAS, tF,'FFC7CE'),(COL_TTARD, tT,'FFEB9C'),
                              (COL_TASIST, tA,'C6EFCE'),(COL_TDS, tDS,'D9D9D9')]:
            c = ws.cell(r, col, val)
            c.font = Font(name='Arial', size=7, bold=True)
            c.alignment = center; c.border = thin_border()
            c.fill = PatternFill('solid', start_color=bg)

    # Bordes generales
    ws.getRange = None  # no aplica

    # Anchos de columna
    ws.column_dimensions[get_column_letter(COL_N)].width = 4
    ws.column_dimensions[get_column_letter(COL_DNI)].width = 10
    ws.column_dimensions[get_column_letter(COL_NOMBRE)].width = 28
    ws.column_dimensions[get_column_letter(COL_CARGO)].width = 14
    for dd in range(1, dias_mes+1):
        ws.column_dimensions[get_column_letter(COL_DIA1+dd-1)].width = 3.2
    for col in [COL_TFALTAS, COL_TTARD, COL_TASIST, COL_TDS]:
        ws.column_dimensions[get_column_letter(col)].width = 4.5

    # Leyenda
    lr = len(trabajadores) + 5
    ws.merge_cells(start_row=lr, start_column=1, end_row=lr, end_column=2)
    c = ws.cell(lr, 1, 'LEYENDA')
    c.font = Font(name='Arial', bold=True, size=7, color='FFFFFF')
    c.fill = hdr_fill; c.alignment = center
    ws.row_dimensions[lr].height = 13

    for li, ley in enumerate(leyenda):
        r = lr + 1 + li
        ws.row_dimensions[r].height = 11
        c1 = ws.cell(r, 1, ley['cod'])
        c1.font = Font(name='Arial', bold=True, size=6)
        fill = get_fill(ley['cod'])
        if fill: c1.fill = fill
        c1.alignment = center; c1.border = thin_border()
        c2 = ws.cell(r, 2, ley['desc'])
        c2.font = Font(name='Arial', size=6)
        if fill: c2.fill = fill
        c2.alignment = left; c2.border = thin_border()

    # Configuracion de pagina A4 horizontal
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.4, bottom=0.4)
    ws.print_title_rows = '1:3'

    return wb


def generar_excel_firmas(trabajadores, dias_mes, mes_nom, anio, leyenda):
    wb = Workbook()
    ws = wb.active
    ws.title = f"FIRMAS {mes_nom} {anio}"

    semanas = []
    d = 1
    while d <= dias_mes:
        semanas.append((d, min(d+6, dias_mes)))
        d = min(d+6, dias_mes) + 1
    sem_colors = ["1F4E79","2E75B6","1F4E79","2E75B6","1F4E79"]

    COL_N = 1
    COL_NOMBRE = 2
    COL_DIA1 = 3
    COL_TFALTAS = COL_DIA1 + dias_mes
    COL_TTARD = COL_TFALTAS + 1
    COL_TASIST = COL_TTARD + 1
    COL_TDIAS = COL_TASIST + 1
    COL_FIRMA = COL_TDIAS + 1
    total_cols = COL_FIRMA

    hdr_fill = PatternFill('solid', start_color='1F3864')
    hdr_font = Font(name='Arial', bold=True, size=7, color='FFFFFF')
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center')

    # Fila 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c = ws.cell(1, 1, f"CONTROL DE ASISTENCIA DEL PERSONAL SUTRAN REGION LA LIBERTAD - {mes_nom} {anio}")
    c.font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
    c.fill = hdr_fill; c.alignment = center
    ws.row_dimensions[1].height = 20

    # Fila 2
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_cols)
    c = ws.cell(2, 1, "GERENCIA DE ARTICULACION TERRITORIAL")
    c.font = Font(name='Arial', bold=True, size=8, color='FFFFFF')
    c.fill = PatternFill('solid', start_color='2E75B6'); c.alignment = center
    ws.row_dimensions[2].height = 14

    # Fila 3 - Semanas
    for col in [COL_N, COL_NOMBRE]:
        ws.cell(3, col).fill = hdr_fill
    for si, (s_ini, s_fin) in enumerate(semanas):
        c1 = COL_DIA1 + s_ini - 1
        c2 = COL_DIA1 + s_fin - 1
        if c1 < c2:
            ws.merge_cells(start_row=3, start_column=c1, end_row=3, end_column=c2)
        c = ws.cell(3, c1, f"SEM {si+1}")
        c.font = Font(name='Arial', bold=True, size=6, color='FFFFFF')
        c.fill = PatternFill('solid', start_color=sem_colors[si])
        c.alignment = center
    for col in range(COL_TFALTAS, total_cols+1):
        ws.cell(3, col).fill = hdr_fill
    ws.row_dimensions[3].height = 12

    # Fila 4 - Encabezados
    for col, h in [(COL_N,'N'),(COL_NOMBRE,'APELLIDOS Y NOMBRES')]:
        c = ws.cell(4, col, h)
        c.font = hdr_font; c.fill = PatternFill('solid', start_color='2E75B6')
        c.alignment = center; c.border = thin_border()

    for dd in range(1, dias_mes+1):
        si = (dd-1)//7
        c = ws.cell(4, COL_DIA1+dd-1, dd)
        c.font = Font(name='Arial', bold=True, size=6, color='FFFFFF')
        c.fill = PatternFill('solid', start_color=sem_colors[si])
        c.alignment = center; c.border = thin_border()

    for col, h, bg in [(COL_TFALTAS,'T.FALTAS','C0504D'),(COL_TTARD,'T.TARD','E36C09'),
                        (COL_TASIST,'T.ASIST','4F81BD'),(COL_TDIAS,'T.DIAS','2E75B6'),
                        (COL_FIRMA,'FIRMA','2E75B6')]:
        c = ws.cell(4, col, h)
        c.font = Font(name='Arial', bold=True, size=6, color='FFFFFF')
        c.fill = PatternFill('solid', start_color=bg)
        c.alignment = center; c.border = thin_border()
    ws.row_dimensions[4].height = 25

    # Datos
    for ti, t in enumerate(trabajadores):
        r = ti + 5
        ws.row_dimensions[r].height = 16
        for col, val, al in [(COL_N, t['item'], center),(COL_NOMBRE, t['nombre_completo'], left)]:
            c = ws.cell(r, col, val)
            c.font = Font(name='Arial', size=7)
            c.alignment = al; c.border = thin_border()

        tF = tT = tA = tD = 0
        for dd in range(1, dias_mes+1):
            est = t['dias'].get(str(dd), t['dias'].get(dd, ''))
            c = ws.cell(r, COL_DIA1+dd-1, est)
            c.font = Font(name='Arial', size=6, bold=bool(est))
            c.alignment = center; c.border = thin_border()
            fill = get_fill(est)
            if fill: c.fill = fill
            if est == 'F': tF += 1
            elif est == 'T': tT += 1
            elif est == 'A': tA += 1; tD += 1

        for col, val, bg in [(COL_TFALTAS,tF,'FFC7CE'),(COL_TTARD,tT,'FFEB9C'),
                              (COL_TASIST,tA,'C6EFCE'),(COL_TDIAS,tD,None)]:
            c = ws.cell(r, col, val)
            c.font = Font(name='Arial', size=7, bold=True)
            c.alignment = center; c.border = thin_border()
            if bg: c.fill = PatternFill('solid', start_color=bg)
        ws.cell(r, COL_FIRMA).border = thin_border()

    # Anchos
    ws.column_dimensions[get_column_letter(COL_N)].width = 4
    ws.column_dimensions[get_column_letter(COL_NOMBRE)].width = 30
    for dd in range(1, dias_mes+1):
        ws.column_dimensions[get_column_letter(COL_DIA1+dd-1)].width = 3.2
    for col in [COL_TFALTAS, COL_TTARD, COL_TASIST, COL_TDIAS]:
        ws.column_dimensions[get_column_letter(col)].width = 5
    ws.column_dimensions[get_column_letter(COL_FIRMA)].width = 18

    # Leyenda
    lr = len(trabajadores) + 6
    ws.merge_cells(start_row=lr, start_column=1, end_row=lr, end_column=2)
    c = ws.cell(lr, 1, 'LEYENDA')
    c.font = Font(name='Arial', bold=True, size=7, color='FFFFFF')
    c.fill = hdr_fill; c.alignment = center
    ws.row_dimensions[lr].height = 13

    for li, ley in enumerate(leyenda):
        r = lr + 1 + li
        ws.row_dimensions[r].height = 11
        c1 = ws.cell(r, 1, ley['cod'])
        c1.font = Font(name='Arial', bold=True, size=6)
        fill = get_fill(ley['cod'])
        if fill: c1.fill = fill
        c1.alignment = center; c1.border = thin_border()
        c2 = ws.cell(r, 2, ley['desc'])
        c2.font = Font(name='Arial', size=6)
        if fill: c2.fill = fill
        c2.alignment = left; c2.border = thin_border()

    # Configuracion A4 horizontal
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.4, bottom=0.4)
    ws.print_title_rows = '1:4'

    return wb


@app.route('/excel', methods=['POST', 'GET'])
def generar_excel():
    try:
        if request.method == 'POST':
            data = request.json
        else:
            data = json.loads(request.args.get('data', '{}'))

        tipo = data.get('tipo', 'mensual')
        trabajadores = data.get('trabajadores', [])
        dias_mes = data.get('dias_mes', 31)
        mes_nom = data.get('mes_nom', 'MES')
        anio = data.get('anio', 2026)
        leyenda = data.get('leyenda', [])

        if tipo == 'firmas':
            wb = generar_excel_firmas(trabajadores, dias_mes, mes_nom, anio, leyenda)
            filename = f"Reporte_Firmas_{mes_nom}_{anio}.xlsx"
        else:
            wb = generar_excel_mensual(trabajadores, dias_mes, mes_nom, anio, leyenda)
            filename = f"Reporte_Asistencia_{mes_nom}_{anio}.xlsx"

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/')
def home():
    return jsonify({'status': 'ok', 'message': 'Servidor Excel SUTRAN UDLL'})


if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=10000)
