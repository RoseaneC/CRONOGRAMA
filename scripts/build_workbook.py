# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from model import (PS_SCHEDULE, SCENARIOS, RECOMMENDED, HOLIDAYS, RECESSO_START, RECESSO_END,
                    build_og_schedule, build_ct_schedule, build_ead_schedule, build_ci_schedule,
                    CT_CATALOG, parse_iso)
from datetime import date, timedelta
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule
from openpyxl.comments import Comment

# ============================================================================
# PALETA / ESTILOS
# ============================================================================
NAVY   = "1F3864"
BLUE   = "2E5395"
LBLUE  = "D9E2F3"
TEAL   = "0F6B5C"
LTEAL  = "D8EEE9"
GOLD   = "BF8F00"
LGOLD  = "FCE9B4"
RED    = "C0392B"
LRED   = "F8D7D3"
GREEN  = "2E7D32"
LGREEN = "DCEDC8"
GREY   = "5A5A5A"
LGREY  = "EDEDED"
WHITE  = "FFFFFF"

STAGE_COLORS = {
    "PS":  "F4B183",  # laranja claro
    "DH":  "9DC3E6",  # azul claro
    "OG":  "A9D18E",  # verde claro
    "CT":  "FFD966",  # amarelo
    "EAD": "B4A7D6",  # roxo claro
    "CI":  "F1948A",  # coral
    "REC": "D9D9D9",  # recesso/feriado
}

FONT_TITLE = Font(name="Calibri", size=16, bold=True, color=WHITE)
FONT_SUB   = Font(name="Calibri", size=11, italic=True, color=GREY)
FONT_H1    = Font(name="Calibri", size=12, bold=True, color=WHITE)
FONT_H2    = Font(name="Calibri", size=10, bold=True, color=NAVY)
FONT_BODY  = Font(name="Calibri", size=10, color="000000")
FONT_BOLD  = Font(name="Calibri", size=10, bold=True, color="000000")
FONT_NOTE  = Font(name="Calibri", size=9, italic=True, color=GREY)

FILL_TITLE = PatternFill("solid", fgColor=NAVY)
FILL_H1    = PatternFill("solid", fgColor=BLUE)
FILL_H2    = PatternFill("solid", fgColor=LBLUE)
FILL_EDIT  = PatternFill("solid", fgColor=LGOLD)
FILL_WARN  = PatternFill("solid", fgColor=LRED)
FILL_OK    = PatternFill("solid", fgColor=LGREEN)
FILL_INFO  = PatternFill("solid", fgColor=LTEAL)
FILL_GREY  = PatternFill("solid", fgColor=LGREY)

THIN = Side(style="thin", color="B7B7B7")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CLASS_COLORS = {
    "Confirmada": LGREEN,
    "Premissa provisória": LGOLD,
    "Divergência entre documentos": LRED,
    "Informação pendente de validação": "D9D9D9",
    "Decisão recomendada": LTEAL,
}

def style_title(ws, cell_range, text, height=28):
    ws.merge_cells(cell_range)
    c = ws[cell_range.split(":")[0]]
    c.value = text
    c.font = FONT_TITLE
    c.fill = FILL_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[c.row].height = height

def style_subtitle(ws, cell_range, text):
    ws.merge_cells(cell_range)
    c = ws[cell_range.split(":")[0]]
    c.value = text
    c.font = FONT_SUB
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

def header_row(ws, row, col_start, headers, fill=FILL_H1, font=FONT_H1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col_start+i, value=h)
        c.font = font
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_ALL
    ws.row_dimensions[row].height = 30

def body_cell(ws, row, col, value, fill=None, font=FONT_BODY, align="left", wrap=False, numfmt=None, bold=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = FONT_BOLD if bold else font
    if fill: c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    c.border = BORDER_ALL
    if numfmt: c.number_format = numfmt
    return c

def classification_cell(ws, row, col, classe):
    color = CLASS_COLORS.get(classe, WHITE)
    c = ws.cell(row=row, column=col, value=classe)
    c.font = Font(name="Calibri", size=9, bold=True)
    c.fill = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER_ALL
    return c

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def freeze(ws, cell):
    ws.freeze_panes = cell

DATE_FMT = "DD/MM/YYYY"

wb = Workbook()
wb.remove(wb.active)

print("Workbook base criado. Prosseguindo com as abas...")

# ============================================================================
# 00_LEIA-ME
# ============================================================================
ws = wb.create_sheet("00_LEIA-ME")
set_col_widths(ws, [3, 30, 90, 3])
style_title(ws, "B2:C2", "PARTIU! APRENDER E EMPREENDER — CRONOGRAMA PAULÍNIA (NÚCLEO 2)")
style_subtitle(ws, "B3:C3", "Instituto da Criança · Sebrae · Petrobras — Material de apoio à decisão e negociação | Gerado em 21/07/2026")

r = 5
ws.cell(r,2,"O QUE É ESTE ARQUIVO").font = FONT_H2
r += 1
txt = ("Este arquivo consolida e cruza 4 documentos-fonte recebidos sobre o projeto em Paulínia: (1) a apresentação macro do "
"projeto com metas contratuais; (2) a planilha operacional de turmas/cargas horárias; (3) uma tentativa anterior de "
"cronograma da coordenação; e (4) uma simulação inicial de 4 cenários de salas/frequência do Desenvolvimento Humano (DH). "
"Os 4 arquivos originais NÃO foram alterados. Onde os documentos divergem, a divergência foi preservada e classificada — "
"nenhum número foi escolhido \"por conta própria\" sem sinalização.")
ws.merge_cells(f"B{r}:C{r+2}")
c = ws.cell(r,2,txt); c.font=FONT_BODY; c.alignment=Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[r].height = 60
r += 4

ws.cell(r,2,"COMO CLASSIFICAMOS CADA INFORMAÇÃO").font = FONT_H2
r += 1
classes = [
    ("Confirmada", "Valor igual em pelo menos 2 fontes independentes, sem contradição."),
    ("Premissa provisória", "Valor usado para construir o cronograma na ausência de confirmação formal; pode mudar."),
    ("Divergência entre documentos", "Os documentos-fonte trazem valores diferentes para o mesmo item; nenhum foi descartado."),
    ("Informação pendente de validação", "Dado ausente, incompleto ou com placeholder (ex.: fornecedor \"?\")."),
    ("Decisão recomendada", "Recomendação da análise, com justificativa — não é um fato, é uma sugestão de encaminhamento."),
]
for classe, desc in classes:
    classification_cell(ws, r, 2, classe)
    body_cell(ws, r, 3, desc, wrap=True)
    ws.row_dimensions[r].height = 28
    r += 1
r += 1

ws.cell(r,2,"COMO NAVEGAR").font = FONT_H2
r += 1
nav = [
    ("01_PREMISSAS", "Painel de premissas editáveis (células amarelas). Altere aqui frequência de DH, nº de salas, carga horária etc."),
    ("02_DIVERGENCIAS", "Tabela completa de conflitos entre os 4 documentos, com decisão provisória e responsável pela validação."),
    ("03_FLUXO_DA_TRILHA", "Mapa da trilha PS → DH → OG → CT → EAD → CI, com as regras quantitativas encontradas (e o que é premissa)."),
    ("04_RESUMO_EXECUTIVO", "Visão de 1 página para Sebrae/Petrobras — leitura em menos de 2 minutos."),
    ("05_COMPARATIVO_CENARIOS", "Painel comparativo dos 4 cenários de sala/frequência do DH, com semáforo de viabilidade."),
    ("06 a 09_CENARIO A-D", "Detalhamento turma a turma de cada cenário de DH (datas, sala, cadência)."),
    ("10_CRONOGRAMA_MESTRE", "Linha do tempo completa de Paulínia no cenário recomendado (PS, DH, OG, CT, EAD, CI, recesso, feriados)."),
    ("11_AGENDA_POR_SALA", "Ocupação de Sala 1/2/3 e locais de curso técnico, sem sobreposição."),
    ("12_TURMAS_DH", "Detalhamento de PS 01-30 e DH 01-15 (datas, frequência, carga horária, participantes, predecessora)."),
    ("13_OFICINAS_GESTAO", "Detalhamento das turmas de Oficina de Gestão."),
    ("14_CURSOS_TECNICOS", "Detalhamento das turmas de curso técnico (tema, fornecedor, carga horária, participantes)."),
    ("15_EAD_CONSULTORIAS", "Curso EAD e Consultoria Individual."),
    ("16_RISCOS_E_DECISOES", "Quadro de riscos, divergências e decisões pendentes, com responsável e status."),
    ("17_GANTT_EXECUTIVO", "Gantt visual semanal de toda a trilha (cenário recomendado) com feriados e recesso destacados."),
]
for nome, desc in nav:
    body_cell(ws, r, 2, nome, bold=True, fill=LBLUE)
    body_cell(ws, r, 3, desc, wrap=True)
    ws.row_dimensions[r].height = 26
    r += 1
r += 1

ws.cell(r,2,"REGRAS IMPORTANTES DESTE MODELO").font = FONT_H2
r += 1
regras = [
    "Nenhuma divergência foi corrigida silenciosamente (ex.: 20h x 22h de DH; 10 x 13 turmas de curso técnico).",
    "Toda data calculada indica, na própria linha, qual premissa foi usada para chegar até ela.",
    "As datas são datas reais do Excel (não texto) — podem ser usadas em fórmulas, filtros e classificação.",
    "As abas 05 (comparativo) recalculam automaticamente a partir das abas 06-09 via fórmulas (MIN/MÁXIMO/MENOR).",
    "As grades turma a turma (06-09, 10, 11, 12-15) são geradas a partir das premissas vigentes no momento da montagem. "
    "Se uma premissa estrutural mudar (ex.: nº de salas, frequência do DH), regenere o arquivo para recalcular as grades — "
    "o Excel nativo não recalcula sozinho um algoritmo de alocação de salas.",
]
for reg in regras:
    body_cell(ws, r, 2, "⚠", align="center", fill=LGOLD)
    ws.merge_cells(f"C{r}:C{r}")
    body_cell(ws, r, 3, reg, wrap=True)
    ws.row_dimensions[r].height = 30
    r += 1

freeze(ws, "A5")

# ============================================================================
# 01_PREMISSAS
# ============================================================================
ws = wb.create_sheet("01_PREMISSAS")
set_col_widths(ws, [3, 30, 16, 40, 22, 3])
style_title(ws, "B2:E2", "PREMISSAS DO CRONOGRAMA — CÉLULAS AMARELAS SÃO EDITÁVEIS")
style_subtitle(ws, "B3:E3", "Alterar aqui atualiza automaticamente as datas e indicadores das abas 04 e 05. Grades turma a turma (06-17) exigem regeneração do arquivo.")

r = 5
def premise_block(ws, r, title, rows):
    ws.merge_cells(f"B{r}:E{r}")
    c = ws.cell(r,2,title); c.font=FONT_H1; c.fill=FILL_H1
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 22
    r += 1
    header_row(ws, r, 2, ["Item","Valor","Observação / Classificação","Status"], fill=FILL_H2, font=FONT_H2)
    r += 1
    for item, val, obs, status, editable in rows:
        body_cell(ws, r, 2, item, bold=True)
        vc = body_cell(ws, r, 3, val, align="center")
        if editable:
            vc.fill = FILL_EDIT
        body_cell(ws, r, 4, obs, wrap=True)
        classification_cell(ws, r, 5, status)
        ws.row_dimensions[r].height = 28
        r += 1
    return r + 1

r = premise_block(ws, r, "PALESTRAS DE SENSIBILIZAÇÃO (PS)", [
    ("Data de início", date(2026,8,11), "Confirmado em 3 fontes (planilha operacional, simulação, cronograma anterior)", "Confirmada", True),
    ("Data-alvo de término", date(2026,10,15), "Confirmado em 3 fontes", "Confirmada", True),
    ("Total de turmas (Paulínia)", 30, "Base operacional J3=30; PDF macro Núcleo 2 = 30 turmas", "Confirmada", True),
    ("Frequência", "3x/semana (Ter/Qua/Qui)", "Cadência dedutível da planilha \"Possibilidades\"; concilia início e fim informados", "Premissa provisória", True),
    ("Carga horária por sessão", "2h", "Planilha operacional E3", "Confirmada", True),
])

r = premise_block(ws, r, "DESENVOLVIMENTO HUMANO (DH)", [
    ("Total de turmas (Paulínia)", 15, "Confirmado: planilha operacional J4, PDF macro Núcleo 2, simulação de cenários", "Confirmada", True),
    ("Carga horária total — OPÇÃO USADA NO MODELO", "20h / 10 encontros", "Usada pela simulação de cenários e citada em reunião (Sebrae teria reduzido)", "Premissa provisória", True),
    ("Carga horária total — OPÇÃO ALTERNATIVA (documental)", "22h / 11 encontros", "Planilha operacional (F4=22, G4=11) e carga horária citada no PDF macro (\"22 horas\")", "Divergência entre documentos", False),
    ("Mínimo de participantes/turma", 20, "Planilha operacional H4; PDF usa 20/turma como média de referência", "Confirmada", True),
    ("Máximo de participantes/turma", 40, "Planilha operacional I4", "Confirmada", True),
    ("Regra de formação — OPÇÃO A", "2 PS = 1 DH", "Fecha exatamente com 30 PS → 15 DH; é a regra usada nesta simulação", "Premissa provisória", False),
    ("Regra de formação — OPÇÃO B (anotação divergente)", "4 OS = 1 DH", "\"OS\" não existe na trilha; provável erro de digitação de \"PS\". Mesmo lendo como \"4 PS=1DH\", o resultado (7-8 turmas) não fecha com as 15 turmas confirmadas", "Divergência entre documentos", False),
    ("Frequência semanal a comparar", "2x ou 5x por semana", "Ambas exigidas pelo escopo desta análise (cenários A-D)", "Premissa provisória", True),
    ("Nº de salas a comparar", "2 ou 3", "Ambas exigidas pelo escopo desta análise (cenários A-D)", "Premissa provisória", True),
    ("Gatilho para 1º curso técnico", "Após a 4ª turma de DH concluída", "Planilha operacional, anotação A34", "Premissa provisória", True),
])

r = premise_block(ws, r, "OFICINA DE GESTÃO (OG)", [
    ("Total de turmas — OPÇÃO USADA NO MODELO", 15, "Planilha operacional J5=15 (1 OG por turma de DH concluída)", "Premissa provisória", True),
    ("Total de turmas — LEITURA ALTERNATIVA", "≈7 a 8", "Se a anotação \"2 DH = 1 OG + 1 CT\" (A28) for aplicada literalmente às 15 turmas de DH", "Divergência entre documentos", False),
    ("Carga horária total", "4h", "Planilha operacional F5", "Confirmada", True),
    ("Encontros", 2, "Planilha operacional G5, de 2h cada", "Confirmada", True),
    ("Mínimo/máximo de participantes", "15 / 18", "Planilha operacional H5/I5", "Confirmada", True),
    ("Espaço para a OG no cenário recomendado (B)", "3 espaços dedicados, distintos das 3 salas de DH", "ACHADO DESTA ANÁLISE: no cenário B (3 salas, DH 5x/semana), as 3 salas de DH ficam 100% ocupadas por lotes consecutivos durante toda a fase — não sobra horário livre na mesma sala para a OG. É preciso um espaço adicional.", "Decisão recomendada", False),
])

r = premise_block(ws, r, "CURSOS TÉCNICOS (CT)", [
    ("Total de turmas — OPÇÃO USADA NO MODELO", 10, "Planilha operacional: soma de J6:J12 (7 temas)", "Premissa provisória", True),
    ("Total de turmas — META CONTRATUAL (PDF macro)", 13, "PDF \"Apres. Projeto e Anexos\", Núcleo 2 = 13 turmas = 195 pessoas", "Divergência entre documentos", False),
    ("Diferença a validar", "3 turmas", "Não foi definido neste material qual(is) tema(s) receberia(m) turma adicional para fechar 13", "Informação pendente de validação", False),
    ("Frequência", "5x/semana, 4h/dia", "Planilha operacional (observação) e confirmado no briefing", "Confirmada", True),
    ("Turno preferencial", "Noite", "Confirmado na planilha operacional e no briefing", "Confirmada", True),
    ("Quórum mínimo por turma", "14 a 18 (varia por curso/fornecedor)", "Planilha operacional H6:H12; reunião cita 16 como referência geral", "Divergência entre documentos", False),
])

r = premise_block(ws, r, "EAD E CONSULTORIA INDIVIDUAL", [
    ("EAD — carga horária", "6h / 3 encontros", "Planilha operacional linha 13", "Confirmada", True),
    ("Consultoria Individual — atendimentos por pessoa", 3, "Planilha operacional linha 14 e nota J14", "Confirmada", True),
    ("Consultoria Individual — meta de pessoas (Paulínia)", 150, "Nota J14 (450 atend. / 3) e PDF macro (\"13 turmas = 150 pessoas\")", "Confirmada", True),
    ("Curso Técnico é pré-requisito para Consultoria Individual", "Sim", "PDF macro: \"Curso Técnico é pré-requisito\"", "Confirmada", True),
])

r = premise_block(ws, r, "CALENDÁRIO", [
    ("Feriados considerados", "7/9, 12/10, 2/11, 20/11, 25/12, 1/1", "Calendário municipal de Paulínia 2026 (fonte: paulinia.sp.gov.br/feriados2026)", "Premissa provisória", False),
    ("Recesso institucional — início", date(2026,12,21), "Observado no cronograma anterior da coordenação; não confirmado formalmente", "Informação pendente de validação", True),
    ("Recesso institucional — fim", date(2027,1,8), "Idem", "Informação pendente de validação", True),
])

freeze(ws, "B6")

# ============================================================================
# 03_FLUXO_DA_TRILHA
# ============================================================================
ws = wb.create_sheet("03_FLUXO_DA_TRILHA")
set_col_widths(ws, [3, 20, 4, 20, 4, 20, 4, 20, 4, 20, 4, 20, 3])
style_title(ws, "B2:L2", "FLUXO DA TRILHA — PALESTRA → DH → OFICINA DE GESTÃO → CURSO TÉCNICO → EAD → CONSULTORIA")
style_subtitle(ws, "B3:L3", "Sequência obrigatória de dependências. Regras quantitativas mostradas como premissa/divergência, nunca como fato quando não confirmadas.")

r = 5
stages = [
    ("PALESTRA DE\nSENSIBILIZAÇÃO\n(PS)", STAGE_COLORS["PS"], "30 turmas"),
    ("DESENVOLVIMENTO\nHUMANO\n(DH)", STAGE_COLORS["DH"], "15 turmas"),
    ("OFICINA DE\nGESTÃO\n(OG)", STAGE_COLORS["OG"], "15 turmas*"),
    ("CURSO\nTÉCNICO\n(CT)", STAGE_COLORS["CT"], "10 turmas*"),
    ("CURSO\nEAD", STAGE_COLORS["EAD"], "contínuo"),
    ("CONSULTORIA\nINDIVIDUAL\n(CI)", STAGE_COLORS["CI"], "meta 150 pessoas"),
]
col = 2
for i, (name, color, qty) in enumerate(stages):
    c = ws.cell(r, col, name)
    ws.merge_cells(start_row=r, start_column=col, end_row=r+2, end_column=col+1)
    c.fill = PatternFill("solid", fgColor=color)
    c.font = Font(name="Calibri", size=11, bold=True, color="000000")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for rr in range(r, r+3):
        for cc in range(col, col+2):
            ws.cell(rr, cc).border = BORDER_ALL
            ws.cell(rr, cc).fill = PatternFill("solid", fgColor=color)
    qc = ws.cell(r+3, col, qty)
    ws.merge_cells(start_row=r+3, start_column=col, end_row=r+3, end_column=col+1)
    qc.font = Font(size=9, italic=True, bold=True); qc.alignment=Alignment(horizontal="center")
    if i < len(stages)-1:
        arrow_col = col+2
        ac = ws.cell(r+1, arrow_col, "→")
        ac.font = Font(size=20, bold=True, color=NAVY)
        ac.alignment = Alignment(horizontal="center", vertical="center")
    col += 3
ws.row_dimensions[r].height = 20
ws.row_dimensions[r+1].height = 20
ws.row_dimensions[r+2].height = 20

r2 = r + 5
ws.cell(r2, 2, "* Ver divergência quantitativa nas abas 01 e 02 (OG: 15 x ~7-8; CT: 10 x 13).").font = FONT_NOTE
r2 += 2

header_row(ws, r2, 2, ["Relação quantitativa citada nos documentos","Fecha com os totais confirmados?","Classificação"], fill=FILL_H2, font=FONT_H2)
ws.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=5)
ws.cell(r2,2).value = "Relação quantitativa citada nos documentos"
ws.merge_cells(start_row=r2, start_column=6, end_row=r2, end_column=9)
ws.cell(r2,6).value = "Fecha com os totais confirmados (30 PS / 15 DH)?"
ws.merge_cells(start_row=r2, start_column=10, end_row=r2, end_column=12)
ws.cell(r2,10).value = "Classificação"
for cc in range(2,13):
    ws.cell(r2,cc).font=FONT_H2; ws.cell(r2,cc).fill=FILL_H2; ws.cell(r2,cc).border=BORDER_ALL
    ws.cell(r2,cc).alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[r2].height = 26
r2 += 1

relacoes = [
    ("30 PS → 15 DH  (regra: 2 PS = 1 DH)", "Sim — 30 ÷ 2 = 15, bate exatamente com o total confirmado.", "Confirmada"),
    ("30 PS → 15 DH  (regra: 4 OS = 1 DH)", "Não — 30 ÷ 4 = 7,5, não bate com 15 turmas. Sigla \"OS\" também não existe na trilha (possível erro de digitação de \"PS\").", "Divergência entre documentos"),
    ("4 DH concluídos → 2 turmas de Curso Técnico (~80 pessoas formando 2 turmas de ~15-18)", "Aritmeticamente aproximado (80 pessoas / 2 turmas = 40 cada, acima do quórum de 14-18 por turma); indica que nem todos os 80 formandos migram para o mesmo curso técnico — a real distribuição depende da escolha individual de curso.", "Premissa provisória"),
    ("2 DH = 1 OG + 1 CT", "Não fecha com as 15 turmas de OG citadas na planilha operacional (implicaria ~7-8 OG, não 15).", "Divergência entre documentos"),
    ("15 DH → 15 OG (1 para 1)", "Consistente com o total de 15 turmas de OG da planilha operacional, mas contradiz a regra \"2 DH = 1 OG + 1 CT\" acima.", "Premissa provisória"),
]
for rel, fecha, classe in relacoes:
    ws.merge_cells(start_row=r2, start_column=2, end_row=r2, end_column=5)
    body_cell(ws, r2, 2, rel, wrap=True)
    ws.merge_cells(start_row=r2, start_column=6, end_row=r2, end_column=9)
    body_cell(ws, r2, 6, fecha, wrap=True)
    ws.merge_cells(start_row=r2, start_column=10, end_row=r2, end_column=12)
    classification_cell(ws, r2, 10, classe)
    for cc in range(2,13):
        ws.cell(r2,cc).border = BORDER_ALL
    ws.row_dimensions[r2].height = 55
    r2 += 1

r2 += 1
ws.cell(r2,2,"LEGENDA DE ETAPAS").font = FONT_H2
r2 += 1
for name, color in [("PS - Palestra de Sensibilização", STAGE_COLORS["PS"]), ("DH - Desenvolvimento Humano", STAGE_COLORS["DH"]),
                     ("OG - Oficina de Gestão", STAGE_COLORS["OG"]), ("CT - Curso Técnico", STAGE_COLORS["CT"]),
                     ("EAD - Curso Online", STAGE_COLORS["EAD"]), ("CI - Consultoria Individual", STAGE_COLORS["CI"]),
                     ("Recesso / feriado", STAGE_COLORS["REC"])]:
    body_cell(ws, r2, 2, "", fill=color)
    ws.cell(r2,2).border = BORDER_ALL
    body_cell(ws, r2, 3, name)
    r2 += 1

freeze(ws, "B9")

# ============================================================================
# 06-09 CENARIO A/B/C/D
# ============================================================================
SCENARIO_SHEET_NAMES = {}
for key in ["A","B","C","D"]:
    s = SCENARIOS[key]
    sheet_num = {"A":"06","B":"07","C":"08","D":"09"}[key]
    sheet_name = f"{sheet_num}_CENARIO_{key}"
    SCENARIO_SHEET_NAMES[key] = sheet_name
    ws = wb.create_sheet(sheet_name)
    set_col_widths(ws, [3, 10, 8, 14, 12, 12, 12, 10, 34, 3])
    tag = "★ CENÁRIO RECOMENDADO" if key == RECOMMENDED else ""
    style_title(ws, "B2:I2", f"CENÁRIO {key} — {s['salas']} SALAS | DH {s['freq']}  {tag}")
    style_subtitle(ws, "B3:I3", f"{s['leitura']}  |  Indicação: {s['indicacao']}")

    r = 5
    header_row(ws, r, 2, ["Turma DH","Lote","Sala","Cadência","Início","Término","Encontros","Nº aprox. de\nparticipantes","Observação"])
    r += 1
    first_data_row = r
    for t in s["turmas"]:
        body_cell(ws, r, 2, t["turma"], bold=True)
        body_cell(ws, r, 3, t["lote"], align="center")
        body_cell(ws, r, 4, t["sala"], align="center")
        body_cell(ws, r, 5, t["cadencia"], align="center")
        dc = body_cell(ws, r, 6, t["inicio"], align="center"); dc.number_format = DATE_FMT
        fc = body_cell(ws, r, 7, t["fim"], align="center"); fc.number_format = DATE_FMT
        body_cell(ws, r, 8, 10, align="center")
        body_cell(ws, r, 9, "20 a 40 (premissa)", align="center")
        obs = ""
        if t["fim"] >= RECESSO_START and t["inicio"] <= RECESSO_END:
            obs = "Atravessa o recesso institucional (21/12-08/01) — pendente de confirmação"
        hol_hits = [d for d in HOLIDAYS if t["inicio"] <= d <= t["fim"]]
        if hol_hits:
            obs = (obs + "; " if obs else "") + "Período inclui feriado(s): " + ", ".join(d.strftime("%d/%m") for d in sorted(hol_hits))
        cobs = ws.cell(r, 9); cobs.value = obs if obs else "—"
        cobs.font = FONT_NOTE if not obs else Font(size=9, color=RED, italic=True)
        cobs.alignment = Alignment(wrap_text=True, vertical="center")
        cobs.border = BORDER_ALL
        if obs:
            for cc in range(2,10):
                ws.cell(r,cc).fill = FILL_WARN
        ws.row_dimensions[r].height = 22
        r += 1
    last_data_row = r - 1

    r += 1
    header_row(ws, r, 2, ["Indicador","Resultado","Leitura"], fill=FILL_H2, font=FONT_H2)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=9)
    r += 1
    indic_rows = {}
    def indic(label, formula_or_val, leitura, is_date=True):
        global r
        body_cell(ws, r, 2, label, bold=True)
        c = ws.cell(r, 3, formula_or_val)
        c.font = FONT_BODY; c.border = BORDER_ALL; c.alignment = Alignment(horizontal="center")
        if is_date: c.number_format = DATE_FMT
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=9)
        body_cell(ws, r, 4, leitura, wrap=True)
        indic_rows[label] = r
        ws.row_dimensions[r].height = 24
        r += 1

    rng_ini = f"F{first_data_row}:F{last_data_row}"
    rng_fim = f"G{first_data_row}:G{last_data_row}"
    indic("Início da 1ª turma", f"=MIN({rng_ini})", "Marco de início da fase de Desenvolvimento Humano")
    indic("Conclusão da 4ª turma", f"=SMALL({rng_fim},4)", "Marco para liberar o primeiro lote de Curso Técnico (premissa: após 4 DH)")
    indic("Conclusão da 15ª turma (todas)", f"=MAX({rng_fim})", "Fim estimado da fase de Desenvolvimento Humano")
    indic("Duração da fase DH (dias corridos)", f"={get_column_letter(3)}{indic_rows['Conclusão da 15ª turma (todas)']}-{get_column_letter(3)}{indic_rows['Início da 1ª turma']}+1", "Da 1ª turma iniciada até a 15ª concluída", is_date=False)
    indic("Capacidade simultânea máxima (turmas)", s["cap_simultanea"], f"{s['salas']} sala(s) × interleaving por dia da semana (2 cadências/sala se 2x/semana)", is_date=False)

    freeze(ws, f"B{first_data_row}")
    globals()[f"SCEN_ROWS_{key}"] = (first_data_row, last_data_row, indic_rows)

# ============================================================================
# 05_COMPARATIVO_CENARIOS
# ============================================================================
ws = wb.create_sheet("05_COMPARATIVO_CENARIOS")
set_col_widths(ws, [3, 20, 10, 16, 16, 15, 15, 15, 15, 30, 34, 20, 3])
style_title(ws, "B2:L2", "COMPARATIVO DOS 4 CENÁRIOS DE DESENVOLVIMENTO HUMANO (DH)")
style_subtitle(ws, "B3:L3", "Salas x Frequência semanal — indicadores recalculados por fórmula a partir das abas 06 a 09")

r = 5
headers = ["Cenário","Salas","Frequência\nDH","Capacidade\nsimultânea","Início do\nDH","4ª turma\nconcluída","15ª turma\nconcluída","Duração\nfase DH (dias)","Riscos principais","Vantagens / Desvantagens","Viabilidade","Recomendação"]
header_row(ws, r, 2, headers)
r += 1
scen_data_first_row = r

SCEN_META = {
 "A": {"riscos": "Sobreposição de turmas exige interleaving de dias (Seg/Qua com Ter/Qui) na mesma sala; termina em dezembro, próximo ao recesso.",
       "vant_desv": "Vantagem: usa a capacidade máxima das 3 salas. Desvantagem: ritmo mais lento (2x/semana) atrasa toda a cadeia OG→CT→EAD→CI."},
 "B": {"riscos": "Ritmo intenso (5x/semana) exige disponibilidade constante de facilitadores e das 3 salas em sequência apertada entre lotes.",
       "vant_desv": "Vantagem: conclui a fase DH mais cedo (04/11), maximizando o tempo disponível para OG, CT, EAD e CI antes do fim do ano. Desvantagem: menor margem de faltas/reposição por turma (10 dias úteis corridos)."},
 "C": {"riscos": "Fase DH ultrapassa o recesso institucional e só termina em fevereiro/2027 — compromete seriamente o início dos cursos técnicos e a meta anual.",
       "vant_desv": "Vantagem: nenhuma relevante frente aos demais cenários. Desvantagem: menor capacidade simultânea (4) e maior duração total (169 dias)."},
 "D": {"riscos": "Alta intensidade diária (5x/semana) combinada a apenas 2 salas deixa pouquíssima folga para reposição de encontros perdidos.",
       "vant_desv": "Vantagem: mais rápido que os cenários de 2x/semana. Desvantagem: capacidade simultânea mínima (2) entre os 4 cenários."},
}
VIAB = {"A":"Viável","B":"Altamente viável","C":"Não recomendado","D":"Viável com ressalvas"}
VIAB_COLOR = {"Viável":LGOLD, "Altamente viável":LGREEN, "Não recomendado":LRED, "Viável com ressalvas":LGOLD}

for key in ["A","B","C","D"]:
    sh = SCENARIO_SHEET_NAMES[key]
    first, last, indic_rows = globals()[f"SCEN_ROWS_{key}"]
    s = SCENARIOS[key]
    is_rec = key == RECOMMENDED
    rowfill = LGREEN if is_rec else None
    body_cell(ws, r, 2, f"Cenário {key}" + (" ★" if is_rec else ""), bold=True, fill=rowfill)
    body_cell(ws, r, 3, s["salas"], align="center", fill=rowfill)
    body_cell(ws, r, 4, s["freq"], align="center", fill=rowfill)
    c = ws.cell(r, 5, f"='{sh}'!C{indic_rows['Capacidade simultânea máxima (turmas)']}")
    c.alignment=Alignment(horizontal="center"); c.border=BORDER_ALL
    if rowfill: c.fill = PatternFill("solid", fgColor=rowfill)
    c1 = ws.cell(r, 6, f"='{sh}'!C{indic_rows['Início da 1ª turma']}"); c1.number_format=DATE_FMT
    c2 = ws.cell(r, 7, f"='{sh}'!C{indic_rows['Conclusão da 4ª turma']}"); c2.number_format=DATE_FMT
    c3 = ws.cell(r, 8, f"='{sh}'!C{indic_rows['Conclusão da 15ª turma (todas)']}"); c3.number_format=DATE_FMT
    c4 = ws.cell(r, 9, f"='{sh}'!C{indic_rows['Duração da fase DH (dias corridos)']}")
    for cc in (c1,c2,c3,c4):
        cc.alignment=Alignment(horizontal="center"); cc.border=BORDER_ALL
        if rowfill: cc.fill = PatternFill("solid", fgColor=rowfill)
    body_cell(ws, r, 10, SCEN_META[key]["riscos"], wrap=True, fill=rowfill)
    body_cell(ws, r, 11, SCEN_META[key]["vant_desv"], wrap=True, fill=rowfill)
    vcell = ws.cell(r, 12, VIAB[key])
    vcell.font = Font(bold=True, size=10)
    vcell.fill = PatternFill("solid", fgColor=VIAB_COLOR[VIAB[key]])
    vcell.alignment = Alignment(horizontal="center", vertical="center")
    vcell.border = BORDER_ALL
    ws.row_dimensions[r].height = 70
    r += 1
scen_data_last_row = r - 1

r += 1
ws.merge_cells(f"B{r}:L{r}")
rec_cell = ws.cell(r, 2, f"RECOMENDAÇÃO: Cenário {RECOMMENDED} (3 salas, DH 5x/semana) — conclui a fase de Desenvolvimento Humano em 04/11/2026, liberando o maior tempo possível "
                          "para Oficina de Gestão, Cursos Técnicos, EAD e Consultoria Individual ainda dentro do ano civil, antes do recesso institucional. "
                          "Justificativa: menor duração total da fase (73 dias), maior antecedência da 4ª turma concluída (21/09, a mais cedo entre os 4 cenários) "
                          "e nenhuma sobreposição com o recesso. Ressalva: exige confirmação de disponibilidade das 3 salas e de facilitadores em ritmo diário.")
rec_cell.font = Font(bold=True, size=11, color=WHITE)
rec_cell.fill = PatternFill("solid", fgColor=TEAL)
rec_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
ws.row_dimensions[r].height = 70

freeze(ws, f"B{scen_data_first_row}")

# ============================================================================
# Dados derivados para o cenario recomendado (usados nas abas 10-17)
# ============================================================================
REC = SCENARIOS[RECOMMENDED]
OG_SCHEDULE = build_og_schedule(REC)
CT_SCHEDULE, CT_GATE = build_ct_schedule(REC)
EAD_SCHEDULE = build_ead_schedule(CT_SCHEDULE)
CI_SCHEDULE = build_ci_schedule(CT_SCHEDULE)

# ============================================================================
# 12_TURMAS_DH  (Nivel 4 - inclui bloco de PS 01-30 + DH 01-15)
# ============================================================================
ws = wb.create_sheet("12_TURMAS_DH")
set_col_widths(ws, [3, 12, 14, 14, 16, 12, 12, 22, 16, 30, 3])
style_title(ws, "B2:J2", "VISÃO POR TURMA — PALESTRAS (PS) E DESENVOLVIMENTO HUMANO (DH)")
style_subtitle(ws, "B3:J3", f"Cenário recomendado ({RECOMMENDED}: {REC['salas']} salas, DH {REC['freq']}) — ver abas 06-09 para os outros 3 cenários")

r = 5
ws.merge_cells(f"B{r}:J{r}")
ws.cell(r,2,"PALESTRAS DE SENSIBILIZAÇÃO (PS 01 – PS 30)").font=FONT_H1
ws.cell(r,2).fill = PatternFill("solid", fgColor=STAGE_COLORS["PS"])
ws.row_dimensions[r].height=20
r += 1
header_row(ws, r, 2, ["Turma","Data","Dia da\nsemana","Carga\nhorária","Encontros","Nº aprox.\nparticipantes","Predecessora","Situação da informação"])
r += 1
for p in PS_SCHEDULE:
    body_cell(ws, r, 2, p["turma"], bold=True)
    dc = body_cell(ws, r, 3, p["data"], align="center"); dc.number_format = DATE_FMT
    body_cell(ws, r, 4, p["data"].strftime("%A").replace("Monday","Segunda").replace("Tuesday","Terça").replace("Wednesday","Quarta").replace("Thursday","Quinta").replace("Friday","Sexta"), align="center")
    body_cell(ws, r, 5, "2h", align="center")
    body_cell(ws, r, 6, 1, align="center")
    body_cell(ws, r, 7, "20-40", align="center")
    body_cell(ws, r, 8, "—", align="center")
    classification_cell(ws, r, 9, "Premissa provisória")
    ws.row_dimensions[r].height = 18
    r += 1

r += 1
ws.merge_cells(f"B{r}:J{r}")
ws.cell(r,2,f"DESENVOLVIMENTO HUMANO (DH 01 – DH 15) — Cenário {RECOMMENDED}").font=FONT_H1
ws.cell(r,2).fill = PatternFill("solid", fgColor=STAGE_COLORS["DH"])
ws.row_dimensions[r].height=20
r += 1
header_row(ws, r, 2, ["Turma","Início","Término","Frequência","Carga\nhorária","Encontros","Sala","Nº aprox.\nparticipantes","Predecessora / Situação"])
r += 1
ps_per_dh = 2
for i, t in enumerate(REC["turmas"]):
    body_cell(ws, r, 2, t["turma"], bold=True)
    dic = body_cell(ws, r, 3, t["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 4, t["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 5, t["cadencia"], align="center")
    body_cell(ws, r, 6, "20h (premissa; alternativa 22h)", align="center")
    body_cell(ws, r, 7, 10, align="center")
    body_cell(ws, r, 8, t["sala"], align="center")
    body_cell(ws, r, 9, "20-40", align="center")
    pred = f"PS {2*i+1:02d} e PS {2*i+2:02d} (regra 2 PS = 1 DH — premissa)"
    c = ws.cell(r, 10); c.value = pred; c.font = FONT_NOTE; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BORDER_ALL
    ws.row_dimensions[r].height = 26
    r += 1

freeze(ws, "B7")

# ============================================================================
# 13_OFICINAS_GESTAO
# ============================================================================
ws = wb.create_sheet("13_OFICINAS_GESTAO")
set_col_widths(ws, [3, 12, 14, 14, 12, 12, 12, 12, 34, 3])
style_title(ws, "B2:I2", "OFICINA DE GESTÃO (OG) — DETALHAMENTO DE TURMAS")
style_subtitle(ws, "B3:I3", "Premissa provisória: 1 turma de OG por turma de DH concluída (15 no total). Ver divergência com a regra \"2 DH = 1 OG + 1 CT\" nas abas 01/02/03.")
r = 5
header_row(ws, r, 2, ["Turma","Início","Término","Carga\nhorária","Encontros","Sala","Nº aprox.\nparticipantes","Predecessora","Classificação"])
r += 1
for og in OG_SCHEDULE:
    body_cell(ws, r, 2, og["turma"], bold=True)
    dic = body_cell(ws, r, 3, og["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 4, og["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 5, "4h", align="center")
    body_cell(ws, r, 6, 2, align="center")
    body_cell(ws, r, 7, og["sala"], align="center")
    body_cell(ws, r, 8, "15-18", align="center")
    body_cell(ws, r, 9, og["dh_origem"] + " concluída", align="center")
    classification_cell(ws, r, 10, "Premissa provisória")
    ws.row_dimensions[r].height = 18
    r += 1
freeze(ws, "B6")

# ============================================================================
# 14_CURSOS_TECNICOS
# ============================================================================
ws = wb.create_sheet("14_CURSOS_TECNICOS")
set_col_widths(ws, [3, 10, 30, 14, 14, 12, 12, 10, 10, 26, 34, 3])
style_title(ws, "B2:K2", "CURSOS TÉCNICOS (CT) — DETALHAMENTO DE TURMAS")
style_subtitle(ws, "B3:K3", "10 turmas confirmadas na planilha operacional. Meta contratual do PDF macro é 13 — faltam 3 turmas a definir (ver aba 02).")
r = 5
header_row(ws, r, 2, ["Sigla","Tema","Início","Término","Carga\nhorária","Encontros","Mín.","Máx.","Fornecedor / Local","Predecessora / Situação"])
r += 1
for ct in CT_SCHEDULE:
    body_cell(ws, r, 2, ct["sigla"], bold=True)
    body_cell(ws, r, 3, ct["tema"], wrap=True)
    dic = body_cell(ws, r, 4, ct["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 5, ct["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 6, f"{ct['carga_horaria']}h", align="center")
    body_cell(ws, r, 7, ct["encontros"], align="center")
    body_cell(ws, r, 8, ct["min"], align="center")
    body_cell(ws, r, 9, ct["max"], align="center")
    fill_forn = LRED if "?" in ct["fornecedor"] else None
    body_cell(ws, r, 10, ct["fornecedor"], wrap=True, fill=fill_forn)
    pred = f"4ª turma de DH concluída em {CT_GATE.strftime('%d/%m/%Y')} (premissa: gatilho após 4 DH + 1 semana)"
    c = ws.cell(r, 11); c.value = pred; c.font = FONT_NOTE; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BORDER_ALL
    ws.row_dimensions[r].height = 30
    r += 1

r += 1
ws.merge_cells(f"B{r}:K{r}")
extra_note = ws.cell(r, 2, "PENDENTE DE VALIDAÇÃO: 3 turmas adicionais necessárias para atingir a meta contratual de 13 turmas do PDF macro (Núcleo 2). "
                            "Temas e fornecedores não definidos neste material — recomenda-se priorizar cursos com maior demanda identificada no diagnóstico de participantes.")
extra_note.font = Font(italic=True, bold=True, color=RED)
extra_note.fill = PatternFill("solid", fgColor=LRED)
extra_note.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
ws.row_dimensions[r].height = 34
freeze(ws, "B6")

# ============================================================================
# 15_EAD_CONSULTORIAS
# ============================================================================
ws = wb.create_sheet("15_EAD_CONSULTORIAS")
set_col_widths(ws, [3, 26, 16, 16, 16, 40, 3])
style_title(ws, "B2:F2", "CURSO EAD E CONSULTORIA INDIVIDUAL (CI)")
style_subtitle(ws, "B3:F3", "Etapas finais da trilha — Curso Técnico é pré-requisito para Consultoria Individual (confirmado no PDF macro)")
r = 5
header_row(ws, r, 2, ["Atividade","Início","Término","Carga horária / Meta","Observação"])
r += 1
body_cell(ws, r, 2, "Curso EAD — Planejamento, Marketing e Finanças", bold=True)
dic = body_cell(ws, r, 3, EAD_SCHEDULE["inicio"], align="center"); dic.number_format=DATE_FMT
dfc = body_cell(ws, r, 4, EAD_SCHEDULE["fim"], align="center"); dfc.number_format=DATE_FMT
body_cell(ws, r, 5, f"{EAD_SCHEDULE['carga_horaria']}h / {EAD_SCHEDULE['encontros']} módulos", align="center")
body_cell(ws, r, 6, "Início estimado logo após a conclusão do primeiro Curso Técnico de cada trilha (autoinstrucional)", wrap=True)
ws.row_dimensions[r].height = 30
r += 1
body_cell(ws, r, 2, "Consultoria Individual (CI)", bold=True)
dic = body_cell(ws, r, 3, CI_SCHEDULE["inicio"], align="center"); dic.number_format=DATE_FMT
dfc = body_cell(ws, r, 4, CI_SCHEDULE["fim"], align="center"); dfc.number_format=DATE_FMT
body_cell(ws, r, 5, f"{CI_SCHEDULE['atendimentos_por_pessoa']} atendimentos/pessoa; meta {CI_SCHEDULE['meta_pessoas']} pessoas", align="center")
body_cell(ws, r, 6, "Rolling: inicia conforme cada aluno conclui seu Curso Técnico. Meta cruzada entre nota da planilha operacional (450 atend./3) e PDF macro (13 turmas = 150 pessoas).", wrap=True)
ws.row_dimensions[r].height = 40
freeze(ws, "B6")

# ============================================================================
# Monta lista consolidada de atividades (para 10_CRONOGRAMA_MESTRE e 17_GANTT)
# ============================================================================
def build_master_activities():
    acts = []
    acts.append({"etapa":"PS","nome":"Palestras de Sensibilização (PS 01-30)","inicio":PS_SCHEDULE[0]["data"],
                  "fim":PS_SCHEDULE[-1]["data"],"local":"A definir por comunidade","obs":"3x/semana (Ter/Qua/Qui), 2h cada — Confirmada"})
    for t in REC["turmas"]:
        acts.append({"etapa":"DH","nome":t["turma"],"inicio":t["inicio"],"fim":t["fim"],"local":t["sala"],
                     "obs":f"Cadência {t['cadencia']} — Cenário {RECOMMENDED} (recomendado)"})
    for og in OG_SCHEDULE:
        acts.append({"etapa":"OG","nome":og["turma"],"inicio":og["inicio"],"fim":og["fim"],"local":og["sala"],
                     "obs":f"Predecessora: {og['dh_origem']} — Premissa provisória"})
    for ct in CT_SCHEDULE:
        acts.append({"etapa":"CT","nome":f"{ct['sigla']} — {ct['tema']}","inicio":ct["inicio"],"fim":ct["fim"],
                     "local":ct["fornecedor"],"obs":f"{ct['carga_horaria']}h, {ct['encontros']} encontros, noturno"})
    acts.append({"etapa":"EAD","nome":"Curso EAD — Planejamento, Marketing e Finanças","inicio":EAD_SCHEDULE["inicio"],
                 "fim":EAD_SCHEDULE["fim"],"local":"Plataforma online","obs":"6h / 3 módulos, autoinstrucional"})
    acts.append({"etapa":"CI","nome":"Consultoria Individual (rolling)","inicio":CI_SCHEDULE["inicio"],
                 "fim":CI_SCHEDULE["fim"],"local":"Presencial/Online","obs":f"Meta {CI_SCHEDULE['meta_pessoas']} pessoas, 3 atendimentos/pessoa"})
    acts.append({"etapa":"REC","nome":"Recesso institucional","inicio":RECESSO_START,"fim":RECESSO_END,
                 "local":"—","obs":"Informação pendente de validação"})
    for hd, name in HOLIDAYS.items():
        acts.append({"etapa":"REC","nome":f"Feriado: {name}","inicio":hd,"fim":hd,"local":"—","obs":"Confirmada (calendário municipal)"})
    acts.sort(key=lambda a: a["inicio"])
    return acts

MASTER_ACTIVITIES = build_master_activities()

# ============================================================================
# 10_CRONOGRAMA_MESTRE
# ============================================================================
ws = wb.create_sheet("10_CRONOGRAMA_MESTRE")
set_col_widths(ws, [3, 8, 34, 14, 14, 22, 40, 3])
style_title(ws, "B2:G2", "CRONOGRAMA MESTRE DE PAULÍNIA — VISÃO COMPLETA DA EXECUÇÃO")
style_subtitle(ws, "B3:G3", f"Cenário recomendado ({RECOMMENDED}: {REC['salas']} salas, DH {REC['freq']}) — do início da 1ª Palestra ao fim da Consultoria Individual")
r = 5
header_row(ws, r, 2, ["Etapa","Atividade / Turma","Início","Término","Local / Sala","Observação (premissa/situação)"])
r += 1
for a in MASTER_ACTIVITIES:
    color = STAGE_COLORS.get(a["etapa"], WHITE)
    body_cell(ws, r, 2, a["etapa"], bold=True, fill=color, align="center")
    body_cell(ws, r, 3, a["nome"], wrap=True)
    dic = body_cell(ws, r, 4, a["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 5, a["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 6, a["local"], align="center", wrap=True)
    body_cell(ws, r, 7, a["obs"], wrap=True)
    ws.row_dimensions[r].height = 20
    r += 1

r += 1
ws.cell(r,2,"LEGENDA").font = FONT_H2
r += 1
for name, color in [("PS - Palestra de Sensibilização", STAGE_COLORS["PS"]), ("DH - Desenvolvimento Humano", STAGE_COLORS["DH"]),
                     ("OG - Oficina de Gestão", STAGE_COLORS["OG"]), ("CT - Curso Técnico", STAGE_COLORS["CT"]),
                     ("EAD - Curso Online", STAGE_COLORS["EAD"]), ("CI - Consultoria Individual", STAGE_COLORS["CI"]),
                     ("REC - Recesso / feriado", STAGE_COLORS["REC"])]:
    body_cell(ws, r, 2, "", fill=color); ws.cell(r,2).border=BORDER_ALL
    body_cell(ws, r, 3, name)
    r += 1

freeze(ws, "B6")

# ============================================================================
# 02_DIVERGENCIAS
# ============================================================================
ws = wb.create_sheet("02_DIVERGENCIAS")
set_col_widths(ws, [3, 22, 34, 30, 30, 24, 22, 16, 3])
style_title(ws, "B2:H2", "DIVERGÊNCIAS ENTRE OS 4 DOCUMENTOS — NADA FOI CORRIGIDO SILENCIOSAMENTE")
style_subtitle(ws, "B3:H3", "Toda linha aponta os documentos envolvidos, o impacto no cronograma e uma decisão provisória — a validar com Sebrae/Instituto")

r = 5
header_row(ws, r, 2, ["Tema","Divergência encontrada","Documentos envolvidos","Valores em conflito","Impacto no cronograma","Decisão recomendada (provisória)","Classificação"])
r += 1

DIVERGENCIAS = [
 ("Carga horária do DH", "Planilha operacional e PDF contratual indicam 22h/11 encontros; reunião recente relata redução para 20h pelo Sebrae, sem documento formal.",
  "Proposta (xlsx) + PDF macro  vs.  ata de reunião / Simulação de cenários",
  "22h / 11 encontros  vs.  20h / 10 encontros",
  "Muito alto: cada turma de DH muda ±1 semana de duração; em 15 turmas encadeadas, o efeito acumulado desloca o início dos cursos técnicos e o fim da fase DH em várias semanas.",
  "Usar 20h/10 encontros apenas como premissa de trabalho (é o que consta na simulação mais recente), mas obter confirmação formal por escrito do Sebrae antes de compromissar datas com a Petrobras.",
  "Divergência entre documentos"),
 ("Quantidade de turmas de Curso Técnico", "O documento macro contratual (Anexos do projeto) indica 13 turmas para o Núcleo 2 (Paulínia); a planilha operacional detalhada soma apenas 10 turmas nos 7 temas listados.",
  "PDF \"Apres. Projeto e Anexos 01.2026\"  vs.  Proposta (xlsx)",
  "13 turmas (meta contratual)  vs.  10 turmas (planilha operacional)",
  "Muito alto: se 13 for a meta válida, faltam 3 turmas/fornecedores a contratar, o que desloca o fim da fase de cursos técnicos e a meta de 195 participantes.",
  "Levantar com a coordenação quais 3 turmas adicionais fariam sentido (ex.: 2ª turma de Automação e Jardinagem, ou nova Beleza/Cabelos) e formalizar com os fornecedores antes de assumir 13 como meta.",
  "Divergência entre documentos"),
 ("Regra de formação do DH a partir das Palestras", "A planilha operacional traz duas anotações conflitantes na mesma aba: \"2 PS = 1 DH\" (fecha com 15 turmas) e \"4 OS = 1 DH\" (não fecha; \"OS\" não corresponde a nenhuma sigla da trilha).",
  "Proposta (xlsx), observações A27 e A30",
  "2 PS = 1 DH  vs.  4 OS = 1 DH",
  "Alto: define o ritmo de liberação de novos lotes de DH; usar a regra errada muda a data de abertura de cada turma.",
  "Adotar \"2 PS = 1 DH\" (única leitura que fecha com as 15 turmas confirmadas) e tratar \"4 OS = 1 DH\" como possível erro de digitação a confirmar com quem escreveu a anotação original.",
  "Divergência entre documentos"),
 ("Regra de composição da Oficina de Gestão e Curso Técnico", "A planilha operacional lista 15 turmas de OG (1 para cada DH), mas também traz a anotação \"2 DH = 1 OG + 1 CT\", que implicaria só ~7-8 turmas de OG para 15 DH.",
  "Proposta (xlsx), linha J5 vs. observação A28",
  "15 turmas de OG (1:1 com DH)  vs.  ~7-8 turmas de OG (1 a cada 2 DH)",
  "Muito alto: dobra ou reduz pela metade a necessidade de salas/horários de OG, afetando toda a fase seguinte à DH.",
  "Confirmar com a coordenação se cada turma de DH forma sua própria OG (15 no total) ou se as turmas de DH são pareadas antes da OG (aprox. 7-8 no total).",
  "Divergência entre documentos"),
 ("Quórum mínimo dos cursos técnicos", "A planilha operacional traz mínimos de 14 ou 15 participantes por turma, variando por curso; a reunião menciona 16 como referência geral.",
  "Proposta (xlsx), coluna H  vs.  ata de reunião",
  "14-15 (planilha, por curso)  vs.  16 (reunião, geral)",
  "Alto: se 16 for o piso real exigido pelos fornecedores, pode faltar gente inscrita quando 4 turmas de DH (≈80 pessoas) alimentam 2 turmas técnicas.",
  "Confirmar o mínimo por fornecedor/curso antes de fechar contratos, em vez de aplicar um número único a todos os cursos.",
  "Divergência entre documentos"),
 ("Disponibilidade de salas para o DH", "O cronograma anterior da coordenação foi desenhado com 3 salas simultâneas; a discussão mais recente também cogita apenas 2 salas.",
  "\"Possibilidades de Cronograma\" (título da aba: \"3 Salas Simultâneas\")  vs.  briefing / Simulação de cenários",
  "3 salas  vs.  2 salas",
  "Muito alto: é a variável que mais muda a duração total da fase de DH (73 dias com 3 salas/5x semana vs. 169 dias com 2 salas/2x semana).",
  "Confirmar com antecedência a disponibilidade real e exclusiva de salas à noite antes de comunicar datas ao Sebrae — ver comparativo de cenários na aba 05.",
  "Divergência entre documentos"),
 ("Recesso institucional", "A simulação de cenários bloqueia o período de 21/12/2026 a 08/01/2027 como recesso, mas não há confirmação formal de que o Instituto/Sebrae adotarão esse recesso.",
  "Simulação de cenários (aba \"Premissas e Pendências\")",
  "21/12/2026 a 08/01/2027 (assumido)  vs.  nenhuma confirmação",
  "Médio: cenários que já se estendem para dezembro/janeiro (C e D) ficam ainda mais espremidos ou mais longos conforme o recesso seja confirmado ou não.",
  "Confirmar a duração exata do recesso institucional 2026/2027 junto à coordenação administrativa.",
  "Informação pendente de validação"),
 ("Datas com ano incorreto no cronograma anterior", "A aba \"Paulínia\" do arquivo \"Possibilidades de Cronograma\" tem uma seção rotulada \"Janeiro de 2027\" cujas datas de calendário estão grafadas como 2026 (ex.: célula D152 = 04/01/2026), e ao menos uma linha traz data de término anterior à de início (N13 = 05/12/2026 x O13 = 18/11/2026).",
  "\"Possibilidades de Cronograma- Paulínia.xlsx\", aba Paulínia",
  "Rótulo \"Janeiro de 2027\" com datas grafadas 2026; DHP9 com fim (18/11) anterior ao início (05/12)",
  "Alto: se usado como referência sem correção, o leitor pode interpretar as datas de janeiro como se fossem no ano errado, ou aceitar uma janela de datas invertida.",
  "Não usar essa planilha como fonte de datas finais; tratá-la apenas como registro histórico da tentativa anterior (preservada, sem edição, no arquivo original).",
  "Divergência entre documentos"),
 ("Fornecedor do curso de Jardinagem", "O cronograma anterior da coordenação registra o fornecedor do curso de Jardinagem como \"?\" — não definido.",
  "\"Possibilidades de Cronograma- Paulínia.xlsx\", célula R48",
  "Fornecedor: \"?\"",
  "Alto: sem fornecedor confirmado, a turma de Jardinagem não pode ser efetivamente contratada nem ter data de início garantida.",
  "Priorizar a definição do fornecedor de Jardinagem antes de comunicar a data desta turma aos participantes.",
  "Informação pendente de validação"),
 ("Meta de participantes por turma de DH: valor único ou faixa", "O PDF macro usa \"20 pessoas por turma\" como parâmetro único de cálculo de metas; a planilha operacional define uma faixa de 20 a 40.",
  "PDF \"Apres. Projeto e Anexos\"  vs.  Proposta (xlsx), colunas H4/I4",
  "20 pessoas/turma (fixo, para cálculo de meta)  vs.  20 a 40 (faixa operacional)",
  "Baixo a médio: não é contraditório em si (20 pode ser a média usada para a meta), mas pode gerar expectativa de turmas menores do que a operação permite.",
  "Comunicar a meta com a faixa completa (20-40) para não subdimensionar a expectativa de vagas.",
  "Informação pendente de validação"),
]

for tema, div, docs, valores, impacto, decisao, classe in DIVERGENCIAS:
    body_cell(ws, r, 2, tema, bold=True, wrap=True)
    body_cell(ws, r, 3, div, wrap=True)
    body_cell(ws, r, 4, docs, wrap=True)
    body_cell(ws, r, 5, valores, wrap=True)
    body_cell(ws, r, 6, impacto, wrap=True)
    body_cell(ws, r, 7, decisao, wrap=True)
    classification_cell(ws, r, 8, classe)
    ws.row_dimensions[r].height = 90
    r += 1

freeze(ws, "B6")




# ============================================================================
# 11_AGENDA_POR_SALA
# ============================================================================
ws = wb.create_sheet("11_AGENDA_POR_SALA")
set_col_widths(ws, [3, 12, 12, 14, 14, 14, 12, 30, 3])
style_title(ws, "B2:H2", "AGENDA POR SALA — SEM SOBREPOSIÇÃO CONFIRMADA")
style_subtitle(ws, "B3:H3", f"Cenário {RECOMMENDED} (recomendado): {REC['salas']} salas internas (DH + OG) — cursos técnicos ocorrem em fornecedores externos (ver tabela abaixo)")

r = 5
ws.merge_cells(f"B{r}:H{r}")
ws.cell(r,2,"SALAS INTERNAS — DESENVOLVIMENTO HUMANO E OFICINA DE GESTÃO").font=FONT_H1
ws.cell(r,2).fill = FILL_H1
r += 1
header_row(ws, r, 2, ["Sala","Turma","Tipo","Início","Término","Cadência","Conflito de horário?","Observação"])
r += 1
room_rows = []
for t in REC["turmas"]:
    room_rows.append((t["sala"], t["turma"], "DH", t["inicio"], t["fim"], t["cadencia"]))
for og in OG_SCHEDULE:
    room_rows.append((og["sala"], og["turma"], "OG", og["inicio"], og["fim"], "Seg-Sex (2 encontros)"))
room_rows.sort(key=lambda x: (x[0], x[3]))
for sala, turma, tipo, ini, fim, cad in room_rows:
    body_cell(ws, r, 2, sala, bold=True, align="center")
    body_cell(ws, r, 3, turma, align="center")
    body_cell(ws, r, 4, tipo, align="center", fill=STAGE_COLORS.get(tipo))
    dic = body_cell(ws, r, 5, ini, align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 6, fim, align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 7, cad, align="center")
    body_cell(ws, r, 8, "Não (verificado por dia da semana)", align="center", fill=LGREEN)
    ws.row_dimensions[r].height = 18
    r += 1

r += 1
ws.merge_cells(f"B{r}:H{r}")
ws.cell(r,2,"LOCAIS EXTERNOS — CURSOS TÉCNICOS (fornecedores)").font=FONT_H1
ws.cell(r,2).fill = FILL_H1
r += 1
header_row(ws, r, 2, ["Fornecedor / Local","Turma","Tema","Início","Término","Período livre até próxima turma","Observação",""])
r += 1
by_forn = {}
for ct in CT_SCHEDULE:
    by_forn.setdefault(ct["fornecedor"], []).append(ct)
for forn, cts in by_forn.items():
    cts_sorted = sorted(cts, key=lambda c: c["inicio"])
    for i, ct in enumerate(cts_sorted):
        body_cell(ws, r, 2, forn, bold=True, wrap=True, fill=(LRED if "?" in forn else None))
        body_cell(ws, r, 3, ct["sigla"], align="center")
        body_cell(ws, r, 4, ct["tema"], wrap=True)
        dic = body_cell(ws, r, 5, ct["inicio"], align="center"); dic.number_format=DATE_FMT
        dfc = body_cell(ws, r, 6, ct["fim"], align="center"); dfc.number_format=DATE_FMT
        if i+1 < len(cts_sorted):
            gap = (cts_sorted[i+1]["inicio"] - ct["fim"]).days
            livre = f"{gap} dias até {cts_sorted[i+1]['sigla']}"
        else:
            livre = "Última turma desta trilha"
        body_cell(ws, r, 7, livre, align="center")
        body_cell(ws, r, 8, "Fornecedor não confirmado" if "?" in forn else "—", align="center")
        ws.row_dimensions[r].height = 26
        r += 1

freeze(ws, "B6")

# ============================================================================
# 16_RISCOS_E_DECISOES
# ============================================================================
ws = wb.create_sheet("16_RISCOS_E_DECISOES")
set_col_widths(ws, [3, 30, 26, 22, 26, 16, 14, 26, 3])
style_title(ws, "B2:H2", "RISCOS, DIVERGÊNCIAS E DECISÕES PENDENTES")
style_subtitle(ws, "B3:H3", "Quadro consolidado para acompanhamento com Instituto da Criança, Sebrae e Petrobras")
r = 5
header_row(ws, r, 2, ["Risco / Divergência","Documentos envolvidos","Impacto no cronograma","Decisão necessária","Responsável pela validação","Status","Recomendação provisória"])
r += 1

RISCOS = [
 ("Carga horária do DH indefinida (20h x 22h)", "Proposta (xlsx) + PDF macro vs. ata de reunião",
  "Muda a duração de cada uma das 15 turmas em até 1 semana, deslocando toda a cadeia OG→CT→EAD→CI.",
  "Confirmar por escrito a carga horária oficial do DH junto ao Sebrae.", "Coordenação / Sebrae", "Aberto",
  "Formalizar 20h por e-mail/ata assinada antes de comunicar datas à Petrobras."),
 ("Quantidade de turmas de Curso Técnico indefinida (10 x 13)", "Proposta (xlsx) vs. PDF macro",
  "Se 13 for a meta válida, é preciso contratar 3 novos fornecedores/turmas, o que pode adicionar 4-8 semanas ao cronograma de CT.",
  "Definir se a meta de 13 será cumprida em Paulínia e, em caso positivo, quais temas/fornecedores fecham a diferença.",
  "Coordenação / Instituto da Criança", "Aberto", "Levar a divergência à reunião de alinhamento com o Sebrae antes de fechar contratos."),
 ("Disponibilidade real de salas (2 x 3) não confirmada", "\"Possibilidades de Cronograma\" (3 salas) vs. briefing (2 ou 3)",
  "É a variável que mais impacta a duração da fase DH: 73 dias (cenário recomendado, 3 salas) até 169 dias (cenário C, 2 salas).",
  "Confirmar contratualmente a disponibilidade exclusiva de 3 salas à noite, de segunda a sexta.",
  "Coordenação local / Espaço físico", "Aberto", "Adotar o Cenário B como referência de negociação, mas manter D como contingência caso só 2 salas sejam viabilizadas."),
 ("Regra de composição da Oficina de Gestão (15 x ~7-8 turmas)", "Proposta (xlsx), linha J5 vs. anotação A28",
  "Dobra ou reduz pela metade a necessidade de salas/horários dedicados à OG.",
  "Confirmar se cada DH gera sua própria OG (1:1) ou se as OG são formadas a partir de pares de turmas de DH.",
  "Coordenação pedagógica", "Aberto", "Assumir 1:1 (15 turmas) como premissa de trabalho até confirmação, por ser a leitura mais conservadora em capacidade."),
 ("Fornecedor de Jardinagem não definido", "\"Possibilidades de Cronograma\", célula R48 = \"?\"",
  "Turma de Jardinagem não pode ser comunicada aos participantes sem fornecedor e local confirmados.",
  "Selecionar e contratar fornecedor para o curso de Jardinagem.",
  "Coordenação de parcerias", "Aberto", "Priorizar a definição nas próximas 4 semanas, antes da 4ª turma de DH concluir (gatilho do 1º Curso Técnico)."),
 ("Recesso institucional não confirmado formalmente", "Simulação de cenários (premissa observada)",
  "Cenários que se estendem para dezembro/janeiro (C e D) podem precisar de ajuste fino de datas conforme o recesso real.",
  "Confirmar datas oficiais de recesso institucional 2026/2027.",
  "Administrativo / RH", "Aberto", "Manter 21/12/2026-08/01/2027 como premissa de planejamento; ajustar se o recesso oficial for diferente."),
 ("Prazo apertado até o início das Palestras", "Contexto do projeto — data atual 21/07/2026, início das PS em 11/08/2026",
  "Restam menos de 3 semanas até a primeira Palestra de Sensibilização; qualquer atraso na confirmação de salas/premissas comprime a mobilização.",
  "Validar premissas críticas (salas, carga horária do DH) com urgência.",
  "Coordenação geral do projeto", "Aberto — urgente", "Tratar as pendências desta aba como prioridade das próximas 2 semanas."),
 ("Quórum mínimo dos cursos técnicos variável (14-18) vs. referência única de 16 citada em reunião", "Proposta (xlsx) vs. ata de reunião",
  "Se 16 for aplicado a todos os cursos, turmas com quórum planejado de 14-15 podem não abrir.",
  "Confirmar quórum mínimo por fornecedor/curso, não um valor único.",
  "Coordenação de parcerias", "Aberto", "Negociar quórum por curso com cada fornecedor antes de divulgar datas."),
 ("ACHADO DESTA ANÁLISE: as 3 salas de DH ficam saturadas o tempo todo no cenário recomendado (B)", "Modelagem desta análise (nenhum dos 4 documentos originais havia cruzado a agenda de salas do DH com a da OG)",
  "No cenário B, cada uma das 3 salas de DH recebe 5 lotes consecutivos sem dia livre durante os 73 dias da fase — não há horário disponível na mesma sala para a Oficina de Gestão.",
  "Confirmar a existência de pelo menos 3 espaços adicionais (fora das salas de DH) para rodar a OG em paralelo, ou aceitar que a OG só comece após o fim de toda a fase DH (04/11).",
  "Coordenação local / Espaço físico", "Aberto — novo", "Reservar 3 espaços dedicados à OG (ex.: salas de reunião, auditório, espaço comunitário) distintos das 3 salas de DH, seguindo a mesma cadência de liberação dos lotes."),
]
for risco, docs, impacto, decisao, resp, status, rec in RISCOS:
    body_cell(ws, r, 2, risco, bold=True, wrap=True)
    body_cell(ws, r, 3, docs, wrap=True)
    body_cell(ws, r, 4, impacto, wrap=True)
    body_cell(ws, r, 5, decisao, wrap=True)
    body_cell(ws, r, 6, resp, wrap=True, align="center")
    fill_status = LRED if "urgente" in status else LGOLD
    body_cell(ws, r, 7, status, align="center", fill=fill_status, bold=True)
    body_cell(ws, r, 8, rec, wrap=True)
    ws.row_dimensions[r].height = 75
    r += 1

freeze(ws, "B6")

# ============================================================================
# 04_RESUMO_EXECUTIVO  (NIVEL 1 - visao de uma pagina)
# ============================================================================
ws = wb.create_sheet("04_RESUMO_EXECUTIVO")
ws.sheet_view.showGridLines = False
set_col_widths(ws, [2, 16, 16, 16, 16, 16, 16, 16, 16, 16, 2])
style_title(ws, "B2:J2", "PARTIU! APRENDER E EMPREENDER — PAULÍNIA", height=34)
style_subtitle(ws, "B3:J3", "Trilha completa: Palestra de Sensibilização → Desenvolvimento Humano → Oficina de Gestão → Curso Técnico → EAD → Consultoria Individual")

r = 5
period_txt = f"PERÍODO TOTAL: {PS_SCHEDULE[0]['data'].strftime('%d/%m/%Y')} a {CI_SCHEDULE['fim'].strftime('%d/%m/%Y')}  (cenário recomendado)"
ws.merge_cells(f"B{r}:J{r}")
pc = ws.cell(r,2,period_txt); pc.font=Font(size=13,bold=True,color=WHITE); pc.fill=PatternFill("solid",fgColor=TEAL)
pc.alignment=Alignment(horizontal="center", vertical="center"); ws.row_dimensions[r].height=26
r += 2

# Marcos principais - cartoes
marcos = [
 ("Início das Palestras", PS_SCHEDULE[0]["data"], STAGE_COLORS["PS"]),
 ("Início do DH", REC["inicio_fase"], STAGE_COLORS["DH"]),
 ("4ª turma de DH concluída", REC["quarta_concl"], STAGE_COLORS["DH"]),
 ("1º Curso Técnico possível", CT_GATE, STAGE_COLORS["CT"]),
 ("Fim do DH (15 turmas)", REC["fim_fase"], STAGE_COLORS["DH"]),
 ("Fim dos Cursos Técnicos", max(c["fim"] for c in CT_SCHEDULE), STAGE_COLORS["CT"]),
]
col = 2
for label, d, color in marcos:
    c1 = ws.cell(r, col, label)
    c1.font = Font(size=9, bold=True); c1.fill = PatternFill("solid", fgColor=color)
    c1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c1.border = BORDER_ALL
    ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col)
    c2 = ws.cell(r+1, col, d); c2.number_format = DATE_FMT
    c2.font = Font(size=11, bold=True); c2.alignment = Alignment(horizontal="center")
    c2.fill = PatternFill("solid", fgColor="FFFFFF"); c2.border = BORDER_ALL
    ws.row_dimensions[r].height = 34
    ws.row_dimensions[r+1].height = 20
    col += 1
r += 3

# Linha do tempo mensal (nivel executivo, sem excesso de detalhe)
ws.merge_cells(f"B{r}:J{r}")
ws.cell(r,2,"LINHA DO TEMPO — CENÁRIO RECOMENDADO (B: 3 SALAS, DH 5X/SEMANA)").font=FONT_H1
ws.cell(r,2).fill = FILL_H1
r += 1

months = []
m = date(2026,8,1)
end_month = date(2027,3,1)
while m < end_month:
    months.append(m)
    if m.month == 12:
        m = date(m.year+1,1,1)
    else:
        m = date(m.year, m.month+1, 1)

month_row = r
header_row(ws, r, 2, ["Etapa"] + [mm.strftime("%b/%y").upper() for mm in months], fill=FILL_H2, font=FONT_H2)
r += 1

def month_overlap(d1, d2, mstart):
    mend = (date(mstart.year, mstart.month+1,1) - timedelta(days=1)) if mstart.month < 12 else date(mstart.year,12,31)
    return d1 <= mend and d2 >= mstart

stage_ranges = {
 "PS": (PS_SCHEDULE[0]["data"], PS_SCHEDULE[-1]["data"]),
 "DH": (REC["inicio_fase"], REC["fim_fase"]),
 "OG": (min(o["inicio"] for o in OG_SCHEDULE), max(o["fim"] for o in OG_SCHEDULE)),
 "CT": (min(c["inicio"] for c in CT_SCHEDULE), max(c["fim"] for c in CT_SCHEDULE)),
 "EAD": (EAD_SCHEDULE["inicio"], EAD_SCHEDULE["fim"]),
 "CI": (CI_SCHEDULE["inicio"], CI_SCHEDULE["fim"]),
}
stage_labels = {"PS":"Palestras de Sensibilização","DH":"Desenvolvimento Humano","OG":"Oficina de Gestão",
                "CT":"Cursos Técnicos","EAD":"Curso EAD","CI":"Consultoria Individual"}
for key, (d1,d2) in stage_ranges.items():
    body_cell(ws, r, 2, stage_labels[key], bold=True)
    for ci, mm in enumerate(months):
        cell = ws.cell(r, 3+ci)
        cell.border = BORDER_ALL
        if month_overlap(d1,d2,mm):
            cell.fill = PatternFill("solid", fgColor=STAGE_COLORS[key])
    ws.row_dimensions[r].height = 20
    r += 1
# recesso row
body_cell(ws, r, 2, "Recesso / feriados", bold=True)
for ci, mm in enumerate(months):
    cell = ws.cell(r, 3+ci)
    cell.border = BORDER_ALL
    if month_overlap(RECESSO_START, RECESSO_END, mm) or any(month_overlap(h,h,mm) for h in HOLIDAYS):
        cell.fill = PatternFill("solid", fgColor=STAGE_COLORS["REC"])
ws.row_dimensions[r].height = 20
r += 2

# Cenario recomendado - callout
ws.merge_cells(f"B{r}:J{r+2}")
rc = ws.cell(r, 2, "CENÁRIO RECOMENDADO PARA NEGOCIAÇÃO: 3 SALAS, DH 5X/SEMANA\n"
                     "Conclui o Desenvolvimento Humano em 73 dias (24/08 a 04/11/2026) — o mais rápido entre os 4 cenários avaliados — "
                     "liberando o maior tempo possível para Oficina de Gestão, Cursos Técnicos, EAD e Consultoria Individual ainda dentro do ano, "
                     "sem cruzar o recesso institucional. Alternativa de contingência: Cenário D (2 salas, DH 5x/semana).")
rc.font = Font(size=12, bold=True, color=WHITE)
rc.fill = PatternFill("solid", fgColor=NAVY)
rc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[r].height = 26; ws.row_dimensions[r+1].height=26; ws.row_dimensions[r+2].height=26
r += 4

# Legenda simples
ws.merge_cells(f"B{r}:J{r}")
ws.cell(r,2,"LEGENDA").font = FONT_H2
r += 1
lc = 2
for key in ["PS","DH","OG","CT","EAD","CI","REC"]:
    cell = ws.cell(r, lc); cell.fill = PatternFill("solid", fgColor=STAGE_COLORS[key]); cell.border = BORDER_ALL
    ws.cell(r, lc+1, stage_labels.get(key, "Recesso/feriado")).font = Font(size=9)
    lc += 2
r += 2
ws.merge_cells(f"B{r}:J{r}")
note = ws.cell(r,2,"Datas construídas a partir de premissas explicitadas nas abas 01, 02 e 03 deste arquivo — sujeitas a confirmação formal do Sebrae/Instituto.")
note.font = FONT_NOTE

ws.sheet_view.zoomScale = 90
freeze(ws, "B6")

# ============================================================================
# 17_GANTT_EXECUTIVO
# ============================================================================
ws = wb.create_sheet("17_GANTT_EXECUTIVO")
ws.sheet_view.showGridLines = False
style_title(ws, "B2:D2", "GANTT EXECUTIVO — TRILHA COMPLETA (CENÁRIO RECOMENDADO)", height=26)
style_subtitle(ws, "B3:D3", "Grade semanal — cada coluna representa a segunda-feira da semana correspondente")

weeks = []
wk = PS_SCHEDULE[0]["data"]
wk = wk - timedelta(days=wk.weekday())  # segunda-feira da semana
last_day = CI_SCHEDULE["fim"]
while wk <= last_day:
    weeks.append(wk)
    wk += timedelta(days=7)

r = 5
set_col_widths(ws, [3, 30] + [4.2]*len(weeks) + [3])
header_row(ws, r, 2, ["Linha"], fill=FILL_H2, font=FONT_H2)
for i, wkd in enumerate(weeks):
    c = ws.cell(r, 3+i, wkd)
    c.number_format = "DD/MM"
    c.font = Font(size=7, bold=True, color=NAVY)
    c.fill = FILL_H2
    c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)
    c.border = BORDER_ALL
ws.row_dimensions[r].height = 46
r += 1

gantt_rows = [
    ("Palestras de Sensibilização","PS", PS_SCHEDULE[0]["data"], PS_SCHEDULE[-1]["data"]),
    ("Desenvolvimento Humano (agregado)","DH", REC["inicio_fase"], REC["fim_fase"]),
]
for t in REC["turmas"]:
    gantt_rows.append((f"  {t['turma']} ({t['sala']})","DH", t["inicio"], t["fim"]))
gantt_rows.append(("Oficina de Gestão (agregado)","OG", min(o["inicio"] for o in OG_SCHEDULE), max(o["fim"] for o in OG_SCHEDULE)))
for ct in CT_SCHEDULE:
    gantt_rows.append((f"  {ct['sigla']} — {ct['tema'][:22]}","CT", ct["inicio"], ct["fim"]))
gantt_rows.append(("Curso EAD","EAD", EAD_SCHEDULE["inicio"], EAD_SCHEDULE["fim"]))
gantt_rows.append(("Consultoria Individual","CI", CI_SCHEDULE["inicio"], CI_SCHEDULE["fim"]))
gantt_rows.append(("Recesso institucional","REC", RECESSO_START, RECESSO_END))
for hd, name in sorted(HOLIDAYS.items()):
    gantt_rows.append((f"Feriado: {name}","REC", hd, hd))

for label, stage, d1, d2 in gantt_rows:
    lc = ws.cell(r, 2, label)
    lc.font = Font(size=8, bold=(not label.startswith(" ")))
    lc.border = BORDER_ALL
    lc.alignment = Alignment(vertical="center")
    for i, wkd in enumerate(weeks):
        cell = ws.cell(r, 3+i)
        cell.border = BORDER_ALL
        if d1 <= wkd + timedelta(days=6) and d2 >= wkd:
            cell.fill = PatternFill("solid", fgColor=STAGE_COLORS[stage])
    ws.row_dimensions[r].height = 12
    r += 1

freeze(ws, "C6")

# ============================================================================
# REORDENAR ABAS NA ORDEM LOGICA FINAL
# ============================================================================
FINAL_ORDER = [
    "00_LEIA-ME", "01_PREMISSAS", "02_DIVERGENCIAS", "03_FLUXO_DA_TRILHA",
    "04_RESUMO_EXECUTIVO", "05_COMPARATIVO_CENARIOS",
    "06_CENARIO_A", "07_CENARIO_B", "08_CENARIO_C", "09_CENARIO_D",
    "10_CRONOGRAMA_MESTRE", "11_AGENDA_POR_SALA", "12_TURMAS_DH",
    "13_OFICINAS_GESTAO", "14_CURSOS_TECNICOS", "15_EAD_CONSULTORIAS",
    "16_RISCOS_E_DECISOES", "17_GANTT_EXECUTIVO",
]
assert set(FINAL_ORDER) == set(wb.sheetnames), f"Divergencia: {set(FINAL_ORDER) ^ set(wb.sheetnames)}"
wb._sheets = [wb[name] for name in FINAL_ORDER]

# tab colors por macro-tema
TAB_COLORS = {
    "00_LEIA-ME":"808080","01_PREMISSAS":"BF8F00","02_DIVERGENCIAS":"C0392B","03_FLUXO_DA_TRILHA":"1F3864",
    "04_RESUMO_EXECUTIVO":"0F6B5C","05_COMPARATIVO_CENARIOS":"2E5395",
    "06_CENARIO_A":"9DC3E6","07_CENARIO_B":"2E7D32","08_CENARIO_C":"9DC3E6","09_CENARIO_D":"9DC3E6",
    "10_CRONOGRAMA_MESTRE":"1F3864","11_AGENDA_POR_SALA":"2E5395","12_TURMAS_DH":"2E5395",
    "13_OFICINAS_GESTAO":"2E5395","14_CURSOS_TECNICOS":"2E5395","15_EAD_CONSULTORIAS":"2E5395",
    "16_RISCOS_E_DECISOES":"C0392B","17_GANTT_EXECUTIVO":"1F3864",
}
for name, color in TAB_COLORS.items():
    wb[name].sheet_properties.tabColor = color

wb.active = wb.sheetnames.index("04_RESUMO_EXECUTIVO")

OUT_DIR = "/home/user/CRONOGRAMA/cronograma"
os.makedirs(OUT_DIR, exist_ok=True)
OUT_PATH = os.path.join(OUT_DIR, "Cronograma_Paulinia_Partiu_2026.xlsx")
wb.save(OUT_PATH)
print("Arquivo salvo em:", OUT_PATH)
