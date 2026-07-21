# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from model import (PS_SCHEDULE, SCENARIOS, RECOMMENDED, HOLIDAYS, RECESSO_START, RECESSO_END,
                    build_og_schedule_b1_dedicado, build_og_schedule_b2_compartilhado,
                    build_ct_schedule, build_ead_schedule, build_ci_schedule, compute_ci_capacity,
                    CT_CATALOG, parse_iso, cadence_weekdays, holiday_hits_real, simulate_lot_delay,
                    MEDIA_PARTICIPANTES_DH)
from datetime import date, timedelta
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

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
    "PS":  "F4B183", "DH":  "9DC3E6", "OG":  "A9D18E", "CT":  "FFD966",
    "EAD": "B4A7D6", "CI":  "F1948A", "REC": "D9D9D9",
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
    "Condicionado à validação": LGOLD,
}

DATE_FMT = "DD/MM/YYYY"

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
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)

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

WD_PT = {0:"Segunda",1:"Terça",2:"Quarta",3:"Quinta",4:"Sexta",5:"Sábado",6:"Domingo"}

wb = Workbook()
wb.remove(wb.active)

# ============================================================================
# Dados derivados (motor de datas) -- usados em varias abas
# ============================================================================
REC = SCENARIOS[RECOMMENDED]
OG_B1 = build_og_schedule_b1_dedicado(REC)           # 3 salas de DH + 3 espacos adicionais (NAO confirmado)
OG_B2 = build_og_schedule_b2_compartilhado(REC)       # 3 salas compartilhadas entre DH e OG (sem espaco extra)
OG_SCHEDULE = OG_B1   # variante de referencia usada nas abas de cronograma mestre / agenda / turmas (ver 07)
CT_SCHEDULE, CT_GATE = build_ct_schedule(REC)
EAD_SCHEDULE = build_ead_schedule(CT_SCHEDULE)
CI_SCHEDULE = build_ci_schedule(CT_SCHEDULE)
CI_CAPACITY_DEFAULT = compute_ci_capacity(CI_SCHEDULE["inicio"])
DELAY_SIM_7D = simulate_lot_delay(REC, lot_to_delay=1, delay_days=7, absorb_gaps=True)
DELAY_SIM_7D_WORST = simulate_lot_delay(REC, lot_to_delay=1, delay_days=7, absorb_gaps=False)

print("Motor de datas carregado. Construindo abas da V2...")

# ============================================================================
# 00_LEIA-ME
# ============================================================================
ws = wb.create_sheet("00_LEIA-ME")
set_col_widths(ws, [3, 26, 88, 3])
style_title(ws, "B2:C2", "PARTIU! APRENDER E EMPREENDER — CRONOGRAMA PAULÍNIA (NÚCLEO 2) — V2")
style_subtitle(ws, "B3:C3", "Instituto da Criança · Sebrae · Petrobras — V2: revisão crítica da V1, com correções de automação, capacidade e condicionamento de recomendação")

r = 5
ws.cell(r,2,"O QUE MUDOU NESTA VERSÃO (V2)").font = FONT_H2
r += 1
txt = ("Esta é uma revisão crítica do arquivo V1, feita a pedido da coordenação. As principais correções: (1) esclarecimento "
"honesto sobre o que é automatizado por fórmula e o que exige regenerar o arquivo; (2) separação entre \"turmas ativas na "
"semana\" e \"capacidade física por noite\"; (3) a Oficina de Gestão passa a ter duas alternativas explícitas (B1 e B2), "
"nenhuma delas confirmada; (4) o cenário recomendado deixou de ser tratado como aprovado; (5) marcos de validação de quórum "
"e simulação de atraso; (6) Cursos Técnicos reclassificados como datas de simulação, não compromissos; (7) cálculo de "
"capacidade da Consultoria Individual. Ver a aba 19_LOG_DE_ALTERACOES para a lista completa, item a item.")
ws.merge_cells(f"B{r}:C{r+3}")
c = ws.cell(r,2,txt); c.font=FONT_BODY; c.alignment=Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[r].height = 75
r += 5

ws.cell(r,2,"MATRIZ DE AUTOMAÇÃO — O QUE MUDA SOZINHO E O QUE EXIGE REGENERAR O ARQUIVO").font = FONT_H2
r += 1
ws.merge_cells(f"B{r}:C{r}")
note = ws.cell(r,2,"Regra geral: este arquivo NÃO tem macro/script embutido. Fórmulas do Excel recalculam sozinhas; "
                    "qualquer coisa que dependa do algoritmo de alocação de salas/datas (Python) só muda depois de rodar "
                    "scripts/build_workbook.py e gerar um novo arquivo.")
note.font = FONT_NOTE; note.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[r].height = 30
r += 1

header_row(ws, r, 2, ["Premissa (aba 01_PREMISSAS)"], fill=FILL_H2, font=FONT_H2)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2)
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=3)
ws.cell(r,3,"Abas atualizadas automaticamente  /  Abas que exigem regenerar (script)")
ws.cell(r,3).font=FONT_H2; ws.cell(r,3).fill=FILL_H2
ws.cell(r,3).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
r += 1

MATRIZ_AUTOMACAO = [
 ("Cenário recomendado (dropdown A/B/C/D)",
  "AUTOMÁTICO: 04_RESUMO_EXECUTIVO e 05_COMPARATIVO_CENARIOS recalculam o destaque e os marcos por fórmula (INDEX/MATCH). "
  "EXIGE REGERAR: 10, 11, 12-15, 17 (cronograma mestre, agenda e turma a turma continuam no cenário B até serem regeneradas)."),
 ("Datas de início/fim das Palestras (PS)",
  "EXIGE REGERAR script/build_workbook.py: todas as datas de PS, e por consequência as datas de DH/OG/CT/EAD/CI, são recalculadas pelo motor de datas em Python (não há fórmula de calendário de turmas no Excel)."),
 ("Carga horária do DH (20h x 22h)",
  "EXIGE REGERAR: muda a duração de cada turma de DH em model.py; nenhuma fórmula recalcula isso no Excel."),
 ("Nº de salas / frequência do DH (cenários A-D)",
  "EXIGE REGERAR: o algoritmo de alocação de sala por lote está em Python (model.py), não em fórmula de planilha."),
 ("Quantidade de turmas de OG / CT / variante B1 x B2",
  "EXIGE REGERAR: schedules gerados em Python (build_og_schedule_b1_dedicado / _b2_compartilhado / build_ct_schedule)."),
 ("Parâmetros de capacidade da Consultoria Individual (nº consultores, horas/dia, dias/semana, atendimentos simultâneos)",
  "AUTOMÁTICO: células editáveis na aba 15_EAD_CONSULTORIAS calculam por FÓRMULA (sem precisar regenerar) as semanas necessárias e a data de conclusão estimada."),
 ("Recesso institucional (datas)",
  "PARCIAL: o texto e os indicadores citam a data por referência de célula (fórmula) nas abas 00/01/04/17; mas o bloqueio desse período no cálculo de datas de turma (para não agendar encontro dentro do recesso) exige regerar o script."),
]
for premissa, detalhe in MATRIZ_AUTOMACAO:
    body_cell(ws, r, 2, premissa, bold=True, wrap=True)
    body_cell(ws, r, 3, detalhe, wrap=True)
    ws.row_dimensions[r].height = 58
    r += 1

r += 1
ws.merge_cells(f"B{r}:C{r}")
script_note = ws.cell(r, 2, "SCRIPT A EXECUTAR PARA REGERAR: python3 scripts/build_workbook.py  (edite scripts/model.py "
                              "para mudar uma premissa estrutural antes de rodar). Gera um novo .xlsx — não sobrescreve versões anteriores automaticamente.")
script_note.font = Font(bold=True, italic=True, color=NAVY)
script_note.fill = PatternFill("solid", fgColor=LBLUE)
script_note.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
ws.row_dimensions[r].height = 32
r += 2

ws.cell(r,2,"COMO CLASSIFICAMOS CADA INFORMAÇÃO").font = FONT_H2
r += 1
classes = [
    ("Confirmada", "Valor igual em pelo menos 2 fontes independentes, sem contradição."),
    ("Premissa provisória", "Valor usado para construir o cronograma na ausência de confirmação formal; pode mudar."),
    ("Divergência entre documentos", "Os documentos-fonte trazem valores diferentes para o mesmo item; nenhum foi descartado."),
    ("Informação pendente de validação", "Dado ausente, incompleto ou com placeholder (ex.: fornecedor \"?\")."),
    ("Decisão recomendada", "Recomendação da análise, com justificativa — não é um fato, é uma sugestão de encaminhamento."),
    ("Condicionado à validação", "Datas/indicadores calculados a partir de premissas ainda não confirmadas (salas, quórum, fornecedor) — NÃO são compromisso oficial."),
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
    ("01_PREMISSAS", "Painel de premissas editáveis, agora com coluna indicando se a mudança é automática ou exige regeneração."),
    ("02_DIVERGENCIAS", "Tabela completa de conflitos entre os 4 documentos-fonte."),
    ("03_FLUXO_DA_TRILHA", "Mapa da trilha PS → DH → OG → CT → EAD → CI, com as regras quantitativas encontradas."),
    ("04_RESUMO_EXECUTIVO", "Visão de 1 página — recomendação condicionada, não uma aprovação."),
    ("05_COMPARATIVO_CENARIOS", "Painel comparativo dos 4 cenários, com os dois indicadores de capacidade separados."),
    ("06 a 09_CENARIO A-D", "Turma a turma de cada cenário + marcos de validação de quórum + simulação de atraso de 7 dias."),
    ("07_CENARIO_B", "Inclui a comparação das duas variantes de espaço para a OG (B1 x B2)."),
    ("10_CRONOGRAMA_MESTRE", "Linha do tempo completa (cenário B, variante B1 de OG) — texto corrigido sobre o recesso."),
    ("11_AGENDA_POR_SALA", "Ocupação de salas — inclui as duas variantes de OG (B1/B2)."),
    ("12_TURMAS_DH", "PS 01-30 e DH 01-15 detalhados."),
    ("13_OFICINAS_GESTAO", "Detalhamento das duas variantes de OG (B1/B2), nenhuma confirmada."),
    ("14_CURSOS_TECNICOS", "Todas as datas classificadas como \"proposta para simulação\", com quórum e responsável pela validação."),
    ("15_EAD_CONSULTORIAS", "Calculadora de capacidade da Consultoria Individual com parâmetros editáveis."),
    ("16_RISCOS_E_DECISOES", "Quadro de riscos, agora incluindo quórum, atraso em cascata e ambiguidade de espaços de OG."),
    ("17_GANTT_EXECUTIVO", "Gantt visual semanal da trilha completa."),
    ("18_VISAO_PARA_PARCEIROS", "Página única para compartilhar com Sebrae/Petrobras — sem detalhe técnico operacional."),
    ("19_LOG_DE_ALTERACOES", "Lista completa das correções feitas nesta V2, item a item, em resposta à revisão crítica."),
]
for nome, desc in nav:
    body_cell(ws, r, 2, nome, bold=True, fill=LBLUE)
    body_cell(ws, r, 3, desc, wrap=True)
    ws.row_dimensions[r].height = 26
    r += 1

freeze(ws, "A5")
print("00_LEIA-ME ok")

# ============================================================================
# 01_PREMISSAS
# ============================================================================
ws = wb.create_sheet("01_PREMISSAS")
set_col_widths(ws, [3, 30, 16, 38, 20, 16, 3])
style_title(ws, "B2:F2", "PREMISSAS DO CRONOGRAMA — V2")
style_subtitle(ws, "B3:F3", "Células amarelas são editáveis. A coluna \"Automação\" diz, para CADA premissa, se a mudança recalcula sozinha ou exige rodar scripts/build_workbook.py. Ver matriz completa na aba 00_LEIA-ME.")

r = 5
ws.cell(r,2,"CENÁRIO RECOMENDADO PARA AS ABAS 04 E 05 (ligado por fórmula)").font=FONT_H1
ws.cell(r,2).fill=FILL_H1
ws.merge_cells(f"B{r}:F{r}")
ws.row_dimensions[r].height=20
r += 1
body_cell(ws, r, 2, "Cenário recomendado (dropdown)", bold=True)
rec_cell = ws.cell(r, 3, RECOMMENDED)
rec_cell.fill = FILL_EDIT; rec_cell.font=Font(bold=True, size=12); rec_cell.alignment=Alignment(horizontal="center")
rec_cell.border = BORDER_ALL
dv = DataValidation(type="list", formula1='"A,B,C,D"', allow_blank=False)
ws.add_data_validation(dv)
dv.add(rec_cell)
CENARIO_CELL_ADDR = f"'01_PREMISSAS'!C{r}"
body_cell(ws, r, 4, "Alimenta por fórmula os destaques e marcos das abas 04_RESUMO_EXECUTIVO e 05_COMPARATIVO_CENARIOS.", wrap=True)
classification_cell(ws, r, 5, "Confirmada")
body_cell(ws, r, 6, "Automático (fórmula)", align="center", fill=LGREEN)
ws.row_dimensions[r].height = 26
r += 2

def premise_block(ws, r, title, rows):
    ws.merge_cells(f"B{r}:F{r}")
    c = ws.cell(r,2,title); c.font=FONT_H1; c.fill=FILL_H1
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 22
    r += 1
    header_row(ws, r, 2, ["Item","Valor","Observação / Classificação","Status","Automação"], fill=FILL_H2, font=FONT_H2)
    r += 1
    for item, val, obs, status, editable, automacao in rows:
        body_cell(ws, r, 2, item, bold=True)
        vc = body_cell(ws, r, 3, val, align="center")
        if editable:
            vc.fill = FILL_EDIT
        body_cell(ws, r, 4, obs, wrap=True)
        classification_cell(ws, r, 5, status)
        auto_fill = LGREEN if automacao.startswith("Automático") else LRED
        body_cell(ws, r, 6, automacao, align="center", wrap=True, fill=auto_fill)
        ws.row_dimensions[r].height = 30
        r += 1
    return r + 1

r = premise_block(ws, r, "PALESTRAS DE SENSIBILIZAÇÃO (PS)", [
    ("Data de início", date(2026,8,11), "Confirmado em 3 fontes", "Confirmada", True, "Exige regeneração (script)"),
    ("Data-alvo de término", date(2026,10,15), "Confirmado em 3 fontes", "Confirmada", True, "Exige regeneração (script)"),
    ("Total de turmas (Paulínia)", 30, "Base operacional J3=30; PDF macro Núcleo 2 = 30 turmas", "Confirmada", True, "Exige regeneração (script)"),
    ("Frequência", "3x/semana (Ter/Qua/Qui)", "Cadência dedutível da planilha \"Possibilidades\"", "Premissa provisória", True, "Exige regeneração (script)"),
])

r = premise_block(ws, r, "DESENVOLVIMENTO HUMANO (DH)", [
    ("Total de turmas (Paulínia)", 15, "Confirmado em 3 fontes", "Confirmada", True, "Exige regeneração (script)"),
    ("Carga horária total — OPÇÃO USADA NO MODELO", "20h / 10 encontros", "Usada pela simulação de cenários e citada em reunião", "Premissa provisória", True, "Exige regeneração (script)"),
    ("Carga horária total — OPÇÃO ALTERNATIVA (documental)", "22h / 11 encontros", "Planilha operacional e PDF macro", "Divergência entre documentos", False, "Exige regeneração (script)"),
    ("Mínimo/Máximo de participantes por turma", "20 / 40", "Planilha operacional H4/I4", "Confirmada", True, "Exige regeneração (script)"),
    ("Regra de formação — OPÇÃO A", "2 PS = 1 DH", "Fecha exatamente com 30 PS → 15 DH", "Premissa provisória", False, "Exige regeneração (script)"),
    ("Regra de formação — OPÇÃO B (anotação divergente)", "4 OS = 1 DH", "\"OS\" não existe na trilha; provável erro de digitação. Não fecha com 15 turmas mesmo lendo como \"PS\"", "Divergência entre documentos", False, "—"),
    ("Frequência semanal a comparar", "2x ou 5x por semana", "Cenários A-D", "Premissa provisória", True, "Exige regeneração (script)"),
    ("Nº de salas a comparar", "2 ou 3", "Cenários A-D", "Premissa provisória", True, "Exige regeneração (script)"),
    ("Gatilho para 1º curso técnico", "Após a 4ª turma de DH concluída", "Planilha operacional, anotação A34", "Premissa provisória", True, "Exige regeneração (script)"),
])

r = premise_block(ws, r, "OFICINA DE GESTÃO (OG) — DUAS VARIANTES, NENHUMA CONFIRMADA", [
    ("Total de turmas — OPÇÃO USADA NO MODELO", 15, "Planilha operacional J5=15 (1 OG por turma de DH concluída)", "Premissa provisória", True, "Exige regeneração (script)"),
    ("Total de turmas — LEITURA ALTERNATIVA", "≈7 a 8", "Anotação \"2 DH = 1 OG + 1 CT\" (A28) aplicada literalmente", "Divergência entre documentos", False, "—"),
    ("Variante B1 — espaço", "3 espaços ADICIONAIS, fora das 3 salas de DH", "NÃO CONFIRMADO. Ver comparação B1 x B2 na aba 07_CENARIO_B.", "Informação pendente de validação", False, "Exige regeneração (script)"),
    ("Variante B2 — espaço", "As mesmas 3 salas de DH (sem espaço adicional)", "OG só pode começar após o fim total da fase de DH (04/11) — ver aba 07.", "Decisão recomendada", False, "Exige regeneração (script)"),
    ("Carga horária / Encontros", "4h / 2 encontros", "Planilha operacional F5/G5", "Confirmada", True, "Exige regeneração (script)"),
    ("Mínimo/máximo de participantes", "15 / 18", "Planilha operacional H5/I5", "Confirmada", True, "Exige regeneração (script)"),
])

r = premise_block(ws, r, "CURSOS TÉCNICOS (CT) — DATAS SÃO PROPOSTA PARA SIMULAÇÃO", [
    ("Total de turmas — OPÇÃO USADA NO MODELO", 10, "Soma de J6:J12 (7 temas)", "Premissa provisória", True, "Exige regeneração (script)"),
    ("Total de turmas — META CONTRATUAL (PDF macro)", 13, "PDF, Núcleo 2 = 13 turmas = 195 pessoas", "Divergência entre documentos", False, "—"),
    ("Diferença a validar", "3 turmas", "Temas/fornecedores não definidos neste material", "Informação pendente de validação", False, "—"),
    ("Frequência / Turno", "5x/semana, 4h/dia, noite", "Planilha operacional e briefing", "Confirmada", True, "Exige regeneração (script)"),
    ("Quórum mínimo por turma", "14 a 18 (varia por curso/fornecedor)", "Planilha operacional H6:H12; reunião cita 16 como referência geral", "Divergência entre documentos", False, "—"),
    ("Classificação de TODAS as datas de CT nesta versão", "Proposta para simulação — pendente de quórum, contratação e validação do fornecedor", "Nenhuma data de CT é compromisso oficial (item 6 da revisão crítica)", "Condicionado à validação", False, "—"),
])

r = premise_block(ws, r, "CONSULTORIA INDIVIDUAL (CI) — CALCULADORA DE CAPACIDADE NA ABA 15", [
    ("Meta de pessoas / atendimentos por pessoa / horas por atendimento", "150 / 3 / 1h", "Nota J14 (450 atend./3) e PDF macro (\"13 turmas = 150 pessoas\")", "Confirmada", True, "Exige regeneração (script)"),
    ("Nº de consultores (padrão)", 2, "Editável na aba 15 — recalcula por fórmula", "Premissa provisória", True, "Automático (fórmula, aba 15)"),
    ("Horas de atendimento por consultor/dia (padrão)", 4, "Editável na aba 15 — recalcula por fórmula", "Premissa provisória", True, "Automático (fórmula, aba 15)"),
    ("Dias de atendimento por semana (padrão)", 5, "Editável na aba 15 — recalcula por fórmula", "Premissa provisória", True, "Automático (fórmula, aba 15)"),
    ("Atendimentos simultâneos (padrão)", 2, "Editável na aba 15 — recalcula por fórmula", "Premissa provisória", True, "Automático (fórmula, aba 15)"),
    ("Curso Técnico é pré-requisito para Consultoria Individual", "Sim", "PDF macro", "Confirmada", True, "Exige regeneração (script)"),
])

r = premise_block(ws, r, "CALENDÁRIO", [
    ("Feriados considerados", "7/9, 12/10, 2/11, 20/11, 25/12, 1/1", "Calendário municipal de Paulínia 2026", "Premissa provisória", False, "Exige regeneração (script)"),
    ("Recesso institucional — início", RECESSO_START, "Observado no cronograma anterior; não confirmado formalmente", "Informação pendente de validação", True, "Exige regeneração (script)"),
    ("Recesso institucional — fim", RECESSO_END, "Idem", "Informação pendente de validação", True, "Exige regeneração (script)"),
])

freeze(ws, "B7")
print("01_PREMISSAS ok")

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
  "Levantar com a coordenação quais 3 turmas adicionais fariam sentido e formalizar com os fornecedores antes de assumir 13 como meta. TODAS as datas de CT nesta V2 são propostas para simulação — ver aba 14.",
  "Divergência entre documentos"),
 ("Regra de formação do DH a partir das Palestras", "A planilha operacional traz duas anotações conflitantes: \"2 PS = 1 DH\" (fecha com 15 turmas) e \"4 OS = 1 DH\" (não fecha; \"OS\" não corresponde a nenhuma sigla da trilha).",
  "Proposta (xlsx), observações A27 e A30",
  "2 PS = 1 DH  vs.  4 OS = 1 DH",
  "Alto: define o ritmo de liberação de novos lotes de DH; usar a regra errada muda a data de abertura de cada turma.",
  "Adotar \"2 PS = 1 DH\" (única leitura que fecha com as 15 turmas confirmadas) e tratar \"4 OS = 1 DH\" como possível erro de digitação a confirmar.",
  "Divergência entre documentos"),
 ("Espaço físico para a Oficina de Gestão", "O cenário recomendado (B) usa 3 salas para o DH, mas essas 3 salas ficam ocupadas por lotes consecutivos durante toda a fase (sem vaga livre). A V1 deste arquivo assumiu 3 espaços ADICIONAIS sem confirmar isso — foi um erro de tratamento apontado na revisão crítica.",
  "Achado desta análise (nenhum dos 4 documentos originais cruzou a agenda de sala do DH com a da OG)",
  "B1: 3 salas de DH + 3 espaços adicionais (6 no total, NÃO confirmado)  vs.  B2: 3 salas compartilhadas (OG só após o fim total do DH, em 04/11)",
  "Muito alto: se B1 não for viabilizado, a OG só pode começar após 04/11 (variante B2), atrasando toda a cadeia até a Consultoria Individual em até ~2 semanas.",
  "Não tratar B1 como confirmado. Levar as duas alternativas (B1 x B2, comparadas na aba 07_CENARIO_B) à coordenação para decisão explícita sobre espaço físico antes de comunicar datas de OG.",
  "Divergência entre documentos"),
 ("Regra de composição da Oficina de Gestão e Curso Técnico", "A planilha operacional lista 15 turmas de OG (1 para cada DH), mas também traz a anotação \"2 DH = 1 OG + 1 CT\", que implicaria só ~7-8 turmas de OG para 15 DH.",
  "Proposta (xlsx), linha J5 vs. observação A28",
  "15 turmas de OG (1:1 com DH)  vs.  ~7-8 turmas de OG (1 a cada 2 DH)",
  "Muito alto: dobra ou reduz pela metade a necessidade de espaço/horário de OG, afetando toda a fase seguinte à DH.",
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
  "Muito alto: é a variável que mais muda a duração total da fase de DH.",
  "Confirmar com antecedência a disponibilidade real e exclusiva de salas à noite antes de comunicar datas ao Sebrae — ver comparativo de cenários na aba 05.",
  "Divergência entre documentos"),
 ("Recesso institucional", "A simulação de cenários bloqueia 21/12/2026 a 08/01/2027 como recesso, mas não há confirmação formal de que o Instituto/Sebrae adotarão esse recesso.",
  "Simulação de cenários (aba \"Premissas e Pendências\")",
  "21/12/2026 a 08/01/2027 (assumido)  vs.  nenhuma confirmação",
  "Médio: cenários que já se estendem para dezembro/janeiro ficam ainda mais espremidos ou mais longos conforme o recesso seja confirmado ou não.",
  "Confirmar a duração exata do recesso institucional 2026/2027 junto à coordenação administrativa.",
  "Informação pendente de validação"),
 ("Datas com ano incorreto no cronograma anterior", "A aba \"Paulínia\" do arquivo \"Possibilidades de Cronograma\" tem uma seção rotulada \"Janeiro de 2027\" cujas datas de calendário estão grafadas como 2026, e ao menos uma linha traz data de término anterior à de início.",
  "\"Possibilidades de Cronograma- Paulínia.xlsx\", aba Paulínia",
  "Rótulo \"Janeiro de 2027\" com datas grafadas 2026; DHP9 com fim (18/11) anterior ao início (05/12)",
  "Alto: se usado sem correção, o leitor pode aceitar datas no ano errado ou uma janela invertida.",
  "Não usar essa planilha como fonte de datas finais; tratá-la apenas como registro histórico da tentativa anterior.",
  "Divergência entre documentos"),
 ("Fornecedor do curso de Jardinagem", "O cronograma anterior da coordenação registra o fornecedor do curso de Jardinagem como \"?\" — não definido.",
  "\"Possibilidades de Cronograma- Paulínia.xlsx\", célula R48",
  "Fornecedor: \"?\"",
  "Alto: sem fornecedor confirmado, a turma de Jardinagem não pode ser efetivamente contratada nem ter data de início garantida.",
  "Priorizar a definição do fornecedor de Jardinagem antes de comunicar a data desta turma aos participantes.",
  "Informação pendente de validação"),
 ("Formação real de turma a partir das Palestras (quórum)", "O cronograma trata a formação de cada lote de DH como certa, mas nenhum documento confirma que toda Palestra realizada converte participantes suficientes em turma fechada.",
  "Nenhum dos 4 documentos discute risco de não-formação de turma",
  "Suposição implícita: 100% das Palestras geram quórum para o próximo lote de DH",
  "Alto: se um lote não atingir quórum, todo o restante da cadeia (a partir daquele lote) atrasa — ver simulação de atraso de 7 dias na aba 07_CENARIO_B.",
  "Adicionar marcos de validação de quórum antes de cada lote (aba 06-09) e tratar a formação de turma como \"condicionada\", não automática.",
  "Divergência entre documentos"),
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
print("02_DIVERGENCIAS ok")

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
    ("OFICINA DE\nGESTÃO\n(OG)", STAGE_COLORS["OG"], "15 turmas* (B1/B2)"),
    ("CURSO\nTÉCNICO\n(CT)", STAGE_COLORS["CT"], "10 turmas* (simulação)"),
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
ws.cell(r2, 2, "* Ver divergência quantitativa e status \"condicionado à validação\" nas abas 01, 02, 07 e 14.").font = FONT_NOTE
r2 += 2

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
    ("30 PS → 15 DH  (regra: 2 PS = 1 DH)", "Sim — 30 ÷ 2 = 15, bate exatamente com o total confirmado. Mas a formação de cada turma depende de quórum real — ver marcos de validação nas abas 06-09.", "Premissa provisória"),
    ("30 PS → 15 DH  (regra: 4 OS = 1 DH)", "Não — 30 ÷ 4 = 7,5, não bate com 15 turmas. Sigla \"OS\" também não existe na trilha.", "Divergência entre documentos"),
    ("4 DH concluídos → 2 turmas de Curso Técnico (~80 pessoas formando 2 turmas de ~15-18)", "Aritmeticamente aproximado; indica que nem todos os 80 formandos migram para o mesmo curso técnico. Datas de CT são propostas para simulação, não compromisso — ver aba 14.", "Premissa provisória"),
    ("2 DH = 1 OG + 1 CT", "Não fecha com as 15 turmas de OG citadas na planilha operacional (implicaria ~7-8 OG, não 15).", "Divergência entre documentos"),
    ("15 DH → 15 OG (1 para 1)", "Consistente com o total de 15 turmas de OG, mas contradiz a regra \"2 DH = 1 OG + 1 CT\". Depende ainda de qual variante de espaço (B1 ou B2) for adotada.", "Premissa provisória"),
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
    ws.row_dimensions[r2].height = 60
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
print("03_FLUXO_DA_TRILHA ok")

# ============================================================================
# 06-09 CENARIO A/B/C/D
# ============================================================================
SCENARIO_SHEET_NAMES = {"A":"06_CENARIO_A","B":"07_CENARIO_B","C":"08_CENARIO_C","D":"09_CENARIO_D"}
SCEN_ROWS = {}

for key in ["A","B","C","D"]:
    s = SCENARIOS[key]
    sheet_name = SCENARIO_SHEET_NAMES[key]
    ws = wb.create_sheet(sheet_name)
    set_col_widths(ws, [3, 10, 8, 14, 12, 12, 12, 10, 34, 3])
    tag = "★ MENOR DURAÇÃO — RECOMENDADO PARA NEGOCIAÇÃO (condicionado)" if key == RECOMMENDED else ""
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
        # V2: so sinaliza feriado se ele cair em um dia REAL de encontro (cadencia real da turma)
        hol_hits = holiday_hits_real(t["inicio"], t["fim"], t["cadencia"])
        if hol_hits:
            obs = (obs + "; " if obs else "") + "Encontro cai em feriado: " + ", ".join(d.strftime("%d/%m") for d in hol_hits)
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
        ws.row_dimensions[r].height = 26
        r += 1

    rng_ini = f"F{first_data_row}:F{last_data_row}"
    rng_fim = f"G{first_data_row}:G{last_data_row}"
    indic("Início da 1ª turma", f"=MIN({rng_ini})", "Marco de início da fase de Desenvolvimento Humano")
    indic("Conclusão da 4ª turma", f"=SMALL({rng_fim},4)", "Marco para liberar o primeiro lote de Curso Técnico (premissa: após 4 DH) — data PROPOSTA, condicionada a quórum real")
    indic("Conclusão da 15ª turma (todas)", f"=MAX({rng_fim})", "Fim estimado da fase de Desenvolvimento Humano")
    indic("Duração da fase DH (dias corridos)", f"={get_column_letter(3)}{indic_rows['Conclusão da 15ª turma (todas)']}-{get_column_letter(3)}{indic_rows['Início da 1ª turma']}+1", "Da 1ª turma iniciada até a 15ª concluída", is_date=False)
    indic("Turmas ativas na mesma semana (pico)", s["turmas_semana_pico"], "Nº de turmas com o período [início,término] tocando a mesma semana corrida — NÃO é o nº de turmas se encontrando na mesma noite.", is_date=False)
    indic("Máximo de turmas na MESMA noite (mesma sala/horário)", s["max_turmas_mesma_noite"], f"Limite físico real: {s['salas']} sala(s) → no máximo {s['salas']} atividades simultâneas por noite. O uso alternado de dias (ex.: Seg/Qua com Ter/Qui) permite mais turmas ATIVAS na semana, mas não aumenta o nº de salas físicas disponíveis por noite.", is_date=False)

    # ---- Marcos de validacao de quorum por lote ----
    r += 1
    ws.merge_cells(f"B{r}:I{r}")
    ws.cell(r,2,"MARCOS DE VALIDAÇÃO DE QUÓRUM POR LOTE — formação de turma é CONDICIONADA, não automática").font=FONT_H1
    ws.cell(r,2).fill = FILL_H1
    r += 1
    header_row(ws, r, 2, ["Lote","Data planejada\nde validação","Data-limite /\njanela de contingência","Impacto de atraso\nde 7 dias no lote","Risco de palestra\nreagendada","Status"])
    r += 1
    for L in s["lots"]:
        lot_start = s["lot_start"][L]
        val_date = lot_start - timedelta(days=7)
        while val_date.weekday() > 4:
            val_date -= timedelta(days=1)
        gap_to_next = s["lot_gaps"].get(L+1, None)
        limite_txt = f"{gap_to_next} dia(s) de folga antes de empurrar o lote {L+1}" if gap_to_next is not None else "Último lote — atraso aqui só afeta o fim da fase DH"
        sim = simulate_lot_delay(s, lot_to_delay=L, delay_days=7, absorb_gaps=True)
        quarta_txt = "sem mudança" if sim["atraso_liquido_quarta"] == 0 else f"+{sim['atraso_liquido_quarta']}d"
        fimfase_txt = "sem mudança" if sim["atraso_liquido_fim_fase"] == 0 else f"+{sim['atraso_liquido_fim_fase']}d"
        impacto_txt = f"4ª turma: {quarta_txt}; fim da fase DH: {fimfase_txt} (considerando a folga natural entre lotes)"
        body_cell(ws, r, 2, f"Lote {L} ({', '.join(t['turma'] for t in s['turmas'] if t['lote']==L)})", bold=True, wrap=True)
        dvc = body_cell(ws, r, 3, val_date, align="center"); dvc.number_format=DATE_FMT
        body_cell(ws, r, 4, limite_txt, wrap=True)
        body_cell(ws, r, 5, impacto_txt, wrap=True)
        body_cell(ws, r, 6, "Palestras deste lote podem precisar de reforço de mobilização/reagendamento se o quórum não fechar.", wrap=True)
        classification_cell(ws, r, 7, "Condicionado à validação")
        ws.row_dimensions[r].height = 46
        r += 1

    # ---- Simulacao de atraso de 7 dias no lote 1 ----
    r += 1
    ws.merge_cells(f"B{r}:I{r}")
    ws.cell(r,2,"SIMULAÇÃO DE IMPACTO — ATRASO DE 7 DIAS NO LOTE 1 (ex.: quórum insuficiente nas primeiras Palestras)").font=FONT_H1
    ws.cell(r,2).fill = FILL_H1
    r += 1
    header_row(ws, r, 2, ["Cenário de folga","4ª turma concluída\n(original)","4ª turma concluída\n(com atraso)","Fim da fase DH\n(original)","Fim da fase DH\n(com atraso)","Leitura"])
    r += 1
    for label, sim in [("Otimista — folgas naturais entre lotes absorvem parte do atraso", simulate_lot_delay(s,1,7,True)),
                        ("Conservador (pior caso) — nenhuma folga é aproveitada, atraso se propaga integralmente", simulate_lot_delay(s,1,7,False))]:
        body_cell(ws, r, 2, label, wrap=True, bold=True)
        c1=body_cell(ws, r, 3, sim["quarta_concl_original"], align="center"); c1.number_format=DATE_FMT
        c2=body_cell(ws, r, 4, sim["quarta_concl_novo"], align="center"); c2.number_format=DATE_FMT
        c3=body_cell(ws, r, 5, sim["fim_fase_original"], align="center"); c3.number_format=DATE_FMT
        c4=body_cell(ws, r, 6, sim["fim_fase_novo"], align="center"); c4.number_format=DATE_FMT
        leitura = f"4ª turma desloca {sim['atraso_liquido_quarta']}d; fim da fase DH desloca {sim['atraso_liquido_fim_fase']}d"
        body_cell(ws, r, 7, leitura, wrap=True)
        ws.row_dimensions[r].height = 34
        r += 1

    # ---- B1 x B2: as duas alternativas de espaco para a OG (somente na aba do cenario B) ----
    if key == "B":
        r += 1
        ws.merge_cells(f"B{r}:I{r}")
        ws.cell(r,2,"OFICINA DE GESTÃO — DUAS VARIANTES DE ESPAÇO (NENHUMA CONFIRMADA) — \"CENÁRIO B\" NÃO DEFINE ISSO SOZINHO").font=FONT_H1
        ws.cell(r,2).fill = FILL_H1
        r += 1
        ws.merge_cells(f"B{r}:I{r}")
        b1b2_note = ws.cell(r,2,"A V1 deste arquivo tratou 3 espaços adicionais de OG como certos. Nesta V2, isso é explicitado como uma "
                                 "escolha em aberto entre duas variantes — o nome do cenário muda conforme a escolha, porque a necessidade "
                                 "total de espaços físicos é diferente em cada uma.")
        b1b2_note.font = FONT_NOTE; b1b2_note.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 30
        r += 1
        header_row(ws, r, 2, ["Variante","Espaço usado","OG início","OG término","Conflitos de sala","Necessidade total de espaços","Risco operacional"])
        r += 1
        og_b1_ini, og_b1_fim = min(o["inicio"] for o in OG_B1), max(o["fim"] for o in OG_B1)
        og_b2_ini, og_b2_fim = min(o["inicio"] for o in OG_B2), max(o["fim"] for o in OG_B2)
        variantes = [
            ("Cenário B1 — 3 salas de DH + 3 espaços ADICIONAIS de OG", "3 espaços dedicados, distintos das 3 salas de DH (NÃO CONFIRMADO)",
             og_b1_ini, og_b1_fim, "Nenhum (espaços dedicados, verificado dia a dia)", "6 espaços no total (3 DH + 3 OG)",
             "Depende de viabilizar 3 espaços adicionais — sem eles, este plano não é executável."),
            ("Cenário B2 — 3 salas totais, compartilhadas entre DH e OG", "As mesmas 3 salas do DH (sem espaço adicional)",
             og_b2_ini, og_b2_fim, "Nenhum (OG só ocupa a sala após o fim total da fase DH, em ondas sequenciais)", "3 espaços no total (sem espaço adicional)",
             f"OG só começa em {og_b2_ini.strftime('%d/%m/%Y')} (após o fim da fase DH) e só termina em {og_b2_fim.strftime('%d/%m/%Y')} — atraso de "
             f"{(og_b2_ini - og_b1_ini).days} dias no início da OG frente à variante B1, empurrando também Curso Técnico e Consultoria Individual para quem depende da OG."),
        ]
        for nome, espaco, ini, fim, conflitos, necessidade, risco in variantes:
            body_cell(ws, r, 2, nome, bold=True, wrap=True)
            body_cell(ws, r, 3, espaco, wrap=True)
            c1 = body_cell(ws, r, 4, ini, align="center"); c1.number_format=DATE_FMT
            c2 = body_cell(ws, r, 5, fim, align="center"); c2.number_format=DATE_FMT
            body_cell(ws, r, 6, conflitos, wrap=True)
            body_cell(ws, r, 7, necessidade, wrap=True)
            body_cell(ws, r, 8, risco, wrap=True)
            ws.row_dimensions[r].height = 70
            r += 1
        r += 1
        ws.merge_cells(f"B{r}:I{r}")
        decisao_cell = ws.cell(r, 2, "DECISÃO NECESSÁRIA: a coordenação precisa optar formalmente entre B1 (mais rápido, mas depende de 3 espaços "
                                       "adicionais não confirmados) e B2 (não depende de espaço extra, mas atrasa a OG e tudo que depende dela). "
                                       "Nenhuma das duas está aprovada nesta versão.")
        decisao_cell.font = Font(bold=True, color=NAVY)
        decisao_cell.fill = PatternFill("solid", fgColor=LGOLD)
        decisao_cell.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        ws.row_dimensions[r].height = 40
        r += 1

    freeze(ws, f"B{first_data_row}")
    SCEN_ROWS[key] = (first_data_row, last_data_row, indic_rows)
    print(f"{sheet_name} ok")

# ============================================================================
# 05_COMPARATIVO_CENARIOS
# ============================================================================
from openpyxl.formatting.rule import FormulaRule

ws = wb.create_sheet("05_COMPARATIVO_CENARIOS")
set_col_widths(ws, [3, 4, 16, 8, 14, 13, 13, 13, 13, 13, 30, 34, 42, 3])
style_title(ws, "B2:M2", "COMPARATIVO DOS 4 CENÁRIOS DE DESENVOLVIMENTO HUMANO (DH)")
style_subtitle(ws, "B3:M3", "Salas x Frequência semanal — indicadores recalculados por fórmula a partir das abas 06 a 09. O destaque de \"cenário recomendado\" segue a célula editável em 01_PREMISSAS.")

# celula auxiliar de referencia ao cenario recomendado (para a formatacao condicional)
ws["O1"] = f"={CENARIO_CELL_ADDR}"
ws["O1"].font = Font(size=1, color="FFFFFF")

r = 5
headers = ["","Cenário","Salas","Frequência\nDH","Turmas ativas\nna semana (pico)","Máx. turmas\nmesma noite","Início do\nDH","4ª turma\nconcluída","15ª turma\nconcluída","Duração\n(dias)","Riscos principais","Vantagens / Desvantagens","Viabilidade (condicionada)"]
header_row(ws, r, 2, headers)
r += 1
scen_data_first_row = r

SCEN_META = {
 "A": {"riscos": "Sobreposição de turmas exige interleaving de dias (Seg/Qua com Ter/Qui) na mesma sala; termina em dezembro, próximo ao recesso.",
       "vant_desv": "Vantagem: usa a capacidade máxima das 3 salas. Desvantagem: ritmo mais lento (2x/semana) atrasa toda a cadeia OG→CT→EAD→CI."},
 "B": {"riscos": "Ritmo intenso (5x/semana) exige disponibilidade constante de facilitadores e das 3 salas em sequência apertada entre lotes; praticamente sem folga para atraso (ver simulação de 7 dias na aba 07). Espaço para a Oficina de Gestão ainda não definido entre as variantes B1/B2.",
       "vant_desv": "Vantagem: conclui a fase DH mais cedo (04/11), abrindo mais espaço para OG, CT, EAD e CI ainda em 2026. Desvantagem: menor margem de faltas/reposição por turma; parte dos Cursos Técnicos e da Consultoria Individual permanece programada para 2027 mesmo neste cenário."},
 "C": {"riscos": "Fase DH ultrapassa o recesso institucional e só termina em fevereiro/2027 — compromete seriamente o início dos cursos técnicos e a meta anual.",
       "vant_desv": "Vantagem: nenhuma relevante frente aos demais cenários. Desvantagem: menor capacidade (2 salas) e maior duração total (169 dias)."},
 "D": {"riscos": "Alta intensidade diária (5x/semana) combinada a apenas 2 salas deixa pouquíssima folga para reposição de encontros perdidos.",
       "vant_desv": "Vantagem: mais rápido que os cenários de 2x/semana. Desvantagem: capacidade mínima (2 salas) entre os 4 cenários."},
}
VIAB_TEXT = {
    "A": "Viável — condicionado à validação de salas e facilitadores.",
    "B": "Menor duração — recomendado para negociação, condicionado à validação de salas, facilitadores, adesão dos participantes e espaços para OG (ver B1 x B2 na aba 07).",
    "C": "Não recomendado — atravessa o recesso e estende a fase DH até fevereiro/2027.",
    "D": "Viável com ressalvas — condicionado à validação de salas, facilitadores e folga mínima entre lotes.",
}
VIAB_COLOR = {"A":LGOLD, "B":LGOLD, "C":LRED, "D":LGOLD}

for key in ["A","B","C","D"]:
    sh = SCENARIO_SHEET_NAMES[key]
    first, last, indic_rows = SCEN_ROWS[key]
    s = SCENARIOS[key]
    body_cell(ws, r, 2, key, align="center")  # coluna auxiliar oculta (chave do cenario)
    body_cell(ws, r, 3, f"Cenário {key}", bold=True)
    body_cell(ws, r, 4, s["salas"], align="center")
    body_cell(ws, r, 5, s["freq"], align="center")
    body_cell(ws, r, 6, s["turmas_semana_pico"], align="center")
    body_cell(ws, r, 7, s["max_turmas_mesma_noite"], align="center")
    c1 = ws.cell(r, 8, f"='{sh}'!C{indic_rows['Início da 1ª turma']}"); c1.number_format=DATE_FMT
    c2 = ws.cell(r, 9, f"='{sh}'!C{indic_rows['Conclusão da 4ª turma']}"); c2.number_format=DATE_FMT
    c3 = ws.cell(r, 10, f"='{sh}'!C{indic_rows['Conclusão da 15ª turma (todas)']}"); c3.number_format=DATE_FMT
    c4 = ws.cell(r, 11, f"='{sh}'!C{indic_rows['Duração da fase DH (dias corridos)']}")
    for cc in (c1,c2,c3,c4):
        cc.alignment=Alignment(horizontal="center"); cc.border=BORDER_ALL
    body_cell(ws, r, 12, SCEN_META[key]["riscos"], wrap=True)
    body_cell(ws, r, 13, SCEN_META[key]["vant_desv"], wrap=True)
    vcell = ws.cell(r, 14, VIAB_TEXT[key])
    vcell.font = Font(bold=True, size=9)
    vcell.fill = PatternFill("solid", fgColor=VIAB_COLOR[key])
    vcell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    vcell.border = BORDER_ALL
    ws.row_dimensions[r].height = 85
    r += 1
scen_data_last_row = r - 1

# formatacao condicional: destaca a LINHA cuja chave (coluna B) bate com o cenario recomendado (O1)
rng = f"B{scen_data_first_row}:N{scen_data_last_row}"
ws.conditional_formatting.add(
    rng,
    FormulaRule(formula=[f"$B{scen_data_first_row}=$O$1"], fill=PatternFill("solid", fgColor=LGREEN), stopIfTrue=False)
)

r += 1
ws.merge_cells(f"B{r}:N{r}")
rec_cell = ws.cell(r, 2, "LEITURA: o cenário destacado em verde segue a célula \"Cenário recomendado\" em 01_PREMISSAS (troque A/B/C/D lá para recalcular o destaque aqui e em 04). "
                          "Padrão atual: Cenário B — menor duração da fase DH (73 dias) entre os 4 cenários avaliados. Isso NÃO significa que o cenário está aprovado: "
                          "a recomendação é condicionada à validação de salas, facilitadores, adesão dos participantes às Palestras/DH e definição do espaço para a Oficina de Gestão (B1 x B2, aba 07).")
rec_cell.font = Font(bold=True, size=11, color=WHITE)
rec_cell.fill = PatternFill("solid", fgColor=TEAL)
rec_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
ws.row_dimensions[r].height = 60

freeze(ws, f"C{scen_data_first_row}")
print("05_COMPARATIVO_CENARIOS ok")

# ============================================================================
# 04_RESUMO_EXECUTIVO  (NIVEL 1 - visao de uma pagina)
# ============================================================================
ws = wb.create_sheet("04_RESUMO_EXECUTIVO")
ws.sheet_view.showGridLines = False
set_col_widths(ws, [2, 16, 16, 16, 16, 16, 16, 16, 16, 16, 2])
style_title(ws, "B2:J2", "PARTIU! APRENDER E EMPREENDER — PAULÍNIA (V2)", height=34)
style_subtitle(ws, "B3:J3", "Trilha completa: Palestra de Sensibilização → Desenvolvimento Humano → Oficina de Gestão → Curso Técnico → EAD → Consultoria Individual")

r = 5
ws.merge_cells(f"B{r}:J{r}")
pc = ws.cell(r,2,'PERÍODO TOTAL: 11/08/2026 até parte de 2027 (ver observação abaixo) — cenário definido em 01_PREMISSAS ("Cenário recomendado")')
pc.font=Font(size=12,bold=True,color=WHITE); pc.fill=PatternFill("solid",fgColor=TEAL)
pc.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True); ws.row_dimensions[r].height=30
r += 2

# Marcos principais - ligados por formula ao cenario recomendado (INDEX/MATCH em 05_COMPARATIVO_CENARIOS)
CMP = "'05_COMPARATIVO_CENARIOS'"
CMP_R1, CMP_R2 = scen_data_first_row, scen_data_last_row  # linhas com dados de cenario na aba 05 (bounded range - evita ref de coluna inteira)
MATCH_KEY = f"MATCH({CENARIO_CELL_ADDR},{CMP}!$B${CMP_R1}:$B${CMP_R2},0)"
marcos_formulas = [
    ("Início das Palestras", "=DATE(2026,8,11)", STAGE_COLORS["PS"]),
    ("Início do DH (cenário sel.)", f"=INDEX({CMP}!$H${CMP_R1}:$H${CMP_R2},{MATCH_KEY})", STAGE_COLORS["DH"]),
    ("4ª turma de DH concluída", f"=INDEX({CMP}!$I${CMP_R1}:$I${CMP_R2},{MATCH_KEY})", STAGE_COLORS["DH"]),
    ("Fim do DH (cenário sel.)", f"=INDEX({CMP}!$J${CMP_R1}:$J${CMP_R2},{MATCH_KEY})", STAGE_COLORS["DH"]),
    ("1º Curso Técnico (proposta p/ simulação)", CT_GATE, STAGE_COLORS["CT"]),
]
col = 2
for label, val, color in marcos_formulas:
    c1 = ws.cell(r, col, label)
    c1.font = Font(size=9, bold=True); c1.fill = PatternFill("solid", fgColor=color)
    c1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c1.border = BORDER_ALL
    ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col)
    c2 = ws.cell(r+1, col, val); c2.number_format = DATE_FMT
    c2.font = Font(size=11, bold=True); c2.alignment = Alignment(horizontal="center")
    c2.fill = PatternFill("solid", fgColor="FFFFFF"); c2.border = BORDER_ALL
    ws.row_dimensions[r].height = 34
    ws.row_dimensions[r+1].height = 20
    col += 1
r += 3
ws.merge_cells(f"B{r}:J{r}")
note1 = ws.cell(r, 2, "Os 4 primeiros marcos recalculam por fórmula conforme o \"Cenário recomendado\" escolhido em 01_PREMISSAS. O marco de "
                        "Curso Técnico segue fixo no Cenário B nesta versão (a grade completa de CT/OG/EAD/CI exige regeneração do arquivo para outro cenário — ver matriz em 00_LEIA-ME).")
note1.font = FONT_NOTE; note1.alignment = Alignment(wrap_text=True)
ws.row_dimensions[r].height = 26
r += 2

# Linha do tempo mensal
ws.merge_cells(f"B{r}:J{r}")
ws.cell(r,2,"LINHA DO TEMPO — CENÁRIO B (referência desta versão)").font=FONT_H1
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

header_row(ws, r, 2, ["Etapa"] + [mm.strftime("%b/%y").upper() for mm in months], fill=FILL_H2, font=FONT_H2)
r += 1

def month_overlap(d1, d2, mstart):
    mend = (date(mstart.year, mstart.month+1,1) - timedelta(days=1)) if mstart.month < 12 else date(mstart.year,12,31)
    return d1 <= mend and d2 >= mstart

stage_ranges = {
 "PS": (PS_SCHEDULE[0]["data"], PS_SCHEDULE[-1]["data"]),
 "DH": (REC["inicio_fase"], REC["fim_fase"]),
 "OG": (min(o["inicio"] for o in OG_B1), max(o["fim"] for o in OG_B1)),
 "CT": (min(c["inicio"] for c in CT_SCHEDULE), max(c["fim"] for c in CT_SCHEDULE)),
 "EAD": (EAD_SCHEDULE["inicio"], EAD_SCHEDULE["fim"]),
 "CI": (CI_SCHEDULE["inicio"], CI_CAPACITY_DEFAULT["data_fim_estimada"]),
}
stage_labels = {"PS":"Palestras de Sensibilização","DH":"Desenvolvimento Humano","OG":"Oficina de Gestão (B1, não confirmada)",
                "CT":"Cursos Técnicos (proposta p/ simulação)","EAD":"Curso EAD","CI":"Consultoria Individual (estimado)"}
for key in ["PS","DH","OG","CT","EAD","CI"]:
    d1, d2 = stage_ranges[key]
    body_cell(ws, r, 2, stage_labels[key], bold=True)
    for ci, mm in enumerate(months):
        cell = ws.cell(r, 3+ci)
        cell.border = BORDER_ALL
        if month_overlap(d1,d2,mm):
            cell.fill = PatternFill("solid", fgColor=STAGE_COLORS[key])
    ws.row_dimensions[r].height = 20
    r += 1
body_cell(ws, r, 2, "Recesso / feriados", bold=True)
for ci, mm in enumerate(months):
    cell = ws.cell(r, 3+ci)
    cell.border = BORDER_ALL
    if month_overlap(RECESSO_START, RECESSO_END, mm) or any(month_overlap(h,h,mm) for h in HOLIDAYS):
        cell.fill = PatternFill("solid", fgColor=STAGE_COLORS["REC"])
ws.row_dimensions[r].height = 20
r += 2

# Cenario recomendado - callout CORRIGIDO (item 4 e 8 da revisao critica)
ws.merge_cells(f"B{r}:J{r+3}")
rc = ws.cell(r, 2, "RECOMENDAÇÃO CONDICIONADA (NÃO É UMA APROVAÇÃO): Cenário B — 3 salas, DH 5x/semana.\n"
                     "Menor duração — recomendado para negociação, condicionado à validação de salas, facilitadores, adesão dos "
                     "participantes e espaços para a Oficina de Gestão (variantes B1 x B2, ver aba 07_CENARIO_B).\n"
                     "O cenário B conclui o Desenvolvimento Humano antes do recesso e antecipa o início das etapas seguintes. "
                     "Parte dos cursos técnicos e das consultorias permanece programada para 2027.")
rc.font = Font(size=11, bold=True, color=WHITE)
rc.fill = PatternFill("solid", fgColor=NAVY)
rc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
for rr in range(r, r+4): ws.row_dimensions[rr].height = 24
r += 5

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
note = ws.cell(r,2,"Datas construídas a partir de premissas explicitadas nas abas 01, 02 e 03 — sujeitas a confirmação formal do Sebrae/Instituto. "
                    "Nenhuma data de Curso Técnico ou Oficina de Gestão nesta versão é compromisso oficial.")
note.font = FONT_NOTE
note.alignment = Alignment(wrap_text=True)
ws.row_dimensions[r].height=24

ws.sheet_view.zoomScale = 90
freeze(ws, "B6")
print("04_RESUMO_EXECUTIVO ok")

# ============================================================================
# Lista consolidada de atividades (para 10_CRONOGRAMA_MESTRE e 17_GANTT)
# ============================================================================
def build_master_activities():
    acts = []
    acts.append({"etapa":"PS","nome":"Palestras de Sensibilização (PS 01-30)","inicio":PS_SCHEDULE[0]["data"],
                  "fim":PS_SCHEDULE[-1]["data"],"local":"A definir por comunidade","obs":"3x/semana (Ter/Qua/Qui), 2h cada — Confirmada"})
    for t in REC["turmas"]:
        acts.append({"etapa":"DH","nome":t["turma"],"inicio":t["inicio"],"fim":t["fim"],"local":t["sala"],
                     "obs":f"Cadência {t['cadencia']} — Cenário {RECOMMENDED} (referência desta versão)"})
    for og in OG_B1:
        acts.append({"etapa":"OG","nome":og["turma"],"inicio":og["inicio"],"fim":og["fim"],"local":og["sala"],
                     "obs":f"Predecessora: {og['dh_origem']} — Variante B1 (espaço ADICIONAL, NÃO confirmado; ver B2 na aba 07)"})
    for ct in CT_SCHEDULE:
        acts.append({"etapa":"CT","nome":f"{ct['sigla']} — {ct['tema']}","inicio":ct["inicio"],"fim":ct["fim"],
                     "local":ct["fornecedor"],"obs":"Data proposta para simulação — pendente de quórum, contratação e validação do fornecedor"})
    acts.append({"etapa":"EAD","nome":"Curso EAD — Planejamento, Marketing e Finanças","inicio":EAD_SCHEDULE["inicio"],
                 "fim":EAD_SCHEDULE["fim"],"local":"Plataforma online","obs":"6h / 3 módulos, autoinstrucional"})
    acts.append({"etapa":"CI","nome":"Consultoria Individual (rolling, capacidade calculada na aba 15)","inicio":CI_SCHEDULE["inicio"],
                 "fim":CI_CAPACITY_DEFAULT["data_fim_estimada"],"local":"Presencial/Online",
                 "obs":f"Meta {CI_SCHEDULE['meta_pessoas']} pessoas x {CI_SCHEDULE['atendimentos_por_pessoa']} atend. — estimativa com parâmetros padrão (ver aba 15)"})
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
set_col_widths(ws, [3, 8, 34, 14, 14, 22, 46, 3])
style_title(ws, "B2:G2", "CRONOGRAMA MESTRE DE PAULÍNIA — VISÃO COMPLETA DA EXECUÇÃO")
style_subtitle(ws, "B3:G3", f"Referência: Cenário {RECOMMENDED} + variante B1 de OG (NÃO confirmada). Parte dos Cursos Técnicos e da Consultoria Individual permanece programada para 2027 — ver 04_RESUMO_EXECUTIVO.")
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
    ws.row_dimensions[r].height = 22
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
print("10_CRONOGRAMA_MESTRE ok")

# ============================================================================
# 11_AGENDA_POR_SALA
# ============================================================================
ws = wb.create_sheet("11_AGENDA_POR_SALA")
set_col_widths(ws, [3, 32, 12, 14, 14, 14, 12, 30, 3])
style_title(ws, "B2:H2", "AGENDA POR SALA — SEM SOBREPOSIÇÃO CONFIRMADA (checagem dia a dia por cadência real)")
style_subtitle(ws, "B3:H3", f"Cenário {RECOMMENDED}: {REC['salas']} salas de DH. Duas variantes de OG mostradas lado a lado (B1: espaço adicional NÃO confirmado; B2: mesmas salas do DH).")

r = 5
ws.merge_cells(f"B{r}:H{r}")
ws.cell(r,2,"SALAS DE DH").font=FONT_H1
ws.cell(r,2).fill = FILL_H1
r += 1
header_row(ws, r, 2, ["Sala","Turma","Tipo","Início","Término","Cadência","Conflito de horário?","Observação"])
r += 1
for t in REC["turmas"]:
    body_cell(ws, r, 2, t["sala"], bold=True, align="center")
    body_cell(ws, r, 3, t["turma"], align="center")
    body_cell(ws, r, 4, "DH", align="center", fill=STAGE_COLORS["DH"])
    dic = body_cell(ws, r, 5, t["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 6, t["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 7, t["cadencia"], align="center")
    body_cell(ws, r, 8, "Não (verificado dia a dia por cadência)", align="center", fill=LGREEN)
    ws.row_dimensions[r].height = 18
    r += 1

r += 1
ws.merge_cells(f"B{r}:H{r}")
ws.cell(r,2,"VARIANTE B1 — OG EM ESPAÇOS ADICIONAIS (NÃO CONFIRMADO)").font=FONT_H1
ws.cell(r,2).fill = FILL_H1
r += 1
header_row(ws, r, 2, ["Espaço","Turma","Tipo","Início","Término","Cadência","Conflito de horário?","Observação"])
r += 1
for og in sorted(OG_B1, key=lambda o: (o["sala"], o["inicio"])):
    body_cell(ws, r, 2, og["sala"], bold=True, wrap=True)
    body_cell(ws, r, 3, og["turma"], align="center")
    body_cell(ws, r, 4, "OG", align="center", fill=STAGE_COLORS["OG"])
    dic = body_cell(ws, r, 5, og["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 6, og["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 7, "Seg-Sex (2 enc.)", align="center")
    body_cell(ws, r, 8, "Não (espaço dedicado)", align="center", fill=LGREEN)
    ws.row_dimensions[r].height = 18
    r += 1

r += 1
ws.merge_cells(f"B{r}:H{r}")
ws.cell(r,2,"VARIANTE B2 — OG NAS MESMAS 3 SALAS DE DH (após o fim total da fase DH)").font=FONT_H1
ws.cell(r,2).fill = FILL_H1
r += 1
header_row(ws, r, 2, ["Sala","Turma","Tipo","Início","Término","Cadência","Conflito de horário?","Observação"])
r += 1
for og in sorted(OG_B2, key=lambda o: (o["sala"], o["inicio"])):
    body_cell(ws, r, 2, og["sala"], bold=True, align="center")
    body_cell(ws, r, 3, og["turma"], align="center")
    body_cell(ws, r, 4, "OG", align="center", fill=STAGE_COLORS["OG"])
    dic = body_cell(ws, r, 5, og["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 6, og["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 7, "Seg-Sex (2 enc.)", align="center")
    body_cell(ws, r, 8, "Não (roda em ondas sequenciais)", align="center", fill=LGREEN)
    ws.row_dimensions[r].height = 18
    r += 1

r += 1
ws.merge_cells(f"B{r}:H{r}")
ws.cell(r,2,"LOCAIS EXTERNOS — CURSOS TÉCNICOS (fornecedores; datas propostas para simulação)").font=FONT_H1
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
        body_cell(ws, r, 8, "Data proposta — fornecedor não confirmado" if "?" in forn else "Data proposta para simulação", align="center")
        ws.row_dimensions[r].height = 26
        r += 1

freeze(ws, "B6")
print("11_AGENDA_POR_SALA ok")

# ============================================================================
# 12_TURMAS_DH  (Nivel 4 - inclui bloco de PS 01-30 + DH 01-15)
# ============================================================================
ws = wb.create_sheet("12_TURMAS_DH")
set_col_widths(ws, [3, 12, 14, 14, 16, 12, 12, 22, 16, 34, 3])
style_title(ws, "B2:J2", "VISÃO POR TURMA — PALESTRAS (PS) E DESENVOLVIMENTO HUMANO (DH)")
style_subtitle(ws, "B3:J3", f"Cenário {RECOMMENDED} (referência desta versão) — a formação de cada turma é CONDICIONADA à Palestra atingir quórum (ver marcos de validação na aba 07)")

r = 5
ws.merge_cells(f"B{r}:J{r}")
ws.cell(r,2,"PALESTRAS DE SENSIBILIZAÇÃO (PS 01 – PS 30)").font=FONT_H1
ws.cell(r,2).fill = PatternFill("solid", fgColor=STAGE_COLORS["PS"])
ws.row_dimensions[r].height=20
r += 1
header_row(ws, r, 2, ["Turma","Data","Dia da\nsemana","Carga\nhorária","Encontros","Nº aprox.\nparticipantes","Predecessora","Situação da informação"])
r += 1
DIA_PT = {0:"Segunda",1:"Terça",2:"Quarta",3:"Quinta",4:"Sexta",5:"Sábado",6:"Domingo"}
for p in PS_SCHEDULE:
    body_cell(ws, r, 2, p["turma"], bold=True)
    dc = body_cell(ws, r, 3, p["data"], align="center"); dc.number_format = DATE_FMT
    body_cell(ws, r, 4, DIA_PT[p["data"].weekday()], align="center")
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
for i, t in enumerate(REC["turmas"]):
    body_cell(ws, r, 2, t["turma"], bold=True)
    dic = body_cell(ws, r, 3, t["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 4, t["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 5, t["cadencia"], align="center")
    body_cell(ws, r, 6, "20h (premissa; alternativa 22h)", align="center")
    body_cell(ws, r, 7, 10, align="center")
    body_cell(ws, r, 8, t["sala"], align="center")
    body_cell(ws, r, 9, "20-40", align="center")
    pred = f"PS {2*i+1:02d} e PS {2*i+2:02d} (regra 2 PS = 1 DH — premissa; formação CONDICIONADA a quórum, lote {t['lote']})"
    c = ws.cell(r, 10); c.value = pred; c.font = FONT_NOTE; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BORDER_ALL
    ws.row_dimensions[r].height = 28
    r += 1

freeze(ws, "B7")
print("12_TURMAS_DH ok")

# ============================================================================
# 13_OFICINAS_GESTAO
# ============================================================================
ws = wb.create_sheet("13_OFICINAS_GESTAO")
set_col_widths(ws, [3, 34, 14, 14, 12, 12, 12, 12, 30, 3])
style_title(ws, "B2:I2", "OFICINA DE GESTÃO (OG) — DUAS VARIANTES, NENHUMA CONFIRMADA")
style_subtitle(ws, "B3:I3", "B1 (espaço adicional) x B2 (mesmas salas do DH) — ver comparativo completo (datas, conflitos, risco) na aba 07_CENARIO_B")

r = 5
ws.merge_cells(f"B{r}:I{r}")
ws.cell(r,2,"VARIANTE B1 — 3 ESPAÇOS ADICIONAIS DEDICADOS (NÃO CONFIRMADO)").font=FONT_H1
ws.cell(r,2).fill = FILL_H1
r += 1
header_row(ws, r, 2, ["Turma","Espaço","Início","Término","Carga\nhorária","Encontros","Nº aprox.\nparticipantes","Predecessora","Classificação"])
r += 1
for og in OG_B1:
    body_cell(ws, r, 2, og["turma"], bold=True)
    body_cell(ws, r, 3, og["sala"], wrap=True)
    dic = body_cell(ws, r, 4, og["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 5, og["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 6, "4h", align="center")
    body_cell(ws, r, 7, 2, align="center")
    body_cell(ws, r, 8, "15-18", align="center")
    body_cell(ws, r, 9, og["dh_origem"] + " concluída", align="center")
    ws.row_dimensions[r].height = 16
    r += 1

r += 1
ws.merge_cells(f"B{r}:I{r}")
ws.cell(r,2,"VARIANTE B2 — MESMAS 3 SALAS DO DH (após o fim total da fase de DH)").font=FONT_H1
ws.cell(r,2).fill = FILL_H1
r += 1
header_row(ws, r, 2, ["Turma","Sala","Início","Término","Carga\nhorária","Encontros","Nº aprox.\nparticipantes","Predecessora","Classificação"])
r += 1
for og in OG_B2:
    body_cell(ws, r, 2, og["turma"], bold=True)
    body_cell(ws, r, 3, og["sala"], align="center")
    dic = body_cell(ws, r, 4, og["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 5, og["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 6, "4h", align="center")
    body_cell(ws, r, 7, 2, align="center")
    body_cell(ws, r, 8, "15-18", align="center")
    body_cell(ws, r, 9, og["dh_origem"] + " concluída", align="center")
    ws.row_dimensions[r].height = 16
    r += 1
r += 1
ws.merge_cells(f"B{r}:I{r}")
note = ws.cell(r, 2, "Nenhuma das duas variantes está confirmada. Enquanto não houver decisão da coordenação sobre o espaço físico, "
                     "as demais abas (10, 11, 12) usam a variante B1 como referência de trabalho, por ser a mais rápida — não porque esteja aprovada.")
note.font = Font(italic=True, color=RED); note.fill = PatternFill("solid", fgColor=LRED)
note.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
ws.row_dimensions[r].height = 34
freeze(ws, "B6")
print("13_OFICINAS_GESTAO ok")

# ============================================================================
# 14_CURSOS_TECNICOS
# ============================================================================
ws = wb.create_sheet("14_CURSOS_TECNICOS")
set_col_widths(ws, [3, 9, 24, 12, 12, 12, 12, 15, 15, 22, 13, 13, 13, 13, 20, 3])
style_title(ws, "B2:O2", "CURSOS TÉCNICOS (CT) — TODAS AS DATAS SÃO PROPOSTA PARA SIMULAÇÃO")
style_subtitle(ws, "B3:O3", "Nenhuma data aqui é compromisso oficial — pendente de quórum, contratação e validação do fornecedor. 10 turmas confirmadas na planilha operacional; meta contratual do PDF macro é 13 (ver aba 02).")
r = 5
header_row(ws, r, 2, ["Sigla","Tema","Início\n(proposto)","Término\n(proposto)","Carga\nhorária","Encontros","DH concluídas\nantes da abertura","Participantes\nestimados disponíveis","Quórum\nmínimo","Fornecedor / Local","Fornecedor\nconfirmado?","Local\nconfirmado?","Data\nvalidada?","Responsável pela\nvalidação","Classificação"])
r += 1
for ct in CT_SCHEDULE:
    body_cell(ws, r, 2, ct["sigla"], bold=True)
    body_cell(ws, r, 3, ct["tema"], wrap=True)
    dic = body_cell(ws, r, 4, ct["inicio"], align="center"); dic.number_format=DATE_FMT
    dfc = body_cell(ws, r, 5, ct["fim"], align="center"); dfc.number_format=DATE_FMT
    body_cell(ws, r, 6, f"{ct['carga_horaria']}h", align="center")
    body_cell(ws, r, 7, ct["encontros"], align="center")
    body_cell(ws, r, 8, ct["dh_concluidas_antes"], align="center")
    body_cell(ws, r, 9, f"~{ct['participantes_estimados']} (estim.)", align="center")
    body_cell(ws, r, 10, f"{ct['min']}-{ct['max']}", align="center")
    fill_forn = LRED if "?" in ct["fornecedor"] else None
    body_cell(ws, r, 11, ct["fornecedor"], wrap=True, fill=fill_forn)
    body_cell(ws, r, 12, "Sim" if ct["fornecedor_confirmado"] else "Não", align="center", fill=(LGREEN if ct["fornecedor_confirmado"] else LRED))
    body_cell(ws, r, 13, "Sim" if ct["local_confirmado"] else "Não", align="center", fill=(LGREEN if ct["local_confirmado"] else LRED))
    body_cell(ws, r, 14, "Não" if not ct["data_validada"] else "Sim", align="center", fill=LRED)
    body_cell(ws, r, 15, ct["responsavel_validacao"], wrap=True)
    classification_cell(ws, r, 16, "Condicionado à validação")
    ws.row_dimensions[r].height = 40
    r += 1

r += 1
ws.merge_cells(f"B{r}:O{r}")
extra_note = ws.cell(r, 2, "PENDENTE DE VALIDAÇÃO: 3 turmas adicionais necessárias para atingir a meta contratual de 13 turmas do PDF macro (Núcleo 2). "
                            "Temas e fornecedores não definidos neste material — recomenda-se priorizar cursos com maior demanda identificada no diagnóstico de participantes. "
                            "\"Participantes estimados disponíveis\" é uma estimativa cumulativa (nº de turmas de DH concluídas × 30 pessoas em média) — não desconta quem já ocupou vaga em outro curso técnico.")
extra_note.font = Font(italic=True, bold=True, color=RED)
extra_note.fill = PatternFill("solid", fgColor=LRED)
extra_note.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[r].height = 46
freeze(ws, "B6")
print("14_CURSOS_TECNICOS ok")

# ============================================================================
# 15_EAD_CONSULTORIAS
# ============================================================================
ws = wb.create_sheet("15_EAD_CONSULTORIAS")
set_col_widths(ws, [3, 34, 16, 16, 16, 42, 3])
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
r += 2

# ---- Calculadora de capacidade da Consultoria Individual (formulas reais, editaveis) ----
ws.merge_cells(f"B{r}:F{r}")
ws.cell(r,2,"CALCULADORA DE CAPACIDADE — CONSULTORIA INDIVIDUAL (células amarelas são editáveis e recalculam sozinhas)").font=FONT_H1
ws.cell(r,2).fill = FILL_H1
r += 1

def calc_row(label, value, obs="", editable=True, numfmt=None, align="center"):
    global r
    body_cell(ws, r, 2, label, bold=True)
    c = ws.cell(r, 3, value)
    c.font=FONT_BODY; c.border=BORDER_ALL; c.alignment=Alignment(horizontal=align)
    if editable: c.fill = FILL_EDIT
    if numfmt: c.number_format = numfmt
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    body_cell(ws, r, 4, obs, wrap=True)
    ws.row_dimensions[r].height = 22
    addr = f"C{r}"
    r += 1
    return addr

a_meta        = calc_row("Meta de pessoas atendidas", CI_SCHEDULE["meta_pessoas"], "Cruzado entre nota da planilha operacional (450 atend./3) e PDF macro (13 turmas = 150 pessoas)")
a_atend_pp    = calc_row("Atendimentos por pessoa", CI_SCHEDULE["atendimentos_por_pessoa"], "Planilha operacional / nota J14")
a_horas_atend = calc_row("Horas por atendimento", CI_SCHEDULE["horas_por_atendimento"], "Planilha operacional linha 14 (E14 = 1h)")
a_consult     = calc_row("Nº de consultores", 2, "PARÂMETRO EDITÁVEL — quantos profissionais atendem em paralelo")
a_horas_dia   = calc_row("Horas de atendimento por consultor/dia", 4, "PARÂMETRO EDITÁVEL")
a_dias_sem    = calc_row("Dias de atendimento por semana", 5, "PARÂMETRO EDITÁVEL")
a_simult      = calc_row("Atendimentos simultâneos (capacidade física de salas/espaços)", 2, "PARÂMETRO EDITÁVEL — pode ser menor que o nº de consultores se faltar espaço")
a_inicio      = calc_row("Data de início da Consultoria Individual", CI_SCHEDULE["inicio"], "Uma semana após o primeiro Curso Técnico concluir (ver aba 14)", numfmt=DATE_FMT)
a_proposta    = calc_row("Data proposta para conclusão (meta da coordenação)", max(c["fim"] for c in CT_SCHEDULE), "EDITÁVEL — ajuste para a meta real da coordenação; aqui usamos o fim do último Curso Técnico como referência", numfmt=DATE_FMT)

r += 1
ws.merge_cells(f"B{r}:F{r}")
ws.cell(r,2,"RESULTADOS (calculados por fórmula)").font=FONT_H1
ws.cell(r,2).fill = FILL_H1
r += 1

def result_row(label, formula, obs="", numfmt=None, align="center"):
    global r
    body_cell(ws, r, 2, label, bold=True)
    c = ws.cell(r, 3, formula)
    c.font=Font(bold=True, color=NAVY); c.border=BORDER_ALL; c.alignment=Alignment(horizontal=align)
    c.fill = PatternFill("solid", fgColor=LTEAL)
    if numfmt: c.number_format = numfmt
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    body_cell(ws, r, 4, obs, wrap=True)
    ws.row_dimensions[r].height = 24
    addr = f"C{r}"
    r += 1
    return addr

r_total_atend = result_row("Total de atendimentos necessários", f"={a_meta}*{a_atend_pp}", "Meta de pessoas × atendimentos por pessoa")
r_total_horas = result_row("Total de horas necessárias", f"={r_total_atend}*{a_horas_atend}", "Total de atendimentos × horas por atendimento")
r_cap_paralela = result_row("Capacidade paralela efetiva", f"=MIN({a_consult},{a_simult})", "O menor entre nº de consultores e atendimentos simultâneos possíveis (espaço/sala)")
r_atend_semana = result_row("Atendimentos possíveis por semana", f"={r_cap_paralela}*{a_horas_dia}*{a_dias_sem}/{a_horas_atend}", "")
r_semanas = result_row("Semanas necessárias para concluir a meta", f"=ROUNDUP({r_total_atend}/{r_atend_semana},0)", "Arredondado para cima (semana parcial conta inteira)")
r_data_fim = result_row("Data de conclusão estimada (primeira data viável)", f"={a_inicio}+{r_semanas}*7", "", numfmt=DATE_FMT)
r_viavel = result_row("Cabe até a data proposta pela coordenação?",
                       f'=IF({r_data_fim}<={a_proposta},"SIM — cabe até a data proposta","NÃO — primeira data viável é "&TEXT({r_data_fim},"DD/MM/YYYY"))',
                       "Compara a data de conclusão estimada com a meta editável acima")

r += 1
ws.merge_cells(f"B{r}:F{r}")
note = ws.cell(r, 2, "Como usar: altere nº de consultores, horas/dia, dias/semana ou atendimentos simultâneos acima — os resultados recalculam "
                     "automaticamente. Este cálculo NÃO considera o bloqueio do recesso institucional nas semanas contadas (simplificação); "
                     "se o intervalo calculado cruzar o recesso, a data real de conclusão tende a ser um pouco mais tarde do que a fórmula indica.")
note.font = FONT_NOTE
note.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[r].height = 40

freeze(ws, "B6")
print("15_EAD_CONSULTORIAS ok")

# ============================================================================
# 16_RISCOS_E_DECISOES
# ============================================================================
ws = wb.create_sheet("16_RISCOS_E_DECISOES")
set_col_widths(ws, [3, 30, 26, 26, 26, 16, 16, 26, 3])
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
  "É a variável que mais impacta a duração da fase DH: 73 dias (cenário B, 3 salas) até 169 dias (cenário C, 2 salas).",
  "Confirmar contratualmente a disponibilidade exclusiva de 3 salas à noite, de segunda a sexta.",
  "Coordenação local / Espaço físico", "Aberto", "Usar o Cenário B como referência de negociação, mas manter D como contingência caso só 2 salas sejam viabilizadas."),
 ("Espaço físico para a Oficina de Gestão NÃO definido (B1 x B2)", "Achado desta análise — nenhum documento original cruzou a agenda de sala do DH com a da OG",
  "B1 precisa de 3 espaços adicionais (6 no total); B2 não precisa de espaço extra, mas atrasa a OG (e tudo que depende dela) até 04/11 ou depois.",
  "Decidir formalmente entre B1 e B2 (ver comparação completa na aba 07_CENARIO_B) antes de comunicar datas de OG a qualquer parceiro.",
  "Coordenação local / Espaço físico", "Aberto — decisão pendente", "Não apresentar nenhuma das duas variantes como definitiva até a decisão."),
 ("Regra de composição da Oficina de Gestão (15 x ~7-8 turmas)", "Proposta (xlsx), linha J5 vs. anotação A28",
  "Dobra ou reduz pela metade a necessidade de espaço/horário dedicado à OG (além da questão do espaço físico acima).",
  "Confirmar se cada DH gera sua própria OG (1:1) ou se as OG são formadas a partir de pares de turmas de DH.",
  "Coordenação pedagógica", "Aberto", "Assumir 1:1 (15 turmas) como premissa de trabalho até confirmação."),
 ("Formação de turma tratada como certa (sem risco de quórum)", "Nenhum dos 4 documentos discute o risco de uma Palestra não gerar quórum suficiente",
  "Um atraso de 7 dias no 1º lote de DH desloca a conclusão da 4ª turma entre +3 e +7 dias, dependendo de quanta folga natural existir entre lotes (ver simulação nas abas 06-09).",
  "Definir um protocolo de reagendamento/reforço de mobilização para Palestras que não atingirem quórum mínimo.",
  "Coordenação pedagógica / Facilitadores", "Aberto — novo", "Adotar os marcos de validação de quórum (aba 06-09) como checkpoint formal antes de cada lote."),
 ("Fornecedor de Jardinagem não definido", "\"Possibilidades de Cronograma\", célula R48 = \"?\"",
  "Turma de Jardinagem não pode ser comunicada aos participantes sem fornecedor e local confirmados.",
  "Selecionar e contratar fornecedor para o curso de Jardinagem.",
  "Coordenação de parcerias", "Aberto", "Priorizar a definição nas próximas 4 semanas, antes da 4ª turma de DH concluir."),
 ("Recesso institucional não confirmado formalmente", "Simulação de cenários (premissa observada)",
  "Cenários que se estendem para dezembro/janeiro podem precisar de ajuste fino de datas conforme o recesso real.",
  "Confirmar datas oficiais de recesso institucional 2026/2027.",
  "Administrativo / RH", "Aberto", "Manter 21/12/2026-08/01/2027 como premissa de planejamento; ajustar se o recesso oficial for diferente."),
 ("Prazo apertado até o início das Palestras", "Contexto do projeto — data atual 21/07/2026, início das PS em 11/08/2026",
  "Restam menos de 3 semanas até a primeira Palestra de Sensibilização; qualquer atraso na confirmação de salas/premissas comprime a mobilização.",
  "Validar premissas críticas (salas, carga horária do DH, espaço de OG) com urgência.",
  "Coordenação geral do projeto", "Aberto — urgente", "Tratar as pendências desta aba como prioridade das próximas 2 semanas."),
 ("Quórum mínimo dos cursos técnicos variável (14-18) vs. referência única de 16 citada em reunião", "Proposta (xlsx) vs. ata de reunião",
  "Se 16 for aplicado a todos os cursos, turmas com quórum planejado de 14-15 podem não abrir.",
  "Confirmar quórum mínimo por fornecedor/curso, não um valor único.",
  "Coordenação de parcerias", "Aberto", "Negociar quórum por curso com cada fornecedor antes de divulgar datas."),
 ("Capacidade da Consultoria Individual não verificada (V1)", "Ausente na V1 deste arquivo — apontado na revisão crítica",
  "Com os parâmetros padrão (2 consultores, 4h/dia, 5 dias/semana), a meta de 450 atendimentos leva ~12 semanas — pode ultrapassar a data que a coordenação tinha em mente.",
  "Definir nº real de consultores disponíveis e comparar com a calculadora da aba 15_EAD_CONSULTORIAS.",
  "Coordenação de parcerias / RH", "Aberto — novo", "Usar a calculadora da aba 15 para testar cenários de nº de consultores antes de comunicar prazo de conclusão da CI."),
 ("Recomendação do Cenário B tratada como aprovação (V1)", "Apontado na revisão crítica — texto da V1 usava \"Altamente viável\"",
  "Risco de a coordenação ou parceiros externos interpretarem o Cenário B como decisão fechada, sem as validações pendentes.",
  "Nenhuma — item já corrigido nesta V2 (ver aba 04 e 05).",
  "—", "Corrigido nesta V2", "Usar sempre o texto \"menor duração — recomendado para negociação, condicionado à validação...\" em qualquer material derivado deste arquivo."),
]
for risco, docs, impacto, decisao, resp, status, rec in RISCOS:
    body_cell(ws, r, 2, risco, bold=True, wrap=True)
    body_cell(ws, r, 3, docs, wrap=True)
    body_cell(ws, r, 4, impacto, wrap=True)
    body_cell(ws, r, 5, decisao, wrap=True)
    body_cell(ws, r, 6, resp, wrap=True, align="center")
    fill_status = LRED if ("urgente" in status or "novo" in status or "pendente" in status) else (LGREEN if "Corrigido" in status else LGOLD)
    body_cell(ws, r, 7, status, align="center", fill=fill_status, bold=True)
    body_cell(ws, r, 8, rec, wrap=True)
    ws.row_dimensions[r].height = 80
    r += 1

freeze(ws, "B6")
print("16_RISCOS_E_DECISOES ok")

# ============================================================================
# 17_GANTT_EXECUTIVO
# ============================================================================
ws = wb.create_sheet("17_GANTT_EXECUTIVO")
ws.sheet_view.showGridLines = False
style_title(ws, "B2:D2", "GANTT EXECUTIVO — TRILHA COMPLETA (CENÁRIO B + VARIANTE B1 DE OG)", height=26)
style_subtitle(ws, "B3:D3", "Grade semanal — cada coluna representa a segunda-feira da semana correspondente. Datas de CT são propostas para simulação.")

weeks = []
wk = PS_SCHEDULE[0]["data"]
wk = wk - timedelta(days=wk.weekday())
last_day = CI_CAPACITY_DEFAULT["data_fim_estimada"]
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
gantt_rows.append(("Oficina de Gestão B1 (agregado, NÃO confirmado)","OG", min(o["inicio"] for o in OG_B1), max(o["fim"] for o in OG_B1)))
for ct in CT_SCHEDULE:
    gantt_rows.append((f"  {ct['sigla']} — {ct['tema'][:22]} (proposta)","CT", ct["inicio"], ct["fim"]))
gantt_rows.append(("Curso EAD","EAD", EAD_SCHEDULE["inicio"], EAD_SCHEDULE["fim"]))
gantt_rows.append(("Consultoria Individual (estimado, ver aba 15)","CI", CI_SCHEDULE["inicio"], CI_CAPACITY_DEFAULT["data_fim_estimada"]))
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
print("17_GANTT_EXECUTIVO ok")

# ============================================================================
# 18_VISAO_PARA_PARCEIROS  (pagina unica, sem detalhe tecnico operacional)
# ============================================================================
ws = wb.create_sheet("18_VISAO_PARA_PARCEIROS")
ws.sheet_view.showGridLines = False
set_col_widths(ws, [2, 20, 20, 20, 20, 20, 20, 20, 20, 2])
style_title(ws, "B2:I2", "PARTIU! APRENDER E EMPREENDER — PAULÍNIA", height=32)
style_subtitle(ws, "B3:I3", "Visão para parceiros (Sebrae / Petrobras) — página única, sem detalhamento operacional")

r = 5
ws.merge_cells(f"B{r}:I{r}")
obj = ws.cell(r, 2, "OBJETIVO: organizar a trilha completa de capacitação em Paulínia (Núcleo 2) — da Palestra de Sensibilização "
                     "à Consultoria Individual — em um cronograma auditável, comparando alternativas de execução e deixando claro "
                     "o que já está confirmado e o que ainda depende de validação conjunta.")
obj.font = Font(size=11, bold=True, color=WHITE)
obj.fill = PatternFill("solid", fgColor=NAVY)
obj.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
ws.row_dimensions[r].height = 46
r += 2

ws.merge_cells(f"B{r}:I{r}")
ws.cell(r,2,"OS 4 CENÁRIOS AVALIADOS").font=FONT_H1
ws.cell(r,2).fill=FILL_H1
r += 1
header_row(ws, r, 2, ["Cenário","Salas","Frequência do DH","Duração da fase de DH","Leitura simples"])
r += 1
PARTNER_VIEW = [
    ("A", 3, "2x por semana", "≈101 dias", "Viável, mas mais lento e mais próximo do recesso."),
    ("B", 3, "5x por semana", "≈73 dias", "Menor duração — recomendado para negociação, condicionado à validação."),
    ("C", 2, "2x por semana", "≈169 dias", "Não recomendado — atravessa o recesso e vai até fevereiro/2027."),
    ("D", 2, "5x por semana", "≈116 dias", "Alternativa de contingência, com pouca folga."),
]
for k, salas, freq, dur, leitura in PARTNER_VIEW:
    fill = LGREEN if k=="B" else None
    body_cell(ws, r, 2, f"Cenário {k}", bold=True, fill=fill)
    body_cell(ws, r, 3, salas, align="center", fill=fill)
    body_cell(ws, r, 4, freq, align="center", fill=fill)
    body_cell(ws, r, 5, dur, align="center", fill=fill)
    body_cell(ws, r, 6, leitura, wrap=True, fill=fill)
    ws.row_dimensions[r].height = 26
    r += 1
r += 1

ws.merge_cells(f"B{r}:I{r}")
reco = ws.cell(r, 2, "RECOMENDAÇÃO CONDICIONADA: Cenário B (3 salas, DH 5x/semana) — menor duração, recomendado para negociação, "
                      "condicionado à validação de salas, facilitadores, adesão dos participantes e definição do espaço para a "
                      "Oficina de Gestão. Não é uma decisão fechada.")
reco.font = Font(size=10, bold=True, color=WHITE)
reco.fill = PatternFill("solid", fgColor=TEAL)
reco.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
ws.row_dimensions[r].height = 40
r += 2

ws.merge_cells(f"B{r}:I{r}")
ws.cell(r,2,"LINHA DO TEMPO MACRO").font=FONT_H1
ws.cell(r,2).fill=FILL_H1
r += 1
macro_phases = [
    ("Palestras + Desenvolvimento Humano", date(2026,8,11), date(2026,11,4), STAGE_COLORS["DH"]),
    ("Oficina de Gestão + Cursos Técnicos + EAD", date(2026,9,11), date(2027,2,4), STAGE_COLORS["CT"]),
    ("Consultoria Individual", CI_SCHEDULE["inicio"], CI_CAPACITY_DEFAULT["data_fim_estimada"], STAGE_COLORS["CI"]),
]
months_p = []
mm = date(2026,8,1)
while mm < date(2027,4,1):
    months_p.append(mm)
    mm = date(mm.year+1,1,1) if mm.month==12 else date(mm.year, mm.month+1, 1)
header_row(ws, r, 2, ["Macro-fase"] + [m.strftime("%b/%y").upper() for m in months_p], fill=FILL_H2, font=FONT_H2)
r += 1
def m_overlap(d1,d2,mstart):
    mend = (date(mstart.year, mstart.month+1,1)-timedelta(days=1)) if mstart.month<12 else date(mstart.year,12,31)
    return d1<=mend and d2>=mstart
for label, d1, d2, color in macro_phases:
    body_cell(ws, r, 2, label, bold=True, wrap=True)
    for ci, mm2 in enumerate(months_p):
        cell = ws.cell(r, 3+ci); cell.border=BORDER_ALL
        if m_overlap(d1,d2,mm2):
            cell.fill = PatternFill("solid", fgColor=color)
    ws.row_dimensions[r].height = 22
    r += 1
r += 1

ws.merge_cells(f"B{r}:I{r}")
ws.cell(r,2,"RECURSOS EXIGIDOS").font=FONT_H1
ws.cell(r,2).fill=FILL_H1
r += 1
recursos = [
    "3 salas dedicadas para o Desenvolvimento Humano, disponíveis à noite, de segunda a sexta.",
    "Espaço adicional para a Oficina de Gestão (variante em discussão — pode ser 3 espaços extras ou o reaproveitamento das mesmas 3 salas, com atraso).",
    "Fornecedores contratados para os cursos técnicos (SENAI, SENAC, Fundo Social e um fornecedor de Jardinagem ainda não definido).",
    "Consultores para a Consultoria Individual — capacidade a confirmar (ver premissas internas).",
]
for item in recursos:
    ws.merge_cells(f"B{r}:I{r}")
    ic = ws.cell(r, 2, "•  " + item); ic.font=FONT_BODY; ic.alignment=Alignment(wrap_text=True, vertical="center", indent=1)
    ws.row_dimensions[r].height = 22
    r += 1
r += 1

ws.merge_cells(f"B{r}:I{r}")
ws.cell(r,2,"DECISÕES PENDENTES").font=FONT_H1
ws.cell(r,2).fill=FILL_H1
r += 1
decisoes = [
    "Confirmar disponibilidade de 3 salas para o Desenvolvimento Humano.",
    "Definir a carga horária oficial do Desenvolvimento Humano (20h ou 22h).",
    "Decidir o espaço físico da Oficina de Gestão (espaço adicional ou salas compartilhadas com o DH).",
    "Confirmar se a meta de Cursos Técnicos em Paulínia é 10 ou 13 turmas.",
    "Confirmar fornecedor do curso de Jardinagem.",
]
for item in decisoes:
    ws.merge_cells(f"B{r}:I{r}")
    ic = ws.cell(r, 2, "•  " + item); ic.font=FONT_BODY; ic.alignment=Alignment(wrap_text=True, vertical="center", indent=1)
    ws.row_dimensions[r].height = 22
    r += 1
r += 1

ws.merge_cells(f"B{r}:I{r}")
aviso = ws.cell(r, 2, "AVISO: as datas técnicas apresentadas neste material dependem de validação de quórum, contratação de "
                       "fornecedores e confirmação de espaços físicos. Nenhuma data aqui é compromisso contratual.")
aviso.font = Font(bold=True, italic=True, color=RED)
aviso.fill = PatternFill("solid", fgColor=LRED)
aviso.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
ws.row_dimensions[r].height = 34

freeze(ws, "B6")
print("18_VISAO_PARA_PARCEIROS ok")

# ============================================================================
# 19_LOG_DE_ALTERACOES
# ============================================================================
ws = wb.create_sheet("19_LOG_DE_ALTERACOES")
set_col_widths(ws, [3, 8, 30, 46, 16, 3])
style_title(ws, "B2:E2", "LOG DE ALTERAÇÕES — V1 → V2")
style_subtitle(ws, "B3:E3", "Item a item, em resposta à revisão crítica solicitada pela coordenação")
r = 5
header_row(ws, r, 2, ["#","Item da revisão crítica","O que foi alterado nesta V2","Abas afetadas"])
r += 1
LOG = [
 (1, "Automação e premissas", "Removida a afirmação de que todas as premissas atualizam sozinhas. Adicionada uma matriz em 00_LEIA-ME e uma coluna \"Automação\" em 01_PREMISSAS "
     "mostrando o que é automático (fórmula) e o que exige rodar scripts/build_workbook.py. O \"Cenário recomendado\" agora é um dropdown que alimenta por fórmula "
     "(INDEX/MATCH) os destaques e marcos das abas 04 e 05.", "00, 01, 04, 05"),
 (2, "Capacidade de salas", "\"Capacidade simultânea\" foi substituída por dois indicadores: \"Turmas ativas na mesma semana (pico)\" e \"Máximo de turmas na mesma noite\" "
     "(igual ao nº de salas, por construção). Texto explica que o interleaving de dias aumenta turmas ativas, não o nº de salas físicas.", "05, 06, 07, 08, 09"),
 (3, "Oficina de Gestão e espaços adicionais", "Criadas as variantes B1 (3 salas de DH + 3 espaços adicionais, NÃO confirmado) e B2 (3 salas compartilhadas, OG só após o fim total do DH). "
     "Comparação completa (datas, conflitos, necessidade de espaços, risco) na aba 07_CENARIO_B. Nenhuma das duas é tratada como confirmada.", "01, 07, 10, 11, 13, 16, 19"),
 (4, "Viabilidade do cenário recomendado", "\"Altamente viável\" substituído por \"Menor duração — recomendado para negociação, condicionado à validação de salas, facilitadores, "
     "adesão dos participantes e espaços para OG.\" Destaque de linha mudou de verde (aprovado) para o texto condicionado; título das abas 06-09 não usa mais \"recomendado\" isolado.", "04, 05, 06-09"),
 (5, "Quórum e contingência", "Adicionado, em cada cenário (06-09), um bloco \"Marcos de validação de quórum por lote\" (data planejada, janela de contingência, impacto de atraso "
     "de 7 dias, risco, status \"Condicionado à formação de turma\") e uma \"Simulação de atraso de 7 dias\" com dois cenários (otimista/conservador).", "03, 06, 07, 08, 09, 16"),
 (6, "Cursos técnicos", "Todas as datas de CT reclassificadas como \"Data proposta para simulação — pendente de quórum, contratação e validação do fornecedor\". "
     "Adicionadas colunas: DH concluídas antes da abertura, participantes estimados, quórum mínimo, fornecedor confirmado?, local confirmado?, data validada?, responsável.", "01, 14"),
 (7, "Consultorias individuais", "Criada calculadora de capacidade com fórmulas reais e parâmetros editáveis (nº de consultores, horas/dia, dias/semana, atendimentos simultâneos). "
     "Calcula semanas necessárias e data de conclusão estimada, e compara com uma data proposta editável.", "01, 15"),
 (8, "Texto do resumo executivo", "Substituído o texto que dava a entender que tudo terminava antes do recesso pelo texto: \"O cenário B conclui o Desenvolvimento Humano antes do "
     "recesso e antecipa o início das etapas seguintes. Parte dos cursos técnicos e das consultorias permanece programada para 2027.\"", "04, 18"),
 (9, "Feriados", "A checagem de feriado nos cenários 2x/semana (A e C) agora verifica se o feriado cai EXATAMENTE em um dia real de encontro da cadência da turma "
     "(ex.: Ter/Qui), e não apenas se está dentro do intervalo [início,término].", "06, 08 (e 07, 09 por consistência)"),
 (10, "Versão para parceiros", "Criada a aba 18_VISAO_PARA_PARCEIROS: página única com objetivo, os 4 cenários, recomendação condicionada, linha do tempo macro, "
      "recursos exigidos, decisões pendentes e aviso de que as datas dependem de validação — sem detalhe técnico operacional.", "18"),
 (11, "Auditoria final", "Repetida a checagem de conflitos reais de sala (dia a dia, por cadência), dependências entre etapas, capacidade das consultorias e "
      "rastreamento de toda data provisória. Nenhuma informação não confirmada aparece como fato nesta versão. Este log resume o resultado.", "19 (esta aba)"),
]
for num, item, mudou, abas in LOG:
    body_cell(ws, r, 2, num, align="center", bold=True)
    body_cell(ws, r, 3, item, bold=True, wrap=True)
    body_cell(ws, r, 4, mudou, wrap=True)
    body_cell(ws, r, 5, abas, wrap=True, align="center")
    ws.row_dimensions[r].height = 85
    r += 1

freeze(ws, "B6")
print("19_LOG_DE_ALTERACOES ok")

# ============================================================================
# REORDENAR ABAS NA ORDEM LOGICA FINAL
# ============================================================================
FINAL_ORDER = [
    "00_LEIA-ME", "01_PREMISSAS", "02_DIVERGENCIAS", "03_FLUXO_DA_TRILHA",
    "04_RESUMO_EXECUTIVO", "05_COMPARATIVO_CENARIOS",
    "06_CENARIO_A", "07_CENARIO_B", "08_CENARIO_C", "09_CENARIO_D",
    "10_CRONOGRAMA_MESTRE", "11_AGENDA_POR_SALA", "12_TURMAS_DH",
    "13_OFICINAS_GESTAO", "14_CURSOS_TECNICOS", "15_EAD_CONSULTORIAS",
    "16_RISCOS_E_DECISOES", "17_GANTT_EXECUTIVO", "18_VISAO_PARA_PARCEIROS",
    "19_LOG_DE_ALTERACOES",
]
assert set(FINAL_ORDER) == set(wb.sheetnames), f"Divergencia: {set(FINAL_ORDER) ^ set(wb.sheetnames)}"
wb._sheets = [wb[name] for name in FINAL_ORDER]

TAB_COLORS = {
    "00_LEIA-ME":"808080","01_PREMISSAS":"BF8F00","02_DIVERGENCIAS":"C0392B","03_FLUXO_DA_TRILHA":"1F3864",
    "04_RESUMO_EXECUTIVO":"0F6B5C","05_COMPARATIVO_CENARIOS":"2E5395",
    "06_CENARIO_A":"9DC3E6","07_CENARIO_B":"2E7D32","08_CENARIO_C":"9DC3E6","09_CENARIO_D":"9DC3E6",
    "10_CRONOGRAMA_MESTRE":"1F3864","11_AGENDA_POR_SALA":"2E5395","12_TURMAS_DH":"2E5395",
    "13_OFICINAS_GESTAO":"2E5395","14_CURSOS_TECNICOS":"2E5395","15_EAD_CONSULTORIAS":"2E5395",
    "16_RISCOS_E_DECISOES":"C0392B","17_GANTT_EXECUTIVO":"1F3864","18_VISAO_PARA_PARCEIROS":"0F6B5C",
    "19_LOG_DE_ALTERACOES":"808080",
}
for name, color in TAB_COLORS.items():
    wb[name].sheet_properties.tabColor = color

wb.active = wb.sheetnames.index("00_LEIA-ME")

OUT_DIR = "/home/user/CRONOGRAMA/cronograma"
os.makedirs(OUT_DIR, exist_ok=True)
OUT_PATH = os.path.join(OUT_DIR, "Cronograma_Paulinia_Partiu_2026_V2.xlsx")
wb.save(OUT_PATH)
print("Arquivo salvo em:", OUT_PATH)
