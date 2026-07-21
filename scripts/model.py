# -*- coding: utf-8 -*-
"""
Motor de datas - PARTIU! Paulinia
Reaproveita a logica de lotes DH da Simulacao_Cronograma_Paulinia_Cenarios.xlsx
(ja auditada: sem conflito real de sala/dia) e estende para OG, CT, EAD, CI.
"""
from datetime import date, timedelta
import json

HOLIDAYS = {
    date(2026,9,7): "Independencia do Brasil",
    date(2026,10,12): "Nossa Senhora Aparecida",
    date(2026,11,2): "Finados",
    date(2026,11,20): "Consciencia Negra",
    date(2026,12,25): "Natal",
    date(2027,1,1): "Confraternizacao Universal",
}
RECESSO_START = date(2026,12,21)
RECESSO_END = date(2027,1,8)

def is_blocked(d):
    return d in HOLIDAYS or (RECESSO_START <= d <= RECESSO_END)

def add_working_slots(start, n_encontros, weekdays, skip_blocked=True):
    """Retorna lista de datas ocupando 'weekdays' (0=Seg..4=Sex) a partir de start, contando n_encontros."""
    dates = []
    d = start
    while len(dates) < n_encontros:
        if d.weekday() in weekdays and not (skip_blocked and is_blocked(d)):
            dates.append(d)
        d += timedelta(days=1)
    return dates

WD = {'Seg':0,'Ter':1,'Qua':2,'Qui':3,'Sex':4,'Sab':5,'Dom':6}

def cadence_weekdays(cad):
    """Converte string de cadencia ('Ter/Qui', 'Seg/Qua', 'Seg-Sex') em conjunto de dias da semana (0=Seg)."""
    if cad in ("Seg-Sex", "Seg a Sex"):
        return {0,1,2,3,4}
    return {WD[p] for p in cad.split("/")}

def holiday_hits_real(inicio, fim, cadencia, holidays=None):
    """Retorna apenas feriados que caem EXATAMENTE em um dia real de encontro da turma
    (data dentro do intervalo E dia da semana pertence a cadencia) -- nao apenas 'dentro do intervalo'."""
    hs = holidays if holidays is not None else HOLIDAYS
    wds = cadence_weekdays(cadencia)
    return sorted(d for d in hs if inicio <= d <= fim and d.weekday() in wds)

# ---------------------------------------------------------------------------
# 1. PALESTRAS DE SENSIBILIZACAO (PS) - CONFIRMADO por 3 fontes
# ---------------------------------------------------------------------------
PS_START = date(2026,8,11)   # terca-feira
PS_END_TARGET = date(2026,10,15)
PS_TOTAL = 30
PS_PER_WEEK = 3
PS_WEEKDAYS = {1,2,3}  # Ter/Qua/Qui (padrao observado na planilha "Possibilidades": E/F/G = Ter/Qua/Qui) -> bate com termino real em 15/10/2026

def build_ps_schedule():
    dates = add_working_slots(PS_START, PS_TOTAL, PS_WEEKDAYS, skip_blocked=False)
    ps = []
    for i, d in enumerate(dates, start=1):
        ps.append({"turma": f"PS {i:02d}", "data": d})
    return ps

PS_SCHEDULE = build_ps_schedule()

# ---------------------------------------------------------------------------
# 2. DESENVOLVIMENTO HUMANO (DH) - 4 CENARIOS
#    Reaproveita exatamente as datas ja calculadas e auditadas na simulacao
#    anterior (sem conflito real de sala, verificado por dia da semana).
# ---------------------------------------------------------------------------
# Formato: (turma, lote, sala, cadencia, inicio, fim, n_encontros)
SCENARIOS_RAW = {
 "A": {  # 3 salas, DH 2x/semana
   "salas": 3, "freq": "2x/semana", "leitura": "Viavel, mas com maior sobreposicao e termino em dezembro.",
   "indicacao": "Alternativa equilibrada",
   "turmas": [
    ("DH 1",1,"Sala 1","Ter/Qui","2026-08-25","2026-09-24"),
    ("DH 2",1,"Sala 2","Ter/Qui","2026-08-25","2026-09-24"),
    ("DH 3",1,"Sala 3","Ter/Qui","2026-08-25","2026-09-24"),
    ("DH 4",2,"Sala 1","Seg/Qua","2026-09-09","2026-10-14"),
    ("DH 5",2,"Sala 2","Seg/Qua","2026-09-09","2026-10-14"),
    ("DH 6",2,"Sala 3","Seg/Qua","2026-09-09","2026-10-14"),
    ("DH 7",3,"Sala 1","Ter/Qui","2026-09-29","2026-10-29"),
    ("DH 8",3,"Sala 2","Ter/Qui","2026-09-29","2026-10-29"),
    ("DH 9",3,"Sala 3","Ter/Qui","2026-09-29","2026-10-29"),
    ("DH 10",4,"Sala 1","Seg/Qua","2026-10-19","2026-11-23"),
    ("DH 11",4,"Sala 2","Seg/Qua","2026-10-19","2026-11-23"),
    ("DH 12",4,"Sala 3","Seg/Qua","2026-10-19","2026-11-23"),
    ("DH 13",5,"Sala 1","Ter/Qui","2026-11-03","2026-12-03"),
    ("DH 14",5,"Sala 2","Ter/Qui","2026-11-03","2026-12-03"),
    ("DH 15",5,"Sala 3","Ter/Qui","2026-11-03","2026-12-03"),
   ]},
 "B": {  # 3 salas, DH 5x/semana
   "salas": 3, "freq": "5x/semana", "leitura": "Maior velocidade; acompanha melhor a formacao das turmas.",
   "indicacao": "Recomendado para negociacao",
   "turmas": [
    ("DH 1",1,"Sala 1","Seg-Sex","2026-08-24","2026-09-04"),
    ("DH 2",1,"Sala 2","Seg-Sex","2026-08-24","2026-09-04"),
    ("DH 3",1,"Sala 3","Seg-Sex","2026-08-24","2026-09-04"),
    ("DH 4",2,"Sala 1","Seg-Sex","2026-09-08","2026-09-21"),
    ("DH 5",2,"Sala 2","Seg-Sex","2026-09-08","2026-09-21"),
    ("DH 6",2,"Sala 3","Seg-Sex","2026-09-08","2026-09-21"),
    ("DH 7",3,"Sala 1","Seg-Sex","2026-09-22","2026-10-05"),
    ("DH 8",3,"Sala 2","Seg-Sex","2026-09-22","2026-10-05"),
    ("DH 9",3,"Sala 3","Seg-Sex","2026-09-22","2026-10-05"),
    ("DH 10",4,"Sala 1","Seg-Sex","2026-10-06","2026-10-20"),
    ("DH 11",4,"Sala 2","Seg-Sex","2026-10-06","2026-10-20"),
    ("DH 12",4,"Sala 3","Seg-Sex","2026-10-06","2026-10-20"),
    ("DH 13",5,"Sala 1","Seg-Sex","2026-10-21","2026-11-04"),
    ("DH 14",5,"Sala 2","Seg-Sex","2026-10-21","2026-11-04"),
    ("DH 15",5,"Sala 3","Seg-Sex","2026-10-21","2026-11-04"),
   ]},
 "C": {  # 2 salas, DH 2x/semana
   "salas": 2, "freq": "2x/semana", "leitura": "Baixa capacidade; atravessa o recesso e termina em fevereiro.",
   "indicacao": "Nao recomendado",
   "turmas": [
    ("DH 1",1,"Sala 1","Ter/Qui","2026-08-25","2026-09-24"),
    ("DH 2",1,"Sala 2","Ter/Qui","2026-08-25","2026-09-24"),
    ("DH 3",1,"Sala 1","Seg/Qua","2026-08-24","2026-09-28"),
    ("DH 4",2,"Sala 2","Seg/Qua","2026-09-09","2026-10-14"),
    ("DH 5",2,"Sala 1","Ter/Qui","2026-09-29","2026-10-29"),
    ("DH 6",2,"Sala 2","Ter/Qui","2026-09-29","2026-10-29"),
    ("DH 7",3,"Sala 1","Seg/Qua","2026-09-30","2026-11-09"),
    ("DH 8",3,"Sala 2","Seg/Qua","2026-10-19","2026-11-23"),
    ("DH 9",3,"Sala 1","Ter/Qui","2026-11-03","2026-12-03"),
    ("DH 10",4,"Sala 2","Ter/Qui","2026-11-03","2026-12-03"),
    ("DH 11",4,"Sala 1","Seg/Qua","2026-11-11","2026-12-14"),
    ("DH 12",4,"Sala 2","Seg/Qua","2026-11-25","2027-01-18"),
    ("DH 13",5,"Sala 1","Ter/Qui","2026-12-08","2027-01-28"),
    ("DH 14",5,"Sala 2","Ter/Qui","2026-12-08","2027-01-28"),
    ("DH 15",5,"Sala 1","Seg/Qua","2026-12-16","2027-02-08"),
   ]},
 "D": {  # 2 salas, DH 5x/semana
   "salas": 2, "freq": "5x/semana", "leitura": "Viavel, porem com pouca folga e alta intensidade diaria.",
   "indicacao": "Plano de contingencia",
   "turmas": [
    ("DH 1",1,"Sala 1","Seg-Sex","2026-08-24","2026-09-04"),
    ("DH 2",1,"Sala 2","Seg-Sex","2026-08-24","2026-09-04"),
    ("DH 3",2,"Sala 1","Seg-Sex","2026-09-08","2026-09-21"),
    ("DH 4",2,"Sala 2","Seg-Sex","2026-09-08","2026-09-21"),
    ("DH 5",3,"Sala 1","Seg-Sex","2026-09-22","2026-10-05"),
    ("DH 6",3,"Sala 2","Seg-Sex","2026-09-22","2026-10-05"),
    ("DH 7",4,"Sala 1","Seg-Sex","2026-10-06","2026-10-20"),
    ("DH 8",4,"Sala 2","Seg-Sex","2026-10-06","2026-10-20"),
    ("DH 9",5,"Sala 1","Seg-Sex","2026-10-21","2026-11-04"),
    ("DH 10",5,"Sala 2","Seg-Sex","2026-10-21","2026-11-04"),
    ("DH 11",6,"Sala 1","Seg-Sex","2026-11-05","2026-11-18"),
    ("DH 12",6,"Sala 2","Seg-Sex","2026-11-05","2026-11-18"),
    ("DH 13",7,"Sala 1","Seg-Sex","2026-11-19","2026-12-03"),
    ("DH 14",7,"Sala 2","Seg-Sex","2026-11-19","2026-12-03"),
    ("DH 15",8,"Sala 1","Seg-Sex","2026-12-04","2026-12-17"),
   ]},
}

def parse_iso(s):
    y,m,d = map(int, s.split("-"))
    return date(y,m,d)

def turmas_ativas_pico_semanal(turmas):
    """Pico de turmas com o intervalo [inicio,fim] tocando a mesma semana corrida (nao e o mesmo
    que 'ao mesmo tempo na mesma noite')."""
    todas_datas = sorted(set([t["inicio"] for t in turmas] + [t["fim"] for t in turmas]))
    pico = 0
    d = min(t["inicio"] for t in turmas)
    fim_total = max(t["fim"] for t in turmas)
    while d <= fim_total:
        semana_ini = d - timedelta(days=d.weekday())
        semana_fim = semana_ini + timedelta(days=6)
        ativas = sum(1 for t in turmas if t["inicio"] <= semana_fim and t["fim"] >= semana_ini)
        pico = max(pico, ativas)
        d += timedelta(days=7)
    return pico

def max_turmas_mesma_noite(turmas):
    """Maximo real de turmas se encontrando na MESMA data e MESMO horario (checagem dia a dia
    pela cadencia real de cada turma) -- deve ser <= numero de salas por construcao."""
    d = min(t["inicio"] for t in turmas)
    fim_total = max(t["fim"] for t in turmas)
    pico = 0
    while d <= fim_total:
        n = sum(1 for t in turmas if t["inicio"] <= d <= t["fim"] and d.weekday() in cadence_weekdays(t["cadencia"]))
        pico = max(pico, n)
        d += timedelta(days=1)
    return pico

def build_lot_gaps(turmas):
    """Para cada lote (a partir do 2o), calcula quantos dias corridos de folga existem entre o
    termino do lote anterior e o inicio deste lote -- ou seja, quanto atraso o lote anterior
    poderia absorver sem empurrar este lote."""
    lots = sorted(set(t["lote"] for t in turmas))
    lot_start = {L: min(t["inicio"] for t in turmas if t["lote"] == L) for L in lots}
    lot_end = {L: max(t["fim"] for t in turmas if t["lote"] == L) for L in lots}
    gaps = {}
    for i in range(1, len(lots)):
        L, Lprev = lots[i], lots[i-1]
        gaps[L] = (lot_start[L] - lot_end[Lprev]).days
    return lots, lot_start, lot_end, gaps

def build_scenario(key):
    raw = SCENARIOS_RAW[key]
    turmas = []
    for (nome, lote, sala, cad, ini, fim) in raw["turmas"]:
        turmas.append({
            "turma": nome, "lote": lote, "sala": sala, "cadencia": cad,
            "inicio": parse_iso(ini), "fim": parse_iso(fim),
        })
    inicio_fase = min(t["inicio"] for t in turmas)
    fim_fase = max(t["fim"] for t in turmas)
    quarta_concl = sorted(t["fim"] for t in turmas)[3]
    duracao = (fim_fase - inicio_fase).days + 1
    turmas_semana_pico = turmas_ativas_pico_semanal(turmas)
    max_noite = max_turmas_mesma_noite(turmas)
    lots, lot_start, lot_end, lot_gaps = build_lot_gaps(turmas)
    return {
        "key": key, "salas": raw["salas"], "freq": raw["freq"],
        "turmas": turmas, "inicio_fase": inicio_fase, "fim_fase": fim_fase,
        "quarta_concl": quarta_concl, "duracao_dias": duracao,
        "turmas_semana_pico": turmas_semana_pico, "max_turmas_mesma_noite": max_noite,
        "lots": lots, "lot_start": lot_start, "lot_end": lot_end, "lot_gaps": lot_gaps,
        "leitura": raw["leitura"], "indicacao": raw["indicacao"],
    }

SCENARIOS = {k: build_scenario(k) for k in SCENARIOS_RAW}
RECOMMENDED = "B"

def simulate_lot_delay(scenario, lot_to_delay=1, delay_days=7, absorb_gaps=True):
    """Simula o atraso de `delay_days` dias no lote `lot_to_delay` (ex.: quorum insuficiente
    das Palestras que alimentam o lote) e propaga o efeito em cascata aos lotes seguintes.

    absorb_gaps=True  (cenario otimista): trata os intervalos naturais entre lotes (finais de
        semana/feriados que ja existiam no cronograma original) como folga que absorve parte
        do atraso antes de empurrar o lote seguinte.
    absorb_gaps=False (cenario conservador/pior caso): nenhuma folga e considerada -- o atraso
        se propaga integralmente, dia a dia, por todos os lotes seguintes.
    Ambos os casos devem ser lidos em conjunto -- a realidade fica entre os dois."""
    lots = scenario["lots"]
    lot_end = dict(scenario["lot_end"])
    lot_start = dict(scenario["lot_start"])
    atraso_acumulado = {}
    for i, L in enumerate(lots):
        if L < lot_to_delay:
            atraso_acumulado[L] = 0
            continue
        if L == lot_to_delay:
            atraso_acumulado[L] = delay_days
            continue
        gap = scenario["lot_gaps"].get(L, 0) if absorb_gaps else 0
        atraso_chegando = atraso_acumulado[lots[i-1]]
        atraso_acumulado[L] = max(0, atraso_chegando - gap)

    shifted = {L: {"start": lot_start[L] + timedelta(days=atraso_acumulado[L]),
                   "end": lot_end[L] + timedelta(days=atraso_acumulado[L]),
                   "atraso_aplicado": atraso_acumulado[L]} for L in lots}

    turmas_novas_fim = []
    for t in scenario["turmas"]:
        atraso_t = atraso_acumulado[t["lote"]]
        turmas_novas_fim.append(t["fim"] + timedelta(days=atraso_t))
    novo_fim_fase = max(turmas_novas_fim)
    novo_quarta_concl = sorted(turmas_novas_fim)[3]

    return {
        "lot_atrasado": lot_to_delay, "dias_atraso": delay_days, "lotes_shift": shifted,
        "fim_fase_original": scenario["fim_fase"], "fim_fase_novo": novo_fim_fase,
        "atraso_liquido_fim_fase": (novo_fim_fase - scenario["fim_fase"]).days,
        "quarta_concl_original": scenario["quarta_concl"], "quarta_concl_novo": novo_quarta_concl,
        "atraso_liquido_quarta": (novo_quarta_concl - scenario["quarta_concl"]).days,
    }

# ---------------------------------------------------------------------------
# 3. OFICINA DE GESTAO (OG) - NAO HA CONSENSO SOBRE O ESPACO FISICO A USAR.
#    Duas variantes sao modeladas explicitamente (nenhuma delas e "o cenario B" sozinha):
#      B1 = 3 salas de DH + 3 espacos ADICIONAIS dedicados a OG (6 espacos no total,
#           NAO confirmados) -- permite iniciar a OG uma semana apos cada DH concluir.
#      B2 = 3 espacos TOTAIS, compartilhados entre DH e OG (nenhum espaco adicional) --
#           como as 3 salas de DH ficam ocupadas por lotes consecutivos sem vaga livre
#           durante toda a fase de DH, a OG so pode comecar a ocupar essas MESMAS salas
#           depois que a fase de DH termina totalmente.
# ---------------------------------------------------------------------------
def build_og_schedule_b1_dedicado(scenario):
    """B1: espacos adicionais dedicados a OG (NAO confirmados) -- OG comeca 1 semana apos
    cada turma de DH concluir, sem disputar sala com o DH."""
    ogs = []
    n_espacos = scenario["salas"]
    for i, t in enumerate(scenario["turmas"]):
        start = t["fim"] + timedelta(days=7)
        while start.weekday() > 4:
            start += timedelta(days=1)
        dates = add_working_slots(start, 2, {0,1,2,3,4})
        og_sala = f"Espaço OG {(i % n_espacos) + 1} (ADICIONAL, não confirmado)"
        ogs.append({
            "turma": t["turma"].replace("DH","OG"), "dh_origem": t["turma"], "sala": og_sala,
            "inicio": dates[0], "fim": dates[-1], "encontros": dates,
        })
    return ogs

def build_og_schedule_b2_compartilhado(scenario):
    """B2: OG usa as MESMAS 3 salas do DH (sem espaco adicional). Como essas salas nao tem
    vaga livre durante a fase de DH (ver achado de auditoria), a OG so comeca a ocupar as
    salas apos o TERMINO TOTAL da fase de DH (fim_fase), em ondas sequenciais de `salas`
    turmas por vez."""
    ogs = []
    salas_nomes = sorted(set(t["sala"] for t in scenario["turmas"]))
    n_salas = len(salas_nomes)
    cursor = scenario["fim_fase"] + timedelta(days=1)
    while cursor.weekday() > 4:
        cursor += timedelta(days=1)
    turmas_ordenadas = sorted(scenario["turmas"], key=lambda t: (t["lote"], t["turma"]))
    onda_start = cursor
    for i, t in enumerate(turmas_ordenadas):
        sala_idx = i % n_salas
        if sala_idx == 0 and i > 0:
            # nova onda: comeca no dia util seguinte ao fim da onda anterior
            onda_start = max(o["fim"] for o in ogs[-n_salas:]) + timedelta(days=1)
            while onda_start.weekday() > 4:
                onda_start += timedelta(days=1)
        dates = add_working_slots(onda_start, 2, {0,1,2,3,4})
        ogs.append({
            "turma": t["turma"].replace("DH","OG"), "dh_origem": t["turma"], "sala": salas_nomes[sala_idx],
            "inicio": dates[0], "fim": dates[-1], "encontros": dates,
        })
    return ogs

# Alias mantido para compatibilidade -- aponta para a variante B1 (a mais rapida, mas com
# espacos NAO confirmados). Todo uso deste alias deve reforcar, no texto do relatorio, que
# se trata de uma premissa provisoria sujeita a validacao -- ver 07_CENARIO_B (B1 x B2).
build_og_schedule = build_og_schedule_b1_dedicado

# ---------------------------------------------------------------------------
# 4. CURSOS TECNICOS (CT) - base operacional (10 turmas, 7 temas) x 2 trilhas paralelas
#    Gate: so pode iniciar apos conclusao da 4a turma de DH (+ buffer de 1 semana)
# ---------------------------------------------------------------------------
CT_CATALOG = [
    # tema, sigla, encontros, carga_horaria, min, max, fornecedor, trilha
    ("Corte e Costura",       "CT1a", 8,  32, 15, 18, "Fundo Social / SENAI",  "T1"),
    ("Corte e Costura",       "CT1b", 8,  32, 15, 18, "Fundo Social / SENAI",  "T1"),
    ("Beleza / Unhas",        "CT3",  10, 40, 15, 18, "Fundo Social / SENAC",  "T1"),
    ("Beleza / Cabelos Cacheados e Crespos", "CT5", 14, 56, 15, 18, "Fundo Social / SENAC", "T1"),
    ("Preparacao de Alimentos","CT4a",14, 56, 15, 18, "Fundo Social / SENAC",  "T1"),
    ("Preparacao de Alimentos","CT4b",14, 56, 15, 18, "Fundo Social / SENAC",  "T1"),
    ("Eletrica",              "CT7a", 20, 80, 14, 16, "SENAI Paulinia",        "T2"),
    ("Eletrica",              "CT7b", 20, 80, 14, 16, "SENAI Paulinia",        "T2"),
    ("Automacao",             "CT6",  20, 80, 14, 16, "SENAI Paulinia",        "T2"),
    ("Jardinagem",            "CT2",  10, 40, 15, 18, "? (fornecedor nao confirmado)", "T3"),
]
CT_WEEKDAYS = {0,1,2,3,4}  # 5x/semana, 4h/dia (noite), conforme premissa

MEDIA_PARTICIPANTES_DH = 30  # media do intervalo 20-40 informado na planilha operacional (premissa)

def build_ct_schedule(scenario):
    """
    IMPORTANTE: todas as datas retornadas aqui sao "datas propostas para simulacao" --
    dependem de quorum real, contratacao de fornecedor e validacao, nenhuma delas e um
    compromisso oficial (ver aba 14_CURSOS_TECNICOS e aba 01_PREMISSAS).
    """
    gate = scenario["quarta_concl"] + timedelta(days=7)
    while gate.weekday() > 4:
        gate += timedelta(days=1)
    track_cursor = {"T1": gate, "T2": gate, "T3": gate}
    dh_fins_ordenados = sorted(t["fim"] for t in scenario["turmas"])
    cts = []
    for tema, sigla, n_enc, ch, mn, mx, fornecedor, trilha in CT_CATALOG:
        start = track_cursor[trilha]
        dates = add_working_slots(start, n_enc, CT_WEEKDAYS)
        dh_concluidas_antes = sum(1 for f in dh_fins_ordenados if f <= start)
        participantes_estimados = min(dh_concluidas_antes * MEDIA_PARTICIPANTES_DH, len(scenario["turmas"]) * MEDIA_PARTICIPANTES_DH)
        fornecedor_confirmado = "?" not in fornecedor
        local_confirmado = fornecedor_confirmado  # mesma fonte de incerteza nesta planilha
        cts.append({
            "sigla": sigla, "tema": tema, "encontros": n_enc, "carga_horaria": ch,
            "min": mn, "max": mx, "fornecedor": fornecedor, "trilha": trilha,
            "inicio": dates[0], "fim": dates[-1],
            "dh_concluidas_antes": dh_concluidas_antes, "participantes_estimados": participantes_estimados,
            "fornecedor_confirmado": fornecedor_confirmado, "local_confirmado": local_confirmado,
            "data_validada": False,
            "responsavel_validacao": "Coordenação de parcerias / fornecedor",
            "classificacao": "Data proposta para simulação — pendente de quórum, contratação e validação do fornecedor.",
        })
        track_cursor[trilha] = dates[-1] + timedelta(days=3)  # pequeno intervalo entre turmas da mesma trilha
    return cts, gate

# ---------------------------------------------------------------------------
# 5. EAD - inicia junto com o 1o CT de cada trilha (auto-instrucional), 3 encontros/2h
# ---------------------------------------------------------------------------
def build_ead_schedule(ct_list):
    first_ct_end = min(c["fim"] for c in ct_list)
    start = first_ct_end + timedelta(days=1)
    while start.weekday() > 4:
        start += timedelta(days=1)
    dates = add_working_slots(start, 3, {0,1,2,3,4})
    return {"inicio": dates[0], "fim": dates[-1], "carga_horaria": 6, "encontros": 3}

# ---------------------------------------------------------------------------
# 6. CONSULTORIA INDIVIDUAL (CI) - rolling, apos CT concluido de cada aluno; meta 150 pessoas
# ---------------------------------------------------------------------------
def build_ci_schedule(ct_list):
    start = min(c["fim"] for c in ct_list) + timedelta(days=7)
    while start.weekday() > 4:
        start += timedelta(days=1)
    return {"inicio": start, "meta_pessoas": 150, "atendimentos_por_pessoa": 3, "horas_por_atendimento": 1}

def compute_ci_capacity(inicio, meta_pessoas=150, atendimentos_por_pessoa=3, horas_por_atendimento=1,
                         n_consultores=2, horas_dia=4, dias_semana=5, atend_simultaneos=2):
    """Calcula se a meta de atendimentos de Consultoria Individual cabe em um horizonte, dado um
    conjunto de parametros editaveis de capacidade. Retorna semanas necessarias e a data de
    conclusao estimada (primeira data viavel)."""
    total_atendimentos = meta_pessoas * atendimentos_por_pessoa
    total_horas_necessarias = total_atendimentos * horas_por_atendimento
    capacidade_paralela = min(n_consultores, atend_simultaneos)
    atendimentos_por_semana = capacidade_paralela * horas_dia * dias_semana / horas_por_atendimento
    if atendimentos_por_semana <= 0:
        return {"semanas_necessarias": None, "data_fim_estimada": None, "total_atendimentos": total_atendimentos,
                "atendimentos_por_semana": 0}
    import math
    semanas = math.ceil(total_atendimentos / atendimentos_por_semana)
    data_fim = inicio + timedelta(weeks=semanas)
    return {
        "semanas_necessarias": semanas, "data_fim_estimada": data_fim,
        "total_atendimentos": total_atendimentos, "total_horas_necessarias": total_horas_necessarias,
        "atendimentos_por_semana": atendimentos_por_semana, "capacidade_paralela": capacidade_paralela,
    }

if __name__ == "__main__":
    print("PS: primeira", PS_SCHEDULE[0]["data"], "ultima", PS_SCHEDULE[-1]["data"], "total", len(PS_SCHEDULE))
    for k, s in SCENARIOS.items():
        print(f"\nCenario {k}: inicio={s['inicio_fase']} quarta={s['quarta_concl']} fim={s['fim_fase']} dur={s['duracao_dias']}d "
              f"semana_pico={s['turmas_semana_pico']} max_noite={s['max_turmas_mesma_noite']}")
    rec = SCENARIOS[RECOMMENDED]
    og1 = build_og_schedule_b1_dedicado(rec)
    og2 = build_og_schedule_b2_compartilhado(rec)
    print("\nOG B1 (dedicado) - primeiras 3:", [(o["turma"], o["sala"], o["inicio"], o["fim"]) for o in og1[:3]])
    print("OG B2 (compartilhado) - inicio/fim geral:", min(o["inicio"] for o in og2), max(o["fim"] for o in og2))
    ct, gate = build_ct_schedule(rec)
    print(f"\nGate CT (cenario {RECOMMENDED}): {gate}")
    for c in ct:
        print(" ", c["sigla"], c["tema"], c["inicio"], "-", c["fim"], c["fornecedor"], "DHconcl:", c["dh_concluidas_antes"], "part.est:", c["participantes_estimados"])
    ead = build_ead_schedule(ct)
    print("\nEAD:", ead)
    ci = build_ci_schedule(ct)
    cap = compute_ci_capacity(ci["inicio"])
    print("CI:", ci, "capacidade:", cap)
    ci = build_ci_schedule(ct)
    print("CI:", ci)
