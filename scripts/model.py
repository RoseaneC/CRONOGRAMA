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
    cap_simultanea = raw["salas"] * (2 if raw["freq"]=="2x/semana" else 1)
    return {
        "key": key, "salas": raw["salas"], "freq": raw["freq"],
        "turmas": turmas, "inicio_fase": inicio_fase, "fim_fase": fim_fase,
        "quarta_concl": quarta_concl, "duracao_dias": duracao,
        "cap_simultanea": cap_simultanea, "leitura": raw["leitura"],
        "indicacao": raw["indicacao"],
    }

SCENARIOS = {k: build_scenario(k) for k in SCENARIOS_RAW}
RECOMMENDED = "B"

# ---------------------------------------------------------------------------
# 3. OFICINA DE GESTAO (OG) - premissa base: 1 OG por turma de DH concluida (15 total)
#    2 encontros de 2h, na semana seguinte ao termino do DH (mesma sala da turma)
# ---------------------------------------------------------------------------
def build_og_schedule(scenario):
    """
    ACHADO DE AUDITORIA: nos cenarios com DH 5x/semana (B e D), as salas de DH ficam
    100% ocupadas por lotes consecutivos de DH durante toda a fase (sem dia livre na
    mesma sala). Por isso a OG NAO pode reutilizar a sala do proprio DH nesses cenarios
    -- precisa de um espaco adicional dedicado. Esse achado e registrado nas abas
    02/03/16. Para cenarios 2x/semana (A e C) ha interleaving de dias e a sala de
    origem poderia, em tese, ser reaproveitada -- mas para manter uma unica premissa
    auditavel e simples de comunicar, o modelo usa sempre um espaco dedicado de OG.
    """
    ogs = []
    for i, t in enumerate(scenario["turmas"]):
        start = t["fim"] + timedelta(days=7)
        while start.weekday() > 4:
            start += timedelta(days=1)
        dates = add_working_slots(start, 2, {0,1,2,3,4})
        og_sala = f"Espaço OG {(i % 3) + 1} (dedicado, fora das 3 salas de DH)"
        ogs.append({
            "turma": t["turma"].replace("DH","OG"), "dh_origem": t["turma"], "sala": og_sala,
            "inicio": dates[0], "fim": dates[-1], "encontros": dates,
        })
    return ogs

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

def build_ct_schedule(scenario):
    gate = scenario["quarta_concl"] + timedelta(days=7)
    while gate.weekday() > 4:
        gate += timedelta(days=1)
    track_cursor = {"T1": gate, "T2": gate, "T3": gate}
    cts = []
    for tema, sigla, n_enc, ch, mn, mx, fornecedor, trilha in CT_CATALOG:
        start = track_cursor[trilha]
        dates = add_working_slots(start, n_enc, CT_WEEKDAYS)
        cts.append({
            "sigla": sigla, "tema": tema, "encontros": n_enc, "carga_horaria": ch,
            "min": mn, "max": mx, "fornecedor": fornecedor, "trilha": trilha,
            "inicio": dates[0], "fim": dates[-1],
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
    start = max(c["fim"] for c in ct_list) - timedelta(days=60)  # comeca a fluir conforme CTs vao terminando
    start = min(c["fim"] for c in ct_list) + timedelta(days=7)
    while start.weekday() > 4:
        start += timedelta(days=1)
    end = max(c["fim"] for c in ct_list) + timedelta(days=30)
    return {"inicio": start, "fim": end, "meta_pessoas": 150, "atendimentos_por_pessoa": 3}

if __name__ == "__main__":
    print("PS: primeira", PS_SCHEDULE[0]["data"], "ultima", PS_SCHEDULE[-1]["data"], "total", len(PS_SCHEDULE))
    for k, s in SCENARIOS.items():
        print(f"\nCenario {k}: inicio={s['inicio_fase']} quarta={s['quarta_concl']} fim={s['fim_fase']} dur={s['duracao_dias']}d cap={s['cap_simultanea']}")
    rec = SCENARIOS[RECOMMENDED]
    og = build_og_schedule(rec)
    print("\nOG (cenario recomendado) - primeiras 3:")
    for o in og[:3]:
        print(" ", o["turma"], o["inicio"], "-", o["fim"], "sala", o["sala"])
    ct, gate = build_ct_schedule(rec)
    print(f"\nGate CT (cenario {RECOMMENDED}): {gate}")
    for c in ct:
        print(" ", c["sigla"], c["tema"], c["inicio"], "-", c["fim"], c["fornecedor"])
    ead = build_ead_schedule(ct)
    print("\nEAD:", ead)
    ci = build_ci_schedule(ct)
    print("CI:", ci)
