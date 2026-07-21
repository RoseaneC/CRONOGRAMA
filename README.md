# Cronograma Paulínia — Partiu! Aprender e Empreender

Cronograma completo (Instituto da Criança · Sebrae · Petrobras) para o Núcleo 2 (Paulínia).

## Arquivos

- `cronograma/Cronograma_Paulinia_Partiu_2026.xlsx` — cronograma completo, 18 abas (premissas, divergências,
  4 cenários de Desenvolvimento Humano, cronograma mestre, agenda por sala, turmas, riscos e Gantt executivo).
- `executivo/Visao_Executiva_Paulinia.html` — visão executiva de uma página (mesma versão publicada como Artifact).
- `scripts/model.py` — motor de datas (Palestras, 4 cenários de DH, Oficina de Gestão, Cursos Técnicos, EAD, Consultoria).
- `scripts/build_workbook.py` — gera o arquivo `.xlsx` a partir do `model.py`.

## Como regenerar o Excel após mudar uma premissa estrutural

As premissas editáveis (carga horária, nº de salas, frequência etc.) ficam na aba `01_PREMISSAS` do próprio Excel,
mas mudanças estruturais (nº de salas, frequência do DH, regra de formação de turmas) exigem regenerar o arquivo,
pois o Excel nativo não recalcula um algoritmo de alocação de salas.

```bash
pip install openpyxl
python3 scripts/build_workbook.py
```

O arquivo é escrito em `cronograma/Cronograma_Paulinia_Partiu_2026.xlsx`.

## Fontes originais (não alteradas)

Este cronograma cruza 4 documentos fornecidos: a apresentação macro do projeto (metas contratuais), a planilha
operacional de turmas, uma tentativa anterior de cronograma da coordenação e uma simulação inicial de 4 cenários.
Nenhum dos 4 arquivos originais foi editado. Divergências entre eles estão documentadas na aba `02_DIVERGENCIAS`
do Excel e nunca foram resolvidas silenciosamente.
