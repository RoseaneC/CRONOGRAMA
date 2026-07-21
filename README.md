# Cronograma Paulínia — Partiu! Aprender e Empreender

Cronograma completo (Instituto da Criança · Sebrae · Petrobras) para o Núcleo 2 (Paulínia).

## Arquivos

- `cronograma/Cronograma_Paulinia_Partiu_2026_V2.xlsx` — **versão atual**, 20 abas. Revisão crítica da V1: automação
  esclarecida (o que recalcula por fórmula x o que exige regerar), capacidade de sala separada em "turmas ativas na
  semana" x "máximo por noite", duas variantes de espaço para a Oficina de Gestão (B1/B2, nenhuma confirmada),
  recomendação tratada como condicionada (não aprovada), marcos de validação de quórum e simulação de atraso de 7
  dias, Cursos Técnicos reclassificados como "proposta para simulação", calculadora de capacidade da Consultoria
  Individual, aba para parceiros e log completo de alterações (`19_LOG_DE_ALTERACOES`).
- `cronograma/Cronograma_Paulinia_Partiu_2026.xlsx` — versão V1 original, mantida sem alterações para histórico.
- `executivo/Visao_Executiva_Paulinia.html` — visão executiva de uma página (mesma versão publicada como Artifact; baseada na V1).
- `scripts/model.py` — motor de datas (Palestras, 4 cenários de DH, variantes B1/B2 de OG, Cursos Técnicos, EAD,
  simulação de atraso, calculadora de capacidade da Consultoria Individual). Reflete a lógica da V2.
- `scripts/build_workbook.py` — gera o arquivo `.xlsx` a partir do `model.py`. Gera atualmente a V2
  (`Cronograma_Paulinia_Partiu_2026_V2.xlsx`) — não sobrescreve a V1.

## Como regenerar o Excel após mudar uma premissa estrutural

As premissas editáveis ficam na aba `01_PREMISSAS`. Uma matriz completa em `00_LEIA-ME` mostra, premissa por
premissa, o que atualiza sozinho (fórmula) e o que exige rodar o script — em resumo: o "Cenário recomendado"
(dropdown A/B/C/D) e a calculadora de capacidade da Consultoria Individual recalculam sozinhos; qualquer mudança
estrutural (nº de salas, frequência do DH, carga horária, variante de OG) exige regenerar o arquivo, pois o Excel
nativo não recalcula um algoritmo de alocação de salas.

```bash
pip install openpyxl
python3 scripts/build_workbook.py
```

O arquivo é escrito em `cronograma/Cronograma_Paulinia_Partiu_2026_V2.xlsx`.

## Fontes originais (não alteradas)

Este cronograma cruza 4 documentos fornecidos: a apresentação macro do projeto (metas contratuais), a planilha
operacional de turmas, uma tentativa anterior de cronograma da coordenação e uma simulação inicial de 4 cenários.
Nenhum dos 4 arquivos originais foi editado. Divergências entre eles estão documentadas na aba `02_DIVERGENCIAS`
do Excel e nunca foram resolvidas silenciosamente.
