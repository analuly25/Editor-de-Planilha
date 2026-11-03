# -*- coding: utf-8 -*-
"""Planilha Engine - arruma automaticamente a planilha recebida.

Uso:  PlanilhaEngine.exe <perfil.json> <arquivo_entrada> <pasta_saida>
"""
import csv
import json
import os
import sys
import unicodedata

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as letra


# Formato Contábil em reais (R$), com separador de milhares e 2 casas.
# (O Excel em português mostra com ponto de milhar e vírgula decimal.)
FORMATO_REAIS = '_("R$"* #,##0.00_);_("R$"* \\(#,##0.00\\);_("R$"* "-"??_);_(@_)'


def norm(s):
    """Normaliza nome de coluna: sem acento, maiúsculo, sem espaços."""
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.strip().upper()


def ler_entrada(caminho):
    ext = os.path.splitext(caminho)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        ws = load_workbook(caminho, read_only=True, data_only=True).worksheets[0]
        return [["" if v is None else v for v in row] for row in ws.iter_rows(values_only=True)]
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(caminho, encoding=enc, newline="") as f:
                amostra = f.read(4096)
                f.seek(0)
                delim = ";" if amostra.count(";") >= amostra.count(",") else ","
                return [r for r in csv.reader(f, delimiter=delim)]
        except UnicodeDecodeError:
            continue
    raise ValueError("Não consegui ler o arquivo (codificação desconhecida).")


def para_numero(v):
    """'1.234,56' / '1234.56' / 12 -> float; texto não numérico fica como está."""
    if isinstance(v, (int, float)):
        return v
    t = str(v).strip().replace("R$", "").strip()
    if not t:
        return ""
    try:
        if "," in t:
            t = t.replace(".", "").replace(",", ".")
        return float(t)
    except ValueError:
        return v


def main():
    if len(sys.argv) < 4:
        print("[ERRO] Uso: PlanilhaEngine.exe <perfil.json> <arquivo_entrada> <pasta_saida>")
        return 2
    perfil_path, entrada, saida_dir = sys.argv[1:4]

    with open(perfil_path, encoding="utf-8") as f:
        p = json.load(f)

    dados = ler_entrada(entrada)
    if not dados:
        print("[ERRO] O arquivo está vazio.")
        return 1
    cab = [str(c).strip() for c in dados[0]]
    linhas = [list(r) + [""] * (len(cab) - len(r)) for r in dados[1:]]
    # descarta linhas totalmente vazias
    linhas = [r for r in linhas if any(str(c).strip() for c in r)]
    idx = {}
    for i, c in enumerate(cab):
        idx.setdefault(norm(c), i)

    vb = p["valor_bruto"]
    necessarias = (p["colunas_visiveis"] + p["colunas_numericas"]
                   + [p["coluna_duplicados"], vb["coluna_base"], vb["coluna_uf"]])
    faltando = sorted({c for c in necessarias if norm(c) not in idx})
    if faltando:
        print("[ERRO] Colunas não encontradas na planilha: " + ", ".join(faltando))
        print("       Confira se este é o arquivo certo (cabeçalho na primeira linha).")
        return 1

    # --- monta a tabela (colunas originais, sem inserir nada) ---
    num_idx = {idx[norm(c)] for c in p["colunas_numericas"]}
    i_preco = idx[norm("PRECO_VIGENTE_DO_CIRCUITO")]

    def trata(i, v):
        if i in num_idx:
            v = para_numero(v)
        if i == i_preco and isinstance(v, float):
            v = round(v, 2)
        return v

    tabela = [[trata(i, v) for i, v in enumerate(r)] for r in linhas]

    c_acesso = idx[norm(p["coluna_duplicados"])] + 1           # 1-based
    c_base = idx[norm(vb["coluna_base"])] + 1
    c_uf = idx[norm(vb["coluna_uf"])] + 1
    c_preco = idx[norm("PRECO_VIGENTE_DO_CIRCUITO")] + 1
    c_bruto = len(cab) + 1
    cab_novo = cab + [vb["nome"]]

    wb = Workbook()
    ws = wb.active
    ws.title = "Planilha"
    ws.append(cab_novo)
    for r in tabela:
        ws.append(r + [None])
    n = len(tabela) + 1

    # --- aba de alíquotas (consultada pela fórmula) ---
    wa = wb.create_sheet("ALIQUOTAS")
    wa.append(["UF", "ICMS"])
    for uf, a in p["aliquotas_icms"].items():
        wa.append([uf, a])
    for row in wa.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "0.0%"
    wa["D1"] = "PIS/COFINS"
    wa["E1"] = vb["pis_cofins"]
    wa["E1"].number_format = "0.00%"
    for c in ("A1", "B1", "D1"):
        wa[c].font = Font(bold=True)

    # --- fórmula do VALOR_BRUTO ---
    L_base, L_uf, L_br = letra(c_base), letra(c_uf), letra(c_bruto)
    for r in range(2, n + 1):
        ws.cell(r, c_bruto).value = (
            f'=IF({L_base}{r}="","",IFERROR(ROUND({L_base}{r}/((1-VLOOKUP({L_uf}{r},ALIQUOTAS!$A:$B,2,FALSE))'
            f'*(1-ALIQUOTAS!$E$1)),2),""))'
        )
        ws.cell(r, c_bruto).number_format = FORMATO_REAIS
        ws.cell(r, c_base).number_format = FORMATO_REAIS

    # preço vigente com formato numérico
    for r in range(2, n + 1):
        ws.cell(r, c_preco).number_format = FORMATO_REAIS
        # ACESSO em formato Geral: permite digitar fórmulas à mão nas células
        ws.cell(r, c_acesso).number_format = "General"

    # --- realce vermelho de duplicados em ACESSO ---
    L_ac = letra(c_acesso)
    if n >= 2:
        ws.conditional_formatting.add(
            f"{L_ac}2:{L_ac}{n}",
            FormulaRule(
                formula=[f'AND({L_ac}2<>"",COUNTIF(${L_ac}$2:${L_ac}${n},{L_ac}2)>1)'],
                fill=PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"),
                font=Font(color="FFFFFF", bold=True),
            ),
        )

    # --- visibilidade, formatação geral ---
    visiveis = {norm(c) for c in p["colunas_visiveis"]}
    visiveis_idx = {idx[v] + 1 for v in visiveis} | {c_bruto}
    for c in range(1, len(cab_novo) + 1):
        ws.column_dimensions[letra(c)].hidden = c not in visiveis_idx
    for c in visiveis_idx:
        ws.column_dimensions[letra(c)].width = 24
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{letra(len(cab_novo))}{n}"
    wb.active = 0

    os.makedirs(saida_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(entrada))[0]
    destino = os.path.join(saida_dir, f"{base}_EDITADA.xlsx")
    wb.save(destino)
    print(f"[OK] {len(tabela)} linha(s) processada(s).")
    print(f"[OK] Arquivo salvo em: {destino}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except PermissionError:
        print("[ERRO] Não consegui salvar. Feche a planilha se ela estiver aberta no Excel.")
        sys.exit(1)
    except Exception as e:
        print(f"[ERRO] {type(e).__name__}: {e}")
        sys.exit(1)
