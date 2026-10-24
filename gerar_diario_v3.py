#!/usr/bin/env python3
"""
GERAR DASHBOARD DE RUPTURA x AGENDA — CD SP
"""

import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.chart import BarChart, PieChart, Reference, LineChart
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import DataPoint
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, column_index_from_string

SCRIPT_DIR = Path(__file__).resolve().parent

# Cobertura (em dias) a partir da qual um item com carga agendada é
# considerado "giro baixo" — ou seja, já tem estoque suficiente e a compra
# adicional agendada é questionável. Ajustável conforme a política de
# estoque do CD (ver Notas_Metodológicas, item 8).
LIMIAR_DIAS_GIRO_BAIXO = 45


FONT_NAME = "Arial"
NAVY = "1F2937"
HEADER_FILL = "1F2937"
HEADER_FONT_COLOR = "FFFFFF"
RED_FILL = "F5B7B1"
ORANGE_FILL = "FDEBD0"
GREY_FILL = "F2F3F4"

# Paleta de cores oficiais do Grupo DPSP (extraídas da logo):
DPSP_CINZA = "ADADAD"
DPSP_VERMELHO = "F21111"
DPSP_AZUL = "63D0F8"

# Mapeamento semântico dos gráficos usando SÓ as 3 cores oficiais + 2 tons
# derivados (mais claro/escuro) apenas para manter distinção visual entre
# "crítico" e "atenção" (a marca não definiu uma 4ª cor para isso).
COR_CRITICO = DPSP_VERMELHO       # vermelho oficial — pior situação / crítico
COR_ATENCAO = "F27575"            # tom mais claro do vermelho — atenção/intermediário
COR_OK = DPSP_AZUL                # azul claro oficial — resolvido / sem risco
COR_NEUTRO = DPSP_CINZA           # cinza oficial — série principal / neutra
COR_SECUNDARIA = "6E6E6E"         # tom mais escuro do cinza — série de comparação

# Cores específicas do gráfico "Curva de ruptura projetada" (ajuste manual
# validado por Helber em 24/08/2026, fora da paleta oficial DPSP):
COR_CURVA_COM_AGENDA_LINHA = "E44438"   # linha "% Zerados COM Agenda"
COR_CURVA_SEM_AGENDA_BARRA = "898989"   # barra "Qtde Zerados SEM Agenda"
COR_CURVA_COM_AGENDA_BARRA = "00B0F0"   # barra "Qtde Zerados COM Agenda"


def _formatar_dias_exibicao(dias):
    """Formata dias para exibição: valores < 0.5 viram '0 dias' (mais legível
    que 0,002), outros são mantidos com até 2 casas decimais."""
    if pd.isna(dias) or dias is None:
        return np.nan
    dias_float = float(dias)
    if dias_float < 0.5:
        return 0  # Será exibido como "0" no gráfico, mais limpo que "0,002"
    return round(dias_float, 2)


def _rotulo_cod_desc(codigo, descricao, max_len=32):
    """Combina código + descrição num único rótulo legível para uso em
    categorias de gráfico (ex.: '877417 — VIT B12 DSP 60 CP MAST')."""
    desc = str(descricao) if descricao not in (None, "") else ""
    if len(desc) > max_len:
        desc = desc[: max_len - 1].rstrip() + "…"
    cod = str(int(codigo)) if isinstance(codigo, float) else str(codigo)
    return f"{cod} — {desc}" if desc else cod


def _sem_grade_vertical(chart):
    """Remove a grade do eixo de categorias (deixa só linhas horizontais discretas)."""
    if hasattr(chart, "x_axis"):
        chart.x_axis.majorGridlines = None
        chart.x_axis.delete = False
    if hasattr(chart, "y_axis"):
        chart.y_axis.delete = False


def _rotular_dados(chart, formato=None):
    """Ativa rótulos de dados (valores) diretamente nas barras/colunas/linhas."""
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.showSerName = False
    chart.dataLabels.showCatName = False
    chart.dataLabels.showLegendKey = False
    if formato:
        chart.dataLabels.numFmt = formato
    return chart


def _colorir_series(chart, cores):
    """Aplica uma cor sólida a cada série do gráfico, na ordem informada."""
    for serie, cor in zip(chart.series, cores):
        serie.graphicalProperties.solidFill = cor
        serie.graphicalProperties.line.noFill = True


def _colorir_fatias_pizza(pie, cores):
    """Aplica uma cor por fatia (data point) em um gráfico de pizza."""
    pontos = []
    for i, cor in enumerate(cores):
        dp = DataPoint(idx=i)
        dp.graphicalProperties.solidFill = cor
        dp.graphicalProperties.line.solidFill = "FFFFFF"
        pontos.append(dp)
    pie.series[0].data_points = pontos


def _escolher_arquivo_gui(titulo, pasta_inicial):
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        caminho = filedialog.askopenfilename(
            title=titulo,
            initialdir=pasta_inicial,
            filetypes=[
                ("Arquivos Excel", "*.xlsx *.xls"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        root.destroy()
        return caminho or None
    except Exception:
        return None


def _ler_excel_robusto(path, rotulo, **kwargs):
    tentativas = ["openpyxl", "xlrd", "calamine"]
    ultimo_erro = None
    for engine in tentativas:
        try:
            return pd.read_excel(path, engine=engine, **kwargs)
        except ImportError:
            continue
        except Exception as e:
            ultimo_erro = e
            continue
    print(f"\nERRO: não consegui abrir o arquivo de {rotulo}:\n  {path}")
    print(f"(detalhe técnico: {ultimo_erro})")
    sys.exit(1)


def _cobertura_legivel(dias, ja_zerado):
    if ja_zerado:
        return "Já zerado"
    if not np.isfinite(dias):
        return "Sem risco em 60 dias"
    if dias >= 1:
        if dias < 30:
            return f"{dias:.1f} dias"
        return f"{dias:.0f} dias"
    horas = dias * 24
    if horas < 1:
        return f"{horas * 60:.0f} minutos"
    return f"{horas:.1f} horas"


def _estilizar_tabela(ws, moeda_cols=None, data_cols=None):
    moeda_cols = moeda_cols or []
    data_cols = data_cols or []
    max_col, max_row = ws.max_column, ws.max_row
    header = [ws.cell(row=1, column=c).value for c in range(1, max_col + 1)]
    for c in range(1, max_col + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(name=FONT_NAME, bold=True, color=HEADER_FONT_COLOR, size=10)
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        ws.column_dimensions[get_column_letter(c)].width = max(
            14, min(38, len(str(header[c - 1])) + 4)
        )
    thin = Side(style="thin", color="D5D8DC")
    for r in range(2, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name=FONT_NAME, size=10)
            cell.border = Border(bottom=thin)
            col_name = header[c - 1]
            if col_name in moeda_cols:
                cell.number_format = '"R$" #,##0.00'
            elif col_name in data_cols:
                cell.number_format = "DD/MM/YYYY"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    return header


def main():
    if len(sys.argv) > 1:
        agenda_path = sys.argv[1]
        ruptura_path = (
            sys.argv[2]
            if len(sys.argv) > 2
            else str(SCRIPT_DIR / "Ruptura_qlikview.xlsx")
        )
        out_path = (
            sys.argv[3]
            if len(sys.argv) > 3
            else str(SCRIPT_DIR / "Dashboard_Ruptura.xlsx")
        )
        norma_path = sys.argv[4] if len(sys.argv) > 4 else None
        if norma_path is None:
            candidatos_norma = sorted(SCRIPT_DIR.glob("420_NORMA*.csv"))
            norma_path = str(candidatos_norma[0]) if candidatos_norma else None
    else:
        agenda_default = SCRIPT_DIR / "Agenda_CD_SP.xlsx"
        ruptura_default = SCRIPT_DIR / "Ruptura_qlikview.xlsx"
        if agenda_default.exists() and ruptura_default.exists():
            agenda_path, ruptura_path = str(agenda_default), str(ruptura_default)
        else:
            print("Não encontrei os arquivos com os nomes padrão nesta pasta.")
            agenda_path = _escolher_arquivo_gui(
                "Selecione o arquivo Agenda_CD_SP.xlsx", str(SCRIPT_DIR)
            )
            if not agenda_path:
                sys.exit(1)
            ruptura_path = _escolher_arquivo_gui(
                "Selecione o arquivo Ruptura_qlikview.xlsx",
                str(Path(agenda_path).parent),
            )
            if not ruptura_path:
                sys.exit(1)
        out_path = str(
            SCRIPT_DIR
            / f"Dashboard_Ruptura_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        candidatos_norma = sorted(SCRIPT_DIR.glob("420_NORMA*.csv"))
        norma_path = str(candidatos_norma[0]) if candidatos_norma else None

    agenda = _ler_excel_robusto(agenda_path, "Agenda", sheet_name="Agenda")
    agenda.columns = [c.strip() for c in agenda.columns]
    agenda = agenda.rename(
        columns={
            "Quantidade total de SKUs": "Qtd_Agendada",
            "Nome completo": "Fornecedor_Nome_Agenda",
            "Dt.início desejada": "Data_Prevista",
            "Denominação do status da data de carga": "Status_Carga",
        }
    )
    agenda["Material"] = pd.to_numeric(agenda["Material"], errors="coerce")
    agenda = agenda.dropna(subset=["Material"])
    agenda["Material"] = agenda["Material"].astype(np.int64)
    agenda_agg = (
        agenda.groupby("Material")
        .agg(
            Qtd_Agendada_Total=("Qtd_Agendada", "sum"),
            Data_Prevista_Mais_Proxima=("Data_Prevista", "min"),
            Fornecedores_Agenda=(
                "Fornecedor_Nome_Agenda",
                lambda s: "; ".join(sorted(set(str(x) for x in s))),
            ),
            Status_Cargas=(
                "Status_Carga",
                lambda s: "; ".join(sorted(set(str(x) for x in s))),
            ),
        )
        .reset_index()
    )

    ruptura_raw = _ler_excel_robusto(ruptura_path, "Ruptura", sheet_name=0)
    ruptura_raw.columns = [c.strip() for c in ruptura_raw.columns]
    totais_row = ruptura_raw[ruptura_raw["Cód. Produto"].isna()]
    pct_ruptura_geral_oficial = (
        totais_row.iloc[0]["%Ruptura"] if not totais_row.empty else None
    )
    valor_ruptura_total_oficial = (
        totais_row.iloc[0]["Valor C/Ruptura"] if not totais_row.empty else None
    )

    ruptura = ruptura_raw[ruptura_raw["Cód. Produto"].notna()].copy()
    ruptura["Cód. Produto"] = ruptura["Cód. Produto"].astype(np.int64)

    # --- Diagnóstico: código de produto duplicado no arquivo de Ruptura ---
    # Isto NÃO altera nenhum dado — apenas expõe o problema para decisão humana.
    # Se um mesmo "Cód. Produto" aparece em mais de uma linha, cada linha carrega
    # seu próprio Estoque/Média, e "Dias Restantes de Estoque" é calculado por
    # LINHA. Ao filtrar o código no Excel, é fácil comparar por engano o Estoque
    # de uma linha com o resultado calculado a partir de outra linha do mesmo
    # código — parecendo um erro de fórmula quando na verdade é dado duplicado.
    dup_mask = ruptura["Cód. Produto"].duplicated(keep=False)
    codigos_duplicados_df = (
        ruptura.loc[
            dup_mask,
            ["Cód. Produto", "Descrição", "Estoque", "Média Qtd", "Valor C/Ruptura"],
        ]
        .sort_values("Cód. Produto")
        .copy()
    )
    n_codigos_duplicados = codigos_duplicados_df["Cód. Produto"].nunique()
    n_linhas_duplicadas = len(codigos_duplicados_df)
    if n_codigos_duplicados:
        print(
            f"\nATENÇÃO: {n_codigos_duplicados} código(s) de produto aparecem "
            f"mais de uma vez no arquivo de Ruptura ({n_linhas_duplicadas} linhas "
            "no total). Veja a aba 'Diagnostico_Duplicados' no Excel gerado — "
            "isso pode distorcer contagens de SKUs e a leitura de 'Dias Restantes "
            "de Estoque' se você comparar Estoque de uma linha com o resultado de outra."
        )

    def parse_moeda(v):
        if pd.isna(v):
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        v = str(v).replace("R$", "").strip().replace(".", "").replace(",", ".")
        try:
            return float(v)
        except ValueError:
            return 0.0

    ruptura["Valor_C_Ruptura_Num"] = ruptura["Valor C/Ruptura"].apply(parse_moeda)

    base = ruptura.merge(
        agenda_agg, left_on="Cód. Produto", right_on="Material", how="left"
    )
    base["Agendado"] = np.where(base["Material"].notna(), "Sim", "Não")
    base["Qtd_Agendada_Total"] = base["Qtd_Agendada_Total"].fillna(0)
    base["Ultima Entrada_dt"] = pd.to_datetime(
        base["Ultima Entrada"], format="%d/%m/%Y", errors="coerce"
    )
    base["Ultima Saida_dt"] = pd.to_datetime(
        base["Ultima\nSaida CD"], format="%d/%m/%Y", errors="coerce"
    )
    base["Media_Diaria"] = pd.to_numeric(base["Média Qtd"], errors="coerce")
    base["Estoque_Num"] = pd.to_numeric(base["Estoque"], errors="coerce").fillna(0)
    with np.errstate(divide="ignore", invalid="ignore"):
        dias_bruto = np.where(
            base["Media_Diaria"] > 0, base["Estoque_Num"] / base["Media_Diaria"], np.nan
        )
    base["Dias_Restantes_Estoque"] = pd.Series(dias_bruto).apply(_formatar_dias_exibicao).values
    base["Tem_Estoque"] = np.where(base["Estoque_Num"] > 0, "Sim", "Não")
    base["Tem_Agenda"] = np.where(
        base["Qtd_Agendada_Total"].fillna(0) > 0, "Sim", "Não"
    )
    base["Cód. + Descrição"] = base.apply(
        lambda row: _rotulo_cod_desc(row["Cód. Produto"], row["Descrição"]), axis=1
    )
    base["Confronto_Estoque_Agenda"] = np.select(
        [
            (base["Tem_Estoque"] == "Sim") & (base["Tem_Agenda"] == "Sim"),
            (base["Tem_Estoque"] == "Sim") & (base["Tem_Agenda"] == "Não"),
            (base["Tem_Estoque"] == "Não") & (base["Tem_Agenda"] == "Sim"),
        ],
        ["Estoque + Agenda", "Estoque sem Agenda", "Agenda sem Estoque"],
        default="Sem estoque e sem agenda",
    )

    zerados = base[base["Estoque"] == 0].copy()
    alerta_ampliado = base[base["Ruptura"] == "C/ Ruptura"].copy()
    alerta_ampliado_estoque_positivo = alerta_ampliado[
        alerta_ampliado["Estoque"] > 0
    ].copy()
    zerados_sem_agenda = zerados[zerados["Agendado"] == "Não"].copy()
    zerados_com_agenda = zerados[zerados["Agendado"] == "Sim"].copy()
    zerados_sem_compra = zerados_sem_agenda[
        zerados_sem_agenda["Comprado"] == "Não Comprado"
    ].copy()

    # --- Compra possivelmente desnecessária: estoque alto (giro baixo) + agenda ---
    # Critério (ver Notas_Metodológicas, item 8): Estoque > 0, existe carga
    # agendada (Qtd. Agendada > 0) e a cobertura atual já é alta (>= LIMIAR
    # dias), ou o item não tem histórico de saída (Média Qtd = 0) mesmo tendo
    # estoque — nesses dois casos, receber mais estoque é questionável.
    base["Dias_Cobertura_Pos_Agenda"] = np.where(
        base["Media_Diaria"] > 0,
        (base["Estoque_Num"] + base["Qtd_Agendada_Total"].fillna(0))
        / base["Media_Diaria"],
        np.nan,
    )
    media_valor_num = pd.to_numeric(base["Média Valor"], errors="coerce").fillna(0)
    media_qtd_num = pd.to_numeric(base["Média Qtd"], errors="coerce").fillna(0)
    valor_unitario_estimado = np.where(
        media_qtd_num > 0, media_valor_num / media_qtd_num, 0.0
    )
    base["Valor_Estimado_Compra_Extra"] = (
        base["Qtd_Agendada_Total"].fillna(0) * valor_unitario_estimado
    )

    itens_estoque_e_agenda_mask = (base["Estoque_Num"] > 0) & (
        base["Qtd_Agendada_Total"].fillna(0) > 0
    )
    giro_baixo_mask = (base["Dias_Restantes_Estoque"] >= LIMIAR_DIAS_GIRO_BAIXO) | (
        base["Media_Diaria"] == 0
    )
    compra_desnecessaria = base[itens_estoque_e_agenda_mask & giro_baixo_mask].copy()
    compra_desnecessaria = compra_desnecessaria.sort_values(
        "Estoque_Num", ascending=False
    )
    total_compra_desnecessaria = len(compra_desnecessaria)
    valor_compra_desnecessaria = compra_desnecessaria["Valor_Estimado_Compra_Extra"].sum()
    total_itens_estoque_e_agenda = int(itens_estoque_e_agenda_mask.sum())
    total_agenda_coerente = total_itens_estoque_e_agenda - total_compra_desnecessaria

    # --- Cruzamento com o relatório de posições no armazém (420_NORMA) ---
    # Mostra, para os itens já sinalizados como "compra possivelmente
    # desnecessária" (giro baixo + estoque alto + mais chegando), EM QUAIS
    # posições físicas do CD esse estoque parado está ocupando espaço.
    norma_posicoes = pd.DataFrame()
    norma_resumo = pd.DataFrame()
    norma_arquivo_ok = False
    if norma_path and Path(norma_path).exists():
        try:
            norma_raw = pd.read_csv(
                norma_path, sep=";", encoding="latin1", skiprows=18, dtype=str
            )
            norma_raw.columns = [c.strip() for c in norma_raw.columns]
            norma_raw["Cód. Produto"] = pd.to_numeric(
                norma_raw["Codigo"], errors="coerce"
            )
            norma_raw = norma_raw.dropna(subset=["Cód. Produto"])
            norma_raw["Cód. Produto"] = norma_raw["Cód. Produto"].astype(np.int64)
            qtd_um = pd.to_numeric(norma_raw["UM"], errors="coerce").fillna(0)
            qtd_uge = pd.to_numeric(norma_raw["UGE1/UM"], errors="coerce").fillna(0)
            norma_raw["Quantidade"] = qtd_um * qtd_uge
            norma_raw["Posição"] = norma_raw["Nº local"]
            norma_raw["Validade"] = pd.to_datetime(
                norma_raw["Data val. real"], format="%d/%m/%Y", errors="coerce"
            )
            norma_posicoes = norma_raw[
                [
                    "Cód. Produto",
                    "descricao produto",
                    "Posição",
                    "Quantidade",
                    "Lote",
                    "Validade",
                ]
            ].rename(columns={"descricao produto": "Descrição"})
            norma_resumo = (
                norma_posicoes.groupby(["Cód. Produto", "Descrição"])
                .agg(
                    Qtde_Total_420=("Quantidade", "sum"),
                    N_Posicoes=("Posição", "nunique"),
                    Validade_Mais_Proxima=("Validade", "min"),
                )
                .reset_index()
            )
            norma_arquivo_ok = True
            print(
                f"\n420_NORMA carregado: {norma_raw['Cód. Produto'].nunique()} "
                f"código(s), {len(norma_posicoes)} posições no total."
            )
        except Exception as exc:
            print(f"\nAVISO: não consegui ler o arquivo 420_NORMA ({norma_path}): {exc}")
    else:
        print(
            "\nAVISO: arquivo 420_NORMA não informado/encontrado — a aba e o "
            "gráfico de posições no armazém não serão gerados nesta execução."
        )

    compra_desnecessaria_com_posicao = pd.DataFrame()
    if norma_arquivo_ok and not norma_resumo.empty:
        compra_desnecessaria_com_posicao = compra_desnecessaria.merge(
            norma_resumo[
                ["Cód. Produto", "Qtde_Total_420", "N_Posicoes", "Validade_Mais_Proxima"]
            ],
            on="Cód. Produto",
            how="inner",
        ).sort_values("Qtde_Total_420", ascending=False)
    total_itens_cruzados_420 = len(compra_desnecessaria_com_posicao)

    materiais_ruptura = set(ruptura["Cód. Produto"].astype(np.int64))
    agenda_extra = agenda[~agenda["Material"].isin(materiais_ruptura)].copy()
    agenda_extra_agg = (
        agenda_extra.groupby("Material")
        .agg(
            Qtd_Agendada_Total=("Qtd_Agendada", "sum"),
            Data_Prevista_Mais_Proxima=("Data_Prevista", "min"),
            Descricao=("Texto breve material", "first"),
            Fornecedores_Agenda=(
                "Fornecedor_Nome_Agenda",
                lambda s: "; ".join(sorted(set(str(x) for x in s))),
            ),
            Status_Cargas=(
                "Status_Carga",
                lambda s: "; ".join(sorted(set(str(x) for x in s))),
            ),
        )
        .reset_index()
    )
    n_agenda_sem_match = len(agenda_extra_agg)

    proj = base.copy()
    if n_agenda_sem_match:
        linhas_extra = pd.DataFrame(
            {
                "Cód. Produto": agenda_extra_agg["Material"],
                "Descrição": agenda_extra_agg["Descricao"],
                "Fornecedor": agenda_extra_agg["Fornecedores_Agenda"],
                "Categoria": "N/D (não consta na Ruptura)",
                "Estoque": 0,
                "Agendado": "Sim",
                "Qtd_Agendada_Total": agenda_extra_agg["Qtd_Agendada_Total"],
                "Data_Prevista_Mais_Proxima": agenda_extra_agg[
                    "Data_Prevista_Mais_Proxima"
                ],
                "Fornecedores_Agenda": agenda_extra_agg["Fornecedores_Agenda"],
                "Status_Cargas": agenda_extra_agg["Status_Cargas"],
                "Média Qtd": np.nan,
                "Média Valor": 0.0,
                "Ultima Entrada_dt": pd.NaT,
                "Ultima Saida_dt": pd.NaT,
                "Valor_C_Ruptura_Num": 0.0,
                "Comprado": "N/D",
                "Bloq": "N/D",
                "Motivo Bloq.": "N/D",
                "Curva": "N/D",
                "Media_Diaria": np.nan,
                "Estoque_Num": 0.0,
                "Dias_Restantes_Estoque": np.nan,
                "Tem_Estoque": "Não",
                "Tem_Agenda": "Sim",
                "Confronto_Estoque_Agenda": "Agenda sem Estoque",
            }
        )
        proj = pd.concat([proj, linhas_extra], ignore_index=True)

    media_estimada_global = ruptura.loc[ruptura["Média Qtd"] > 0, "Média Qtd"].median()
    proj["Media_Diaria"] = pd.to_numeric(proj["Média Qtd"], errors="coerce")
    proj["Media_Estimada"] = proj["Media_Diaria"].isna()
    proj["Media_Diaria"] = proj["Media_Diaria"].fillna(media_estimada_global)

    proj["Estoque_Atual"] = proj["Estoque"].fillna(0).astype(float)
    proj["Qtd_Agendada"] = proj["Qtd_Agendada_Total"].fillna(0).astype(float)
    proj["Estoque_Total_Com_Agenda"] = proj["Estoque_Atual"] + proj["Qtd_Agendada"]

    proj["Dias_Restantes_Estoque"] = np.where(
        proj["Media_Diaria"] > 0,
        proj["Estoque_Atual"] / proj["Media_Diaria"],
        np.nan,
    )
    proj["Dias_Restantes_Estoque"] = pd.Series(proj["Dias_Restantes_Estoque"]).apply(_formatar_dias_exibicao).values

    media_arr = proj["Media_Diaria"].values.astype(float)
    estoque_arr = proj["Estoque_Atual"].values.astype(float)
    total_arr = proj["Estoque_Total_Com_Agenda"].values.astype(float)

    with np.errstate(divide="ignore", invalid="ignore"):
        dias_sem = np.where(
            media_arr > 0,
            estoque_arr / media_arr,
            np.where(estoque_arr <= 0, 0, np.inf),
        )
        dias_com = np.where(
            media_arr > 0,
            total_arr / media_arr,
            np.where(total_arr <= 0, 0, np.inf),
        )

    proj["Dias_Ate_Zerar_Sem_Agenda"] = pd.Series(dias_sem).apply(_formatar_dias_exibicao).values
    proj["Dias_Ate_Zerar_Com_Agenda"] = pd.Series(dias_com).apply(_formatar_dias_exibicao).values
    proj["Estoque_Proj_30d_Sem_Agenda"] = np.maximum(estoque_arr - media_arr * 30, 0)
    proj["Estoque_Proj_30d_Com_Agenda"] = np.maximum(total_arr - media_arr * 30, 0)
    proj["Estoque_Proj_60d_Sem_Agenda"] = np.maximum(estoque_arr - media_arr * 60, 0)
    proj["Estoque_Proj_60d_Com_Agenda"] = np.maximum(total_arr - media_arr * 60, 0)

    zerado_hoje     = estoque_arr <= 0          # SEM agenda: só o estoque físico atual
    zerado_hoje_com = total_arr   <= 0          # COM agenda: estoque físico + toda a agenda prevista
    zerado_30_sem = proj["Estoque_Proj_30d_Sem_Agenda"].values <= 0
    zerado_30_com = proj["Estoque_Proj_30d_Com_Agenda"].values <= 0
    zerado_60_sem = proj["Estoque_Proj_60d_Sem_Agenda"].values <= 0
    zerado_60_com = proj["Estoque_Proj_60d_Com_Agenda"].values <= 0

    def _situacao(i):
        if zerado_hoje[i]:
            return "Já zerado hoje"
        if zerado_30_com[i]:
            return "Zera em ≤30 dias MESMO com Agenda (ação urgente)"
        if zerado_60_com[i]:
            return "Zera entre 31-60 dias mesmo com Agenda"
        if zerado_30_sem[i]:
            return "Zeraria em ≤30d sem Agenda — Agenda resolve"
        if zerado_60_sem[i]:
            return "Zeraria em ≤60d sem Agenda — Agenda resolve"
        return "Sem risco projetado em 60 dias"

    proj["Situação_Projeção"] = [_situacao(i) for i in range(len(proj))]
    proj["Cobertura_Legivel_Sem_Agenda"] = [
        _cobertura_legivel(dias_sem[i], zerado_hoje[i]) for i in range(len(proj))
    ]
    proj["Cobertura_Legivel_Com_Agenda"] = [
        _cobertura_legivel(dias_com[i], total_arr[i] <= 0) for i in range(len(proj))
    ]

    dias_range = np.arange(0, 61, 5)  # 0,5,10,...,60 — passo de 5 dias
    curva_sem = (dias_sem[:, None] <= dias_range[None, :]).sum(axis=0)
    curva_com = (dias_com[:, None] <= dias_range[None, :]).sum(axis=0)

    mask_relevante = zerado_60_sem | zerado_60_com
    out_projecao_itens = (
        proj.loc[
            mask_relevante,
            [
                "Cód. Produto",
                "Descrição",
                "Fornecedor",
                "Categoria",
                "Estoque_Atual",
                "Qtd_Agendada",
                "Estoque_Total_Com_Agenda",
                "Media_Diaria",
                "Media_Estimada",
                "Dias_Ate_Zerar_Sem_Agenda",
                "Dias_Ate_Zerar_Com_Agenda",
                "Cobertura_Legivel_Sem_Agenda",
                "Cobertura_Legivel_Com_Agenda",
                "Estoque_Proj_30d_Sem_Agenda",
                "Estoque_Proj_30d_Com_Agenda",
                "Estoque_Proj_60d_Sem_Agenda",
                "Estoque_Proj_60d_Com_Agenda",
                "Situação_Projeção",
                "Comprado",
                "Bloq",
                "Curva",
            ],
        ]
        .rename(
            columns={
                "Estoque_Atual": "Estoque Atual",
                "Qtd_Agendada": "Qtd. Agendada",
                "Estoque_Total_Com_Agenda": "Estoque Total (c/ Agenda)",
                "Media_Diaria": "Média Saída Diária (un)",
                "Media_Estimada": "Média Estimada? (sem histórico próprio)",
                "Dias_Ate_Zerar_Sem_Agenda": "Dias até Zerar (sem Agenda)",
                "Dias_Ate_Zerar_Com_Agenda": "Dias até Zerar (com Agenda)",
                "Cobertura_Legivel_Sem_Agenda": "Cobertura Legível (sem Agenda)",
                "Cobertura_Legivel_Com_Agenda": "Cobertura Legível (com Agenda)",
                "Estoque_Proj_30d_Sem_Agenda": "Estoque em 30d (sem Agenda)",
                "Estoque_Proj_30d_Com_Agenda": "Estoque em 30d (com Agenda)",
                "Estoque_Proj_60d_Sem_Agenda": "Estoque em 60d (sem Agenda)",
                "Estoque_Proj_60d_Com_Agenda": "Estoque em 60d (com Agenda)",
            }
        )
        .sort_values("Dias até Zerar (com Agenda)", ascending=True)
    )

    situacao_counts = out_projecao_itens["Situação_Projeção"].value_counts()

    n_universo = len(proj)

    projecao_kpis = {
        "Universo total (Ruptura + itens só na Agenda)": n_universo,
        "Itens só na Agenda sem correspondência na Ruptura (média estimada)": n_agenda_sem_match,
        "Zerados HOJE — SEM Agenda (estoque físico ≤ 0)": int(zerado_hoje.sum()),
        "Zerados HOJE — COM Agenda (estoque físico + agenda ≤ 0)": int(zerado_hoje_com.sum()),
        "--- Projeção 30 dias ---": "",
        "Zerados projetados em 30 dias — SEM Agenda": int(zerado_30_sem.sum()),
        "Zerados projetados em 30 dias — COM Agenda": int(zerado_30_com.sum()),
        "Itens que a Agenda evita zerar até 30 dias": int(
            (zerado_30_sem & ~zerado_30_com).sum()
        ),
        "Itens NOVOS que vão zerar em 30 dias MESMO com Agenda (ação urgente)": int(
            (zerado_30_com & ~zerado_hoje).sum()
        ),
        "--- Projeção 60 dias ---": "",
        "Zerados projetados em 60 dias — SEM Agenda": int(zerado_60_sem.sum()),
        "Zerados projetados em 60 dias — COM Agenda": int(zerado_60_com.sum()),
        "Itens que a Agenda evita zerar até 60 dias": int(
            (zerado_60_sem & ~zerado_60_com).sum()
        ),
        "Itens NOVOS que vão zerar em 60 dias MESMO com Agenda (ação urgente)": int(
            (zerado_60_com & ~zerado_hoje).sum()
        ),
    }

    curva_df = pd.DataFrame(
        {
            "Dia": dias_range,
            "Qtde Zerados SEM Agenda": curva_sem,
            "Qtde Zerados COM Agenda": curva_com,
            "% Zerados SEM Agenda": curva_sem / len(proj),
            "% Zerados COM Agenda": curva_com / len(proj),
        }
    )

    total_itens = len(base)
    total_zerados = len(zerados)
    total_zerados_com_agenda = len(zerados_com_agenda)
    total_zerados_sem_agenda = len(zerados_sem_agenda)
    total_zerados_sem_compra = len(zerados_sem_compra)
    total_alerta_ampliado = len(alerta_ampliado)
    total_alerta_estoque_positivo = len(alerta_ampliado_estoque_positivo)
    valor_zerados = zerados["Valor_C_Ruptura_Num"].sum()
    valor_zerados_sem_agenda = zerados_sem_agenda["Valor_C_Ruptura_Num"].sum()

    kpis = {
        "Data da análise": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "Total de itens na base (Ruptura)": total_itens,
        "% Ruptura geral (oficial QlikView)": pct_ruptura_geral_oficial,
        "Valor total C/Ruptura (oficial QlikView)": valor_ruptura_total_oficial,
        "Itens com Estoque = 0 (zerados)": total_zerados,
        "% de itens zerados sobre total": (
            total_zerados / total_itens if total_itens else 0
        ),
        "Zerados COM chegada agendada": total_zerados_com_agenda,
        "Zerados SEM nenhuma chegada agendada": total_zerados_sem_agenda,
        "Zerados sem agenda E sem pedido de compra (crítico)": total_zerados_sem_compra,
        "Valor (R$) parado em ruptura (zerados)": valor_zerados,
        "Valor (R$) em ruptura zerada sem previsão de chegada": valor_zerados_sem_agenda,
        "--- Achado adicional (definição ampliada do QlikView) ---": "",
        "Itens sinalizados como C/Ruptura pelo QlikView (flag oficial)": total_alerta_ampliado,
        "  dos quais COM estoque físico > 0 mas < 1 dia de cobertura": total_alerta_estoque_positivo,
        "Itens com compra agendada aparentemente desnecessária (giro baixo)": total_compra_desnecessaria,
        "Valor estimado em compras possivelmente desnecessárias (R$)": valor_compra_desnecessaria,
    }

    cols_saida = [
        "Cód. Produto",
        "Descrição",
        "Cód. + Descrição",
        "Fornecedor",
        "Categoria",
        "Estoque",
        "Agendado",
        "Qtd_Agendada_Total",
        "Data_Prevista_Mais_Proxima",
        "Fornecedores_Agenda",
        "Status_Cargas",
        "Média Qtd",
        "Média Valor",
        "Ultima Entrada_dt",
        "Ultima Saida_dt",
        "Valor_C_Ruptura_Num",
        "Comprado",
        "Bloq",
        "Motivo Bloq.",
        "Curva",
        "Dias_Restantes_Estoque",
        "Tem_Estoque",
        "Tem_Agenda",
        "Confronto_Estoque_Agenda",
    ]

    def preparar_saida(df):
        d = df[cols_saida].copy()
        d = d.rename(
            columns={
                "Média Qtd": "Média Saída Diária (un)",
                "Média Valor": "Média Valor Saída (R$)",
                "Ultima Entrada_dt": "Última Entrada",
                "Ultima Saida_dt": "Última Saída CD",
                "Valor_C_Ruptura_Num": "Valor em Ruptura (R$)",
                "Dias_Restantes_Estoque": "Dias Restantes de Estoque",
            }
        )
        return d.sort_values("Valor em Ruptura (R$)", ascending=False)

    out_zerados = preparar_saida(zerados)
    out_zerados_sem_agenda = preparar_saida(zerados_sem_agenda)
    out_zerados_com_agenda = preparar_saida(zerados_com_agenda)
    out_alerta_ampliado = preparar_saida(alerta_ampliado_estoque_positivo)

    itens_estoque_agenda = base[
        (base["Estoque_Num"] > 0) & (base["Qtd_Agendada_Total"].fillna(0) > 0)
    ].copy()
    itens_estoque_agenda = (
        itens_estoque_agenda[
            [
                "Cód. Produto",
                "Descrição",
                "Cód. + Descrição",
                "Fornecedor",
                "Categoria",
                "Estoque",
                "Qtd_Agendada_Total",
                "Dias_Restantes_Estoque",
                "Média Qtd",
                "Confronto_Estoque_Agenda",
                "Comprado",
                "Bloq",
                "Curva",
            ]
        ]
        .rename(
            columns={
                "Qtd_Agendada_Total": "Qtd. Agendada",
                "Dias_Restantes_Estoque": "Dias Restantes de Estoque",
                "Média Qtd": "Média Saída Diária (un)",
            }
        )
        .sort_values("Dias Restantes de Estoque", ascending=True)
    )

    top10_maiores_estoques = itens_estoque_agenda.sort_values(
        "Estoque", ascending=False
    ).head(10)

    out_compra_desnecessaria = (
        compra_desnecessaria[
            [
                "Cód. Produto",
                "Descrição",
                "Cód. + Descrição",
                "Fornecedor",
                "Categoria",
                "Estoque_Num",
                "Qtd_Agendada_Total",
                "Dias_Restantes_Estoque",
                "Dias_Cobertura_Pos_Agenda",
                "Media_Diaria",
                "Valor_Estimado_Compra_Extra",
                "Data_Prevista_Mais_Proxima",
                "Fornecedores_Agenda",
                "Comprado",
                "Curva",
            ]
        ]
        .rename(
            columns={
                "Estoque_Num": "Estoque Atual",
                "Qtd_Agendada_Total": "Qtd. Agendada",
                "Dias_Restantes_Estoque": "Dias de Cobertura Atual (giro)",
                "Dias_Cobertura_Pos_Agenda": "Dias de Cobertura Após Chegar a Agenda",
                "Media_Diaria": "Média Saída Diária (un)",
                "Valor_Estimado_Compra_Extra": "Valor Estimado da Compra Extra (R$)",
                "Data_Prevista_Mais_Proxima": "Próxima Chegada Prevista",
                "Fornecedores_Agenda": "Fornecedor(es) Agenda",
            }
        )
        .copy()
    )
    top10_compra_desnecessaria = out_compra_desnecessaria.head(10)

    kpi_df = pd.DataFrame(list(kpis.items()), columns=["Indicador", "Valor"])
    proj_kpi_df = pd.DataFrame(
        list(projecao_kpis.items()), columns=["Indicador", "Valor"]
    )
    situacao_df = situacao_counts.rename_axis("Situação").reset_index(
        name="Qtde de Itens"
    )

    cols_item_tab = [
        "Cód. Produto",
        "Descrição",
        "Estoque Atual",
        "Qtd. Agendada",
        "Estoque Total (c/ Agenda)",
        "Média Saída Diária (un)",
        "Cobertura Legível (sem Agenda)",
        "Cobertura Legível (com Agenda)",
        "Situação_Projeção",
    ]
    top_itens_projecao_df = (
        out_projecao_itens[out_projecao_itens["Situação_Projeção"] != "Já zerado hoje"][
            cols_item_tab
        ]
        .head(40)
        .rename(columns={"Situação_Projeção": "Situação"})
    )
    item_tab_startrow = len(proj_kpi_df) + 3 + len(curva_df) + 1 + 100

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        kpi_df.to_excel(writer, sheet_name="Dashboard", index=False)
        out_zerados.to_excel(writer, sheet_name="Zerados_x_Agenda", index=False)
        out_zerados_sem_agenda.to_excel(
            writer, sheet_name="Criticos_Sem_Agenda", index=False
        )
        out_zerados_com_agenda.to_excel(
            writer, sheet_name="Zerados_Com_Agenda", index=False
        )
        out_alerta_ampliado.to_excel(
            writer, sheet_name="Alerta_Ampliado_Cobertura", index=False
        )
        itens_estoque_agenda.to_excel(
            writer, sheet_name="Estoque_x_Agenda", index=False
        )
        proj_kpi_df.to_excel(
            writer, sheet_name="Projecao_30_60_Dias", index=False, startrow=0
        )
        curva_df.to_excel(
            writer,
            sheet_name="Projecao_30_60_Dias",
            index=False,
            startrow=len(proj_kpi_df) + 3,
        )
        top_itens_projecao_df.to_excel(
            writer,
            sheet_name="Projecao_30_60_Dias",
            index=False,
            startrow=item_tab_startrow,
        )
        out_projecao_itens.to_excel(
            writer, sheet_name="Detalhe_Projecao_60d", index=False
        )
        situacao_df.to_excel(
            writer,
            sheet_name="Detalhe_Projecao_60d",
            index=False,
            startrow=0,
            startcol=len(out_projecao_itens.columns) + 2,
        )
        out_compra_desnecessaria.to_excel(
            writer, sheet_name="Compra_Desnecessaria", index=False
        )
        pd.DataFrame(
            {
                "Info": [
                    f"Critério: estoque > 0, com carga agendada (Qtd. Agendada > 0) "
                    f"e cobertura atual ≥ {LIMIAR_DIAS_GIRO_BAIXO} dias (ou sem "
                    "histórico de saída) — ver Notas_Metodológicas, item 8."
                ]
            }
        ).to_excel(
            writer,
            sheet_name="Compra_Desnecessaria",
            index=False,
            header=False,
            startrow=len(out_compra_desnecessaria) + 2,
        )
        if n_codigos_duplicados:
            codigos_duplicados_df.to_excel(
                writer, sheet_name="Diagnostico_Duplicados", index=False
            )
        else:
            pd.DataFrame(
                {
                    "Info": [
                        "Nenhum 'Cód. Produto' duplicado foi encontrado no "
                        "arquivo de Ruptura desta execução."
                    ]
                }
            ).to_excel(writer, sheet_name="Diagnostico_Duplicados", index=False)
        if norma_arquivo_ok:
            out_norma_posicoes = norma_posicoes[
                norma_posicoes["Cód. Produto"].isin(
                    compra_desnecessaria_com_posicao["Cód. Produto"]
                )
            ].sort_values(["Cód. Produto", "Quantidade"], ascending=[True, False])
            out_norma_posicoes.to_excel(
                writer, sheet_name="Posicoes_Estoque_420Norma", index=False
            )
            pd.DataFrame(
                {
                    "Info": [
                        "Posições físicas no CD (arquivo 420_NORMA) para os itens "
                        "já sinalizados na aba Compra_Desnecessaria — mostra ONDE "
                        "o estoque parado está fisicamente guardado. Só aparecem "
                        "aqui os itens que existem nos dois arquivos ao mesmo tempo."
                    ]
                }
            ).to_excel(
                writer,
                sheet_name="Posicoes_Estoque_420Norma",
                index=False,
                header=False,
                startrow=len(out_norma_posicoes) + 2,
            )

    _formatar(
        out_path,
        {
            "kpis": kpis,
            "proj_kpis": projecao_kpis,
            "curva_df": curva_df,
            "kpi_df": kpi_df,
            "proj_kpi_df": proj_kpi_df,
            "situacao_df": situacao_df,
            "out_zerados": out_zerados,
            "out_zerados_sem_agenda": out_zerados_sem_agenda,
            "out_zerados_com_agenda": out_zerados_com_agenda,
            "out_alerta_ampliado": out_alerta_ampliado,
            "itens_estoque_agenda": itens_estoque_agenda,
            "out_projecao_itens": out_projecao_itens,
            "top_itens_projecao_df": top_itens_projecao_df,
            "item_tab_startrow": item_tab_startrow,
            "n_universo": len(proj),
            "n_agenda_sem_match": n_agenda_sem_match,
            "total_itens": total_itens,
            "total_zerados": total_zerados,
            "total_zerados_com_agenda": total_zerados_com_agenda,
            "total_zerados_sem_agenda": total_zerados_sem_agenda,
            "total_zerados_sem_compra": total_zerados_sem_compra,
            "total_alerta_ampliado": total_alerta_ampliado,
            "total_alerta_estoque_positivo": total_alerta_estoque_positivo,
            "n_codigos_duplicados": n_codigos_duplicados,
            "n_linhas_duplicadas": n_linhas_duplicadas,
            "zerados_hoje_n":     int(zerado_hoje.sum()),      # SEM agenda
            "zerados_hoje_com_n": int(zerado_hoje_com.sum()),  # COM agenda (física + prevista)
            "zerados_30_sem_n": int(zerado_30_sem.sum()),
            "zerados_30_com_n": int(zerado_30_com.sum()),
            "zerados_60_sem_n": int(zerado_60_sem.sum()),
            "zerados_60_com_n": int(zerado_60_com.sum()),
            "agenda_evita_30": int((zerado_30_sem & ~zerado_30_com).sum()),
            "agenda_evita_60": int((zerado_60_sem & ~zerado_60_com).sum()),
            "critico_mesmo_com_agenda_30": int((zerado_30_com & ~zerado_hoje).sum()),
            "critico_mesmo_com_agenda_60": int((zerado_60_com & ~zerado_hoje).sum()),
            "n_top_itens_tab": len(top_itens_projecao_df),
            "n_situacao": len(situacao_df),
            "curva_linha_inicio": len(proj_kpi_df) + 3,
            "n_proj_kpis": len(proj_kpi_df),
            "n_itens_proj": len(out_projecao_itens),
            "out_compra_desnecessaria": out_compra_desnecessaria,
            "top10_maiores_estoques": top10_maiores_estoques,
            "n_compra_desnecessaria": total_compra_desnecessaria,
            "valor_compra_desnecessaria": valor_compra_desnecessaria,
            "total_agenda_coerente": total_agenda_coerente,
            "total_itens_estoque_e_agenda": total_itens_estoque_e_agenda,
            "norma_arquivo_ok": norma_arquivo_ok,
            "total_itens_cruzados_420": total_itens_cruzados_420,
            "compra_desnecessaria_com_posicao": compra_desnecessaria_com_posicao,
        },
    )
    print(f"\nPronto! Dashboard salvo em: {out_path}")


def _formatar(path, ctx):
    wb = load_workbook(path)
    moeda_cols = ["Média Valor Saída (R$)", "Valor em Ruptura (R$)"]
    data_cols = ["Data_Prevista_Mais_Proxima", "Última Entrada", "Última Saída CD"]

    # --- Aba única "GRAFICOS": todos os gráficos consolidados em um só lugar,
    # numa sequência lógica (panorama -> risco atual -> projeção -> excesso de
    # estoque), em vez de espalhados pelas abas de dados. As séries continuam
    # apontando para os dados reais nas outras abas; só o "quadro" visual muda.
    graficos_ws = wb.create_sheet("GRAFICOS", 0)
    graficos_ws.sheet_view.showGridLines = False

    # --- Cabeçalho com a logo do Grupo DPSP, se o arquivo estiver disponível ---
    logo_path = SCRIPT_DIR / "logo_dpsp.png"
    logo_ok = logo_path.exists()
    if logo_ok:
        try:
            img = XLImage(str(logo_path))
            img.height = 90
            img.width = 90
            graficos_ws.add_image(img, "A1")
        except Exception as exc:
            print(f"AVISO: não consegui inserir a logo ({exc}).")
            logo_ok = False
    graficos_ws.cell(
        row=1, column=4, value="Dashboard de Ruptura x Agenda — Grupo DPSP"
    ).font = Font(name=FONT_NAME, size=16, bold=True, color=NAVY)
    graficos_ws.cell(
        row=3, column=4, value=f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} | Helber Lima"
    ).font = Font(name=FONT_NAME, size=10, italic=True, color=COR_SECUNDARIA)

    # Layout replicado do modelo ajustado manualmente pelo usuário: cada chave
    # -> (coluna, linha, largura_cm, altura_cm, posição da legenda ou None).
    # Linhas deslocadas +9 em relação ao modelo original para abrir espaço
    # para o cabeçalho com a logo acima dos gráficos.
    _GRID = {
        "zerados_bar": ("A", 11, 20.52, 9.17, None),
        "pizza_zerados": ("N", 11, 21.88, 9.2, "r"),
        "necessaria_questionavel": ("A", 29, 20.43, 9.08, "r"),
        "situacao": ("N", 29, 22.0, 9.0, None),
        "hoje_30_60": ("A", 47, 20.46, 11.99, "r"),
        "top10_valor_ruptura": ("N", 47, 21.88, 12.0, None),
        "curva": ("A", 70, 43.54, 11.0, "b"),
        "top10_compra_desnecessaria_unificado": ("A", 94, 43.54, 13.0, None),
        "posicoes_420norma": ("A", 121, 20.46, 13.0, "r"),
        "top10_maiores_estoques": ("M", 121, 22.17, 13.0, "r"),
        "menor_cobertura": ("A", 147, 20.34, 14.7, None),
        "top_urgentes": ("M", 147, 22.02, 21.52, "b"),
    }

    # Mapeamento: chave do gráfico → (aba-fonte, descrição do link clicável)
    # O link é inserido logo acima de cada gráfico na aba GRAFICOS,
    # permitindo que o usuário clique e vá direto à aba com os dados detalhados.
    _FONTES = {
        "zerados_bar": (
            "Dashboard",
            "📋 Fonte: Dashboard — KPIs de ruptura e agenda (clique para abrir)",
        ),
        "pizza_zerados": (
            "Dashboard",
            "📋 Fonte: Dashboard — composição de itens zerados vs. com estoque (clique para abrir)",
        ),
        "necessaria_questionavel": (
            "Dashboard",
            "📋 Fonte: Dashboard — compras agendadas: necessárias x questionáveis (clique para abrir)",
        ),
        "situacao": (
            "Detalhe_Projecao_60d",
            "📋 Fonte: Detalhe_Projecao_60d — situação de cada item na projeção (clique para abrir)",
        ),
        "hoje_30_60": (
            "Projecao_30_60_Dias",
            "📋 Fonte: Projecao_30_60_Dias — resumo de zerados hoje / 30d / 60d (clique para abrir)",
        ),
        "top10_valor_ruptura": (
            "Zerados_x_Agenda",
            "📋 Fonte: Zerados_x_Agenda — top 10 itens por valor parado em ruptura (clique para abrir)",
        ),
        "curva": (
            "Projecao_30_60_Dias",
            "📋 Fonte: Projecao_30_60_Dias — curva de ruptura projetada em 60 dias (clique para abrir)",
        ),
        "top10_compra_desnecessaria_unificado": (
            "Compra_Desnecessaria",
            "📋 Fonte: Compra_Desnecessaria — top 10 compras possivelmente desnecessárias (clique para abrir)",
        ),
        "posicoes_420norma": (
            "Posicoes_Estoque_420Norma",
            "📋 Fonte: Posicoes_Estoque_420Norma — posições físicas no armazém 420_NORMA (clique para abrir)",
        ),
        "top10_maiores_estoques": (
            "Estoque_x_Agenda",
            "📋 Fonte: Estoque_x_Agenda — top 10 maiores estoques com agenda (clique para abrir)",
        ),
        "menor_cobertura": (
            "Estoque_x_Agenda",
            "📋 Fonte: Estoque_x_Agenda — itens com menor cobertura de estoque (clique para abrir)",
        ),
        "top_urgentes": (
            "Detalhe_Projecao_60d",
            "📋 Fonte: Detalhe_Projecao_60d — top itens mais urgentes da projeção (clique para abrir)",
        ),
    }

    def _colocar_grafico(chart, chave):
        """Posiciona um gráfico na aba GRAFICOS reproduzindo o layout, o
        tamanho e a legenda do modelo ajustado manualmente, e corrige o bug
        de gráfico ficar em branco quando linhas da tabela de origem são
        ocultadas pelo usuário.

        Além do gráfico, insere logo acima:
          • (opcional) a logo DPSP em miniatura — se logo_dpsp.png existir;
          • link clicável para a aba de dados de origem do gráfico.
        """
        col, row, largura, altura, legend_pos = _GRID[chave]
        chart.visible_cells_only = False
        chart.display_blanks = "gap"
        chart.width = largura
        chart.height = altura
        if legend_pos is None:
            chart.legend = None
        else:
            chart.legend = Legend(legendPos=legend_pos)
        graficos_ws.add_chart(chart, f"{col}{row}")

        # ── Linha de metadados (logo + link de fonte) ──────────────────────
        # Inserida na linha imediatamente acima do gráfico (row - 1).
        # A row mínima é 5 para não sobrescrever o cabeçalho principal (A1-A4).
        meta_row = max(5, row - 1)
        col_num = column_index_from_string(col)

        # Logo DPSP em miniatura (se disponível) — ancoramos na mesma célula
        # do link para ela aparecer à esquerda do texto de fonte.
        if logo_ok:
            try:
                logo_mini = XLImage(str(logo_path))
                logo_mini.height = 18   # altura em pontos (≈ 1 linha do Excel)
                logo_mini.width  = 54   # largura proporcional (~3:1)
                graficos_ws.add_image(logo_mini, f"{col}{meta_row}")
            except Exception:
                pass  # falha silenciosa — segue sem logo

        # Link clicável para a aba de dados de origem
        if chave in _FONTES:
            aba_fonte, descricao_fonte = _FONTES[chave]
            # Coluna do link: 1 à frente do logo (quando há logo) ou mesma coluna
            link_col_num = col_num + (2 if logo_ok else 0)
            cell_link = graficos_ws.cell(
                row=meta_row, column=link_col_num, value=descricao_fonte
            )
            # Hyperlink interno → aba de dados (abre direto em A1)
            aba_safe = f"'{aba_fonte}'" if " " in aba_fonte else aba_fonte
            cell_link.hyperlink = f"#{aba_safe}!A1"
            cell_link.font = Font(
                name=FONT_NAME,
                size=8,
                color="0563C1",
                underline="single",
                italic=True,
            )
            # Garante que a linha de metadados tenha altura suficiente
            if graficos_ws.row_dimensions[meta_row].height in (None, 0):
                graficos_ws.row_dimensions[meta_row].height = 18

    for sheet_name in [
        "Zerados_x_Agenda",
        "Criticos_Sem_Agenda",
        "Zerados_Com_Agenda",
        "Alerta_Ampliado_Cobertura",
        "Estoque_x_Agenda",
    ]:
        ws = wb[sheet_name]
        header = _estilizar_tabela(ws, moeda_cols=moeda_cols, data_cols=data_cols)
        col_idx = {name: i + 1 for i, name in enumerate(header)}

        if ws.max_row < 2:
            # Aba sem nenhuma linha de dados (só cabeçalho) — nada a formatar/plotar.
            continue

        if "Agendado" in col_idx:
            cl = get_column_letter(col_idx["Agendado"])
            ws.conditional_formatting.add(
                f"{cl}2:{cl}{ws.max_row}",
                CellIsRule(
                    operator="equal",
                    formula=['"Não"'],
                    fill=PatternFill("solid", fgColor=RED_FILL),
                ),
            )
        if "Comprado" in col_idx:
            cl = get_column_letter(col_idx["Comprado"])
            ws.conditional_formatting.add(
                f"{cl}2:{cl}{ws.max_row}",
                CellIsRule(
                    operator="equal",
                    formula=['"Não Comprado"'],
                    fill=PatternFill("solid", fgColor=ORANGE_FILL),
                ),
            )
        if "Valor em Ruptura (R$)" in col_idx:
            cl = get_column_letter(col_idx["Valor em Ruptura (R$)"])
            ws.conditional_formatting.add(
                f"{cl}2:{cl}{ws.max_row}",
                ColorScaleRule(
                    start_type="min",
                    start_color="D5F5E3",
                    end_type="max",
                    end_color="F1948A",
                ),
            )
        if "Dias Restantes de Estoque" in col_idx:
            cl = get_column_letter(col_idx["Dias Restantes de Estoque"])
            for r in range(2, ws.max_row + 1):
                ws.cell(
                    row=r, column=col_idx["Dias Restantes de Estoque"]
                ).number_format = "0.00"
        if sheet_name == "Estoque_x_Agenda" and "Confronto_Estoque_Agenda" in col_idx:
            cl = get_column_letter(col_idx["Confronto_Estoque_Agenda"])
            ws.conditional_formatting.add(
                f"{cl}2:{cl}{ws.max_row}",
                CellIsRule(
                    operator="equal",
                    formula=['"Estoque + Agenda"'],
                    fill=PatternFill("solid", fgColor="D5F5E3"),
                ),
            )
            ws.conditional_formatting.add(
                f"{cl}2:{cl}{ws.max_row}",
                CellIsRule(
                    operator="equal",
                    formula=['"Estoque sem Agenda"'],
                    fill=PatternFill("solid", fgColor="FDEBD0"),
                ),
            )
            ws.conditional_formatting.add(
                f"{cl}2:{cl}{ws.max_row}",
                CellIsRule(
                    operator="equal",
                    formula=['"Agenda sem Estoque"'],
                    fill=PatternFill("solid", fgColor="F5B7B1"),
                ),
            )
            top_n = min(10, ws.max_row - 1)
            chart = BarChart()
            chart.type = "bar"
            chart.style = 2
            chart.title = "Top itens com menor cobertura de estoque"
            chart.y_axis.title = "Dias restantes"
            data = Reference(
                ws,
                min_col=col_idx["Dias Restantes de Estoque"],
                min_row=1,
                max_row=top_n + 1,
            )
            cats = Reference(
                ws,
                min_col=col_idx["Cód. + Descrição"],
                min_row=2,
                max_row=top_n + 1,
            )
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.width = 24
            chart.height = 12
            chart.legend = None
            _sem_grade_vertical(chart)
            _rotular_dados(chart, formato="0")
            valores_dias = [
                ws.cell(row=r, column=col_idx["Dias Restantes de Estoque"]).value or 0
                for r in range(2, top_n + 2)
            ]
            cores_barras = [
                COR_CRITICO if v < 7 else COR_ATENCAO if v < 15 else COR_SECUNDARIA
                for v in valores_dias
            ]
            pontos = []
            for i, cor in enumerate(cores_barras):
                dp = DataPoint(idx=i)
                dp.graphicalProperties.solidFill = cor
                pontos.append(dp)
            chart.series[0].data_points = pontos
            _colocar_grafico(chart, "menor_cobertura")

    if "Compra_Desnecessaria" in wb.sheetnames:
        ws_cd = wb["Compra_Desnecessaria"]
        n_cd = ctx["n_compra_desnecessaria"]
        header_cd = _estilizar_tabela(
            ws_cd,
            moeda_cols=["Valor Estimado da Compra Extra (R$)"],
            data_cols=["Próxima Chegada Prevista"],
        )
        col_idx_cd = {name: i + 1 for i, name in enumerate(header_cd) if name}
        for col_nome in (
            "Dias de Cobertura Atual (giro)",
            "Dias de Cobertura Após Chegar a Agenda",
        ):
            if col_nome in col_idx_cd:
                for r in range(2, n_cd + 2):
                    ws_cd.cell(row=r, column=col_idx_cd[col_nome]).number_format = "0.0"
        nota_row = n_cd + 3
        ws_cd.cell(row=nota_row, column=1).font = Font(
            name=FONT_NAME, italic=True, size=10, color=NAVY
        )
        ws_cd.column_dimensions["B"].width = 32

    if "Diagnostico_Duplicados" in wb.sheetnames:
        ws_dup = wb["Diagnostico_Duplicados"]
        ws_dup.sheet_view.showGridLines = False
        if ctx["n_codigos_duplicados"]:
            _estilizar_tabela(
                ws_dup, moeda_cols=["Valor C/Ruptura"], data_cols=[]
            )
            for r in range(2, ws_dup.max_row + 1):
                for c in range(1, ws_dup.max_column + 1):
                    ws_dup.cell(row=r, column=c).fill = PatternFill(
                        "solid", fgColor=RED_FILL
                    )
            ws_dup.cell(row=ws_dup.max_row + 2, column=1).value = (
                f"{ctx['n_codigos_duplicados']} código(s) de produto duplicados "
                f"({ctx['n_linhas_duplicadas']} linhas). Recomenda-se validar com "
                "quem gera o export do QlikView se essas linhas devem ser somadas "
                "por código ou se representam registros distintos."
            )
            ws_dup.cell(row=ws_dup.max_row, column=1).font = Font(
                name=FONT_NAME, italic=True, size=10, color=NAVY
            )
        else:
            ws_dup.column_dimensions["A"].width = 90
            ws_dup.cell(row=1, column=1).font = Font(name=FONT_NAME, size=11)

    ws = wb["Detalhe_Projecao_60d"]
    header_proj = _estilizar_tabela(ws, moeda_cols=[], data_cols=[])
    col_idx_proj = {name: i + 1 for i, name in enumerate(header_proj) if name}

    if "Situação_Projeção" in col_idx_proj and ctx["n_itens_proj"] > 0:
        cl = get_column_letter(col_idx_proj["Situação_Projeção"])
        rng = f"{cl}2:{cl}{ctx['n_itens_proj'] + 1}"
        ws.conditional_formatting.add(
            rng,
            CellIsRule(
                operator="equal",
                formula=['"Zera em ≤30 dias MESMO com Agenda (ação urgente)"'],
                fill=PatternFill("solid", fgColor=RED_FILL),
            ),
        )
        ws.conditional_formatting.add(
            rng,
            CellIsRule(
                operator="equal",
                formula=['"Zera entre 31-60 dias mesmo com Agenda"'],
                fill=PatternFill("solid", fgColor=ORANGE_FILL),
            ),
        )
        ws.conditional_formatting.add(
            rng,
            CellIsRule(
                operator="equal",
                formula=['"Já zerado hoje"'],
                fill=PatternFill("solid", fgColor="D7BDE2"),
            ),
        )

    for nome_col in ["Dias até Zerar (sem Agenda)", "Dias até Zerar (com Agenda)"]:
        if nome_col in col_idx_proj:
            for r in range(2, ctx["n_itens_proj"] + 2):
                ws.cell(row=r, column=col_idx_proj[nome_col]).number_format = "00.000"
    for nome_col in [
        "Cobertura Legível (sem Agenda)",
        "Cobertura Legível (com Agenda)",
    ]:
        if nome_col in col_idx_proj:
            ws.column_dimensions[get_column_letter(col_idx_proj[nome_col])].width = 20

    # ANTES: col_sit = len(header_proj) + 3
    # BUG: header_proj já inclui as 2 colunas reais da tabela "Situação"/"Qtde de
    # Itens" (escritas por main() em X/Y) + 2 colunas de espaço em branco entre
    # elas e a tabela principal. Somar "+3" a esse total pulava 4 colunas à
    # frente dos dados reais, caindo em células em branco (por isso os gráficos
    # ficavam sem dados e as colunas AB/AC apareciam vazias). A correção busca a
    # coluna pelo NOME real já escrito, o que é robusto independente de gaps.
    col_sit = col_idx_proj.get("Situação")
    if col_sit is None:
        # Fallback defensivo, caso a coluna não seja encontrada por algum motivo:
        # conta apenas colunas com nome real (ignora os gaps em branco).
        col_sit = len([h for h in header_proj if h]) + 3
    max_row_sit = ctx["n_situacao"] + 1
    for c in (col_sit, col_sit + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(name=FONT_NAME, bold=True, color=HEADER_FONT_COLOR, size=10)
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
    ws.column_dimensions[get_column_letter(col_sit)].width = 48
    ws.column_dimensions[get_column_letter(col_sit + 1)].width = 16

    chart_sit = BarChart()
    chart_sit.type = "bar"
    chart_sit.title = "Itens em risco (≤60 dias) por situação"
    chart_sit.style = 2
    chart_sit.add_data(
        Reference(ws, min_col=col_sit + 1, min_row=1, max_row=max_row_sit),
        titles_from_data=True,
    )
    chart_sit.set_categories(
        Reference(ws, min_col=col_sit, min_row=2, max_row=max_row_sit)
    )
    chart_sit.width, chart_sit.height, chart_sit.legend = 22, 9, None
    _sem_grade_vertical(chart_sit)
    _rotular_dados(chart_sit, formato="0")

    def _cor_por_situacao(rotulo):
        rotulo = str(rotulo)
        if "Já zerado" in rotulo or "≤30 dias MESMO" in rotulo:
            return COR_CRITICO
        if "31-60 dias" in rotulo:
            return COR_ATENCAO
        if "Agenda resolve" in rotulo:
            return COR_OK
        return COR_SECUNDARIA

    pontos_sit = []
    for i in range(max_row_sit - 1):
        rotulo = ws.cell(row=2 + i, column=col_sit).value
        dp = DataPoint(idx=i)
        dp.graphicalProperties.solidFill = _cor_por_situacao(rotulo)
        pontos_sit.append(dp)
    chart_sit.series[0].data_points = pontos_sit
    _colocar_grafico(chart_sit, "situacao")

    col_top = col_sit + 3
    headers_top = [
        "Cód. + Descrição (Top 10)",
        "Descrição (Top 10)",
        "Dias até Zerar sem Agenda (Top 10)",
        "Dias até Zerar com Agenda (Top 10)",
        "Cobertura (Top 10, sem Agenda)",
        "Cobertura (Top 10, com Agenda)",
    ]
    for j, h in enumerate(headers_top):
        cell = ws.cell(row=1, column=col_top + j, value=h)
        cell.font = Font(name=FONT_NAME, bold=True, color=HEADER_FONT_COLOR, size=10)
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    col_cod = col_idx_proj["Cód. Produto"]
    col_desc = col_idx_proj["Descrição"]
    col_dsem = col_idx_proj["Dias até Zerar (sem Agenda)"]
    col_dcom = col_idx_proj["Dias até Zerar (com Agenda)"]
    col_legsem = col_idx_proj["Cobertura Legível (sem Agenda)"]
    col_legcom = col_idx_proj["Cobertura Legível (com Agenda)"]

    ainda_nao_zerado = [
        i + 2
        for i in range(ctx["n_itens_proj"])
        if ws.cell(row=i + 2, column=col_idx_proj["Situação_Projeção"]).value
        != "Já zerado hoje"
    ]
    linhas_top = ainda_nao_zerado[:10]
    for i, r_origem in enumerate(linhas_top):
        r_destino = i + 2
        ws.cell(
            row=r_destino,
            column=col_top,
            value=_rotulo_cod_desc(
                ws.cell(row=r_origem, column=col_cod).value,
                ws.cell(row=r_origem, column=col_desc).value,
            ),
        )
        ws.cell(
            row=r_destino,
            column=col_top + 1,
            value=ws.cell(row=r_origem, column=col_desc).value,
        )
        d_sem = ws.cell(row=r_origem, column=col_dsem).value
        d_com = ws.cell(row=r_origem, column=col_dcom).value
        ws.cell(
            row=r_destino,
            column=col_top + 2,
            value=d_sem if (d_sem is not None and np.isfinite(d_sem)) else None,
        ).number_format = "0.000"
        ws.cell(
            row=r_destino,
            column=col_top + 3,
            value=d_com if (d_com is not None and np.isfinite(d_com)) else None,
        ).number_format = "0.000"
        ws.cell(
            row=r_destino,
            column=col_top + 4,
            value=ws.cell(row=r_origem, column=col_legsem).value,
        )
        ws.cell(
            row=r_destino,
            column=col_top + 5,
            value=ws.cell(row=r_origem, column=col_legcom).value,
        )

    ws.column_dimensions[get_column_letter(col_top)].width = 14
    ws.column_dimensions[get_column_letter(col_top + 1)].width = 42
    ws.column_dimensions[get_column_letter(col_top + 2)].width = 20
    ws.column_dimensions[get_column_letter(col_top + 3)].width = 20
    ws.column_dimensions[get_column_letter(col_top + 4)].width = 20
    ws.column_dimensions[get_column_letter(col_top + 5)].width = 20
    ws.cell(
        row=len(linhas_top) + 3,
        column=col_top,
        value="Obs: exclui itens já zerados hoje — foco em alerta preventivo.",
    )
    ws.cell(row=len(linhas_top) + 3, column=col_top).font = Font(
        name=FONT_NAME, italic=True, size=9
    )

    chart_top = BarChart()
    chart_top.type = "bar"
    chart_top.title = (
        "Top itens AINDA NÃO zerados mais urgentes — dias até zerar (sem x com Agenda)"
    )
    chart_top.y_axis.title = "Dias até zerar"
    chart_top.style = 2
    chart_top.add_data(
        Reference(
            ws,
            min_col=col_top + 2,
            max_col=col_top + 3,
            min_row=1,
            max_row=len(linhas_top) + 1,
        ),
        titles_from_data=True,
    )
    chart_top.set_categories(
        Reference(ws, min_col=col_top, min_row=2, max_row=len(linhas_top) + 1)
    )
    chart_top.width, chart_top.height = 26, 16
    chart_top.y_axis.scaling.orientation = "maxMin"
    _sem_grade_vertical(chart_top)
    _rotular_dados(chart_top, formato="0")
    _colorir_series(chart_top, [COR_CRITICO, COR_OK])
    _colocar_grafico(chart_top, "top_urgentes")

    ws = wb["Projecao_30_60_Dias"]
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 20
    ws.cell(row=1, column=1).font = Font(
        name=FONT_NAME, bold=True, color=HEADER_FONT_COLOR, size=10
    )
    ws.cell(row=1, column=2).font = Font(
        name=FONT_NAME, bold=True, color=HEADER_FONT_COLOR, size=10
    )
    ws.cell(row=1, column=1).fill = PatternFill("solid", fgColor=HEADER_FILL)
    ws.cell(row=1, column=2).fill = PatternFill("solid", fgColor=HEADER_FILL)

    for r in range(2, ctx["n_proj_kpis"] + 1):
        ind, val = ws.cell(row=r, column=1), ws.cell(row=r, column=2)
        eh_separador = str(ind.value).startswith("---")
        ind.font = Font(name=FONT_NAME, size=10, bold=eh_separador)
        val.font = Font(name=FONT_NAME, size=10)
        if eh_separador:
            ind.fill = PatternFill("solid", fgColor=GREY_FILL)
            val.fill = PatternFill("solid", fgColor=GREY_FILL)

    ws.insert_rows(1, amount=2)
    ws["A1"] = "PROJEÇÃO DE RUPTURA — PRÓXIMOS 30 E 60 DIAS (com Agenda x sem Agenda)"
    ws["A1"].font = Font(name=FONT_NAME, bold=True, size=14, color=NAVY)
    ws.merge_cells("A1:B1")

    n_dias_curva = len(ctx["curva_df"])  # nº de linhas da tabela (13, passo de 5d)
    tabela_curva_row = ctx["curva_linha_inicio"] + 2
    header_curva_row = tabela_curva_row + 1
    for c in range(1, 6):
        cell = ws.cell(row=header_curva_row, column=c)
        cell.font = Font(name=FONT_NAME, bold=True, color=HEADER_FONT_COLOR, size=10)
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
    for r in range(header_curva_row + 1, header_curva_row + 1 + n_dias_curva):
        for c in (4, 5):
            ws.cell(row=r, column=c).number_format = "0.0%"
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 20

    line = LineChart()
    line.title = "Curva de ruptura projetada — 60 dias (% de SKUs em ruptura + Qtde)"
    line.y_axis.title = "% de SKUs em ruptura (eixo esquerdo)"
    line.x_axis.title = "Dias a partir de hoje"
    line.style = 2
    # Dados para as linhas (percentuais)
    data_ref = Reference(
        ws,
        min_col=4,
        max_col=5,
        min_row=header_curva_row,
        max_row=header_curva_row + n_dias_curva,
    )
    cats_ref = Reference(
        ws,
        min_col=1,
        min_row=header_curva_row + 1,
        max_row=header_curva_row + n_dias_curva,
    )
    line.add_data(data_ref, titles_from_data=True)
    line.set_categories(cats_ref)
    line.width, line.height = 26, 11
    line.y_axis.numFmt = "0%"
    _sem_grade_vertical(line)
    cores_linha = [COR_CRITICO, COR_CURVA_COM_AGENDA_LINHA]
    for s, cor in zip(line.series, cores_linha):
        s.smooth = False
        s.marker.symbol = "none"
        s.graphicalProperties.line.solidFill = cor
        s.graphicalProperties.line.width = 22000

    # Gráfico de barras para quantidade de SKUs (série secundária)
    bar_qty = BarChart()
    bar_qty.type = "col"
    bar_qty.title = None
    # Dados para as barras (quantidades)
    data_qty_ref = Reference(
        ws,
        min_col=2,
        max_col=3,
        min_row=header_curva_row,
        max_row=header_curva_row + n_dias_curva,
    )
    bar_qty.add_data(data_qty_ref, titles_from_data=True)
    bar_qty.set_categories(cats_ref)
    bar_qty.y_axis.title = "Quantidade de SKUs zerados (eixo direito)"
    bar_qty.y_axis.axId = 200
    bar_qty.y_axis.crosses = "max"
    _colorir_series(
        bar_qty, [COR_CURVA_SEM_AGENDA_BARRA, COR_CURVA_COM_AGENDA_BARRA]
    )
    bar_qty.legend = None
    _rotular_dados(bar_qty, formato="0")
    bar_qty.dataLabels.position = "outEnd"

    # Combinar: linha + barras
    line += bar_qty

    _colocar_grafico(line, "curva")

    resumo_start = header_curva_row + n_dias_curva + 3
    ws.cell(row=resumo_start, column=1, value="Horizonte")
    ws.cell(row=resumo_start, column=2, value="SEM Agenda")
    ws.cell(row=resumo_start, column=3, value="COM Agenda")

    valores = {
        ("Em 30 dias", "sem"): ctx["proj_kpis"][
            "Zerados projetados em 30 dias — SEM Agenda"
        ],
        ("Em 30 dias", "com"): ctx["proj_kpis"][
            "Zerados projetados em 30 dias — COM Agenda"
        ],
        ("Em 60 dias", "sem"): ctx["proj_kpis"][
            "Zerados projetados em 60 dias — SEM Agenda"
        ],
        ("Em 60 dias", "com"): ctx["proj_kpis"][
            "Zerados projetados em 60 dias — COM Agenda"
        ],
    }

    # "Hoje SEM Agenda" = itens com estoque físico ≤ 0  (zerado_hoje)
    # "Hoje COM Agenda" = itens que ficam zerados MESMO somando toda a agenda prevista
    #                     (zerado_hoje_com = total_arr ≤ 0).
    # Esta linha serve de BASE DE COMPARAÇÃO: mostra quantos itens a agenda do dia
    # já resolve vs. quantos continuam críticos mesmo com o recebimento esperado.
    ws.cell(row=resumo_start + 1, column=1, value="Hoje (*)")
    ws.cell(
        row=resumo_start + 1,
        column=2,
        value=ctx["zerados_hoje_n"],        # SEM agenda: estoque físico atual ≤ 0
    )
    ws.cell(
        row=resumo_start + 1,
        column=3,
        value=ctx["zerados_hoje_com_n"],    # COM agenda: físico + previsto ≤ 0
    )
    ws.cell(row=resumo_start + 2, column=1, value="Em 30 dias")
    ws.cell(row=resumo_start + 2, column=2, value=valores[("Em 30 dias", "sem")])
    ws.cell(row=resumo_start + 2, column=3, value=valores[("Em 30 dias", "com")])
    ws.cell(row=resumo_start + 3, column=1, value="Em 60 dias")
    ws.cell(row=resumo_start + 3, column=2, value=valores[("Em 60 dias", "sem")])
    ws.cell(row=resumo_start + 3, column=3, value=valores[("Em 60 dias", "com")])

    # Nota explicativa sobre o (*) e a interpretação das duas colunas no "Hoje"
    nota_proj_row = resumo_start + 5
    agenda_resolve_hoje = ctx["zerados_hoje_n"] - ctx["zerados_hoje_com_n"]
    nota_texto = (
        f"(*) Hoje SEM Agenda ({ctx['zerados_hoje_n']}): itens com estoque físico zerado agora. "
        f"Hoje COM Agenda ({ctx['zerados_hoje_com_n']}): itens que continuam zerados MESMO "
        "somando todo o estoque previsto na agenda — servem de BASE REAL de ruptura após o recebimento. "
        f"A diferença ({agenda_resolve_hoje} item(ns)) representa o que a agenda do dia já resolve hoje. "
        "Nas projeções de 30 e 60 dias, a agenda acumula ainda mais impacto positivo. "
        f"Universo: {ctx['n_universo']} itens "
        f"(Ruptura + {ctx['n_agenda_sem_match']} item(ns) só na Agenda, média de saída estimada)."
    )
    cell_nota = ws.cell(row=nota_proj_row, column=1, value=nota_texto)
    cell_nota.font = Font(name=FONT_NAME, italic=True, size=9, color=COR_SECUNDARIA)
    cell_nota.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(f"A{nota_proj_row}:C{nota_proj_row}")
    ws.row_dimensions[nota_proj_row].height = 60

    bar2 = BarChart()
    bar2.type = "col"
    bar2.grouping = "clustered"
    bar2.title = "SKUs em ruptura: hoje x 30 dias x 60 dias"
    bar2.y_axis.title = "Qtde de SKUs"
    bar2.style = 2
    bar2.add_data(
        Reference(
            ws, min_col=2, max_col=3, min_row=resumo_start, max_row=resumo_start + 3
        ),
        titles_from_data=True,
    )
    bar2.set_categories(
        Reference(ws, min_col=1, min_row=resumo_start + 1, max_row=resumo_start + 3)
    )
    bar2.width, bar2.height = 16, 9
    _sem_grade_vertical(bar2)
    _rotular_dados(bar2, formato="0")
    _colorir_series(bar2, [COR_CRITICO, COR_OK])
    _colocar_grafico(bar2, "hoje_30_60")

    item_tab_row = ctx["item_tab_startrow"] + 3
    ws.cell(
        row=item_tab_row - 1,
        column=1,
        value="Itens ainda não zerados mais urgentes (top 40) — cobertura vs. Média de Saída Diária",
    )
    ws.cell(row=item_tab_row - 1, column=1).font = Font(
        name=FONT_NAME, bold=True, size=12, color=NAVY
    )

    header_item_tab = [ws.cell(row=item_tab_row, column=c).value for c in range(1, 10)]
    for c, _ in enumerate(header_item_tab, start=1):
        cell = ws.cell(row=item_tab_row, column=c)
        cell.font = Font(name=FONT_NAME, bold=True, color=HEADER_FONT_COLOR, size=10)
        cell.fill = PatternFill("solid", fgColor=HEADER_FILL)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    larguras = [14, 42, 14, 14, 20, 20, 22, 22, 40]
    for c, larg in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(
            ws.column_dimensions[get_column_letter(c)].width or 0, larg
        )

    for r in range(item_tab_row + 1, item_tab_row + ctx["n_top_itens_tab"] + 1):
        for c in (3, 4, 5, 6):
            ws.cell(row=r, column=c).number_format = "#,##0.0"
        ws.cell(row=r, column=1).font = Font(name=FONT_NAME, size=10)

    ws = wb["Dashboard"]
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 22
    ws.cell(row=1, column=1).font = Font(
        name=FONT_NAME, bold=True, color=HEADER_FONT_COLOR, size=10
    )
    ws.cell(row=1, column=2).font = Font(
        name=FONT_NAME, bold=True, color=HEADER_FONT_COLOR, size=10
    )
    ws.cell(row=1, column=1).fill = PatternFill("solid", fgColor=HEADER_FILL)
    ws.cell(row=1, column=2).fill = PatternFill("solid", fgColor=HEADER_FILL)

    for r in range(2, ws.max_row + 1):
        ind, val = ws.cell(row=r, column=1), ws.cell(row=r, column=2)
        eh_separador = str(ind.value).startswith("---")
        ind.font = Font(name=FONT_NAME, size=10, bold=eh_separador)
        val.font = Font(name=FONT_NAME, size=10)
        if ind.value and "%" in str(ind.value):
            val.number_format = "0.00%"
        if ind.value and "Valor" in str(ind.value):
            val.number_format = '"R$" #,##0.00'
        if eh_separador:
            ind.fill = PatternFill("solid", fgColor=GREY_FILL)
            val.fill = PatternFill("solid", fgColor=GREY_FILL)

    ws.insert_rows(1, amount=2)
    ws["A1"] = "DASHBOARD DE RUPTURA x AGENDA DE RECEBIMENTO — CD SP"
    ws["A1"].font = Font(name=FONT_NAME, bold=True, size=14, color=NAVY)
    ws.merge_cells("A1:B1")

    chart_start = ws.max_row + 3
    labels_fonte = [
        "Zerados COM chegada agendada",
        "Zerados SEM nenhuma chegada agendada",
        "Zerados sem agenda E sem pedido de compra (crítico)",
    ]
    found = {}
    for r in range(4, ws.max_row + 1):
        key = ws.cell(row=r, column=1).value
        if key in labels_fonte:
            found[key] = ws.cell(row=r, column=2).value

    # As 3 categorias do gráfico precisam ser MUTUAMENTE EXCLUSIVAS para que a
    # soma das barras feche o total de zerados. "Zerados SEM nenhuma chegada
    # agendada" (KPI acima) inclui, como subconjunto, os que também não têm
    # pedido de compra — por isso aqui subtraímos esse subconjunto antes de
    # usar no gráfico, sem alterar o KPI original (que continua mostrando o
    # total real de zerados sem agenda).
    sem_agenda_total = found.get("Zerados SEM nenhuma chegada agendada", 0)
    sem_compra_critico = found.get(
        "Zerados sem agenda E sem pedido de compra (crítico)", 0
    )
    sem_agenda_com_pedido_compra = sem_agenda_total - sem_compra_critico

    labels_grafico = [
        "Zerados COM chegada agendada",
        "Zerados SEM agenda, MAS com pedido de compra em aberto",
        "Zerados sem agenda E sem pedido de compra (crítico)",
    ]
    valores_grafico = [
        found.get("Zerados COM chegada agendada", 0),
        sem_agenda_com_pedido_compra,
        sem_compra_critico,
    ]

    ws.cell(row=chart_start, column=1, value="Situação")
    ws.cell(row=chart_start, column=2, value="Qtde de itens")
    for i, (lab, val) in enumerate(zip(labels_grafico, valores_grafico)):
        ws.cell(row=chart_start + 1 + i, column=1, value=lab)
        ws.cell(row=chart_start + 1 + i, column=2, value=val)

    chart = BarChart()
    chart.type = "col"
    chart.title = "Itens com estoque zerado — cobertura de agenda"
    chart.y_axis.title = "Qtde de SKUs (as 3 barras somam o total de zerados)"
    chart.style = 2
    chart.add_data(
        Reference(ws, min_col=2, min_row=chart_start, max_row=chart_start + 3),
        titles_from_data=True,
    )
    chart.set_categories(
        Reference(ws, min_col=1, min_row=chart_start + 1, max_row=chart_start + 3)
    )
    chart.width, chart.height, chart.legend = 16, 8.5, None
    _sem_grade_vertical(chart)
    _rotular_dados(chart, formato="0")
    cores_dash = [COR_OK, COR_ATENCAO, COR_CRITICO]
    pontos_dash = []
    for i, cor in enumerate(cores_dash):
        dp = DataPoint(idx=i)
        dp.graphicalProperties.solidFill = cor
        pontos_dash.append(dp)
    chart.series[0].data_points = pontos_dash
    _colocar_grafico(chart, "zerados_bar")

    pie_start = chart_start + 6
    total_itens_val = zerados_val = None
    for r in range(4, chart_start):
        k = ws.cell(row=r, column=1).value
        if k == "Total de itens na base (Ruptura)":
            total_itens_val = ws.cell(row=r, column=2).value
        if k == "Itens com Estoque = 0 (zerados)":
            zerados_val = ws.cell(row=r, column=2).value

    ws.cell(row=pie_start, column=1, value="Composição da base")
    ws.cell(row=pie_start, column=2, value="Qtde")
    ws.cell(row=pie_start + 1, column=1, value="Itens zerados (Estoque = 0)")
    ws.cell(row=pie_start + 1, column=2, value=zerados_val or 0)
    ws.cell(row=pie_start + 2, column=1, value="Itens com estoque disponível")
    ws.cell(
        row=pie_start + 2, column=2, value=(total_itens_val or 0) - (zerados_val or 0)
    )

    pie = PieChart()

    pie.title = "Participação dos itens zerados na base total"
    pie.style = 2
    pie.add_data(
        Reference(ws, min_col=2, min_row=pie_start, max_row=pie_start + 2),
        titles_from_data=True,
    )
    pie.set_categories(
        Reference(ws, min_col=1, min_row=pie_start + 1, max_row=pie_start + 2)
    )
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = True
    pie.dataLabels.numFmt = "#,##0"
    pie.width, pie.height = 12, 8.5
    _colorir_fatias_pizza(pie, [COR_CRITICO, COR_OK])
    _colocar_grafico(pie, "pizza_zerados")

    # --- NOVO: Compras agendadas com estoque disponível: necessárias x
    # questionáveis --- (mini-tabela no Dashboard + gráfico)
    nq_start = pie_start + 6
    ws.cell(row=nq_start, column=1, value="Itens com Estoque + Agenda ao mesmo tempo")
    ws.cell(row=nq_start, column=2, value="Qtde")
    ws.cell(row=nq_start + 1, column=1, value="Agenda coerente (giro normal/alto)")
    ws.cell(row=nq_start + 1, column=2, value=ctx["total_agenda_coerente"])
    ws.cell(
        row=nq_start + 2, column=1, value="Compra possivelmente desnecessária (giro baixo)"
    )
    ws.cell(row=nq_start + 2, column=2, value=ctx["n_compra_desnecessaria"])

    pie_nq = PieChart()
    pie_nq.title = "Compras agendadas com estoque disponível: necessárias x questionáveis"
    pie_nq.style = 2
    pie_nq.add_data(
        Reference(ws, min_col=2, min_row=nq_start, max_row=nq_start + 2),
        titles_from_data=True,
    )
    pie_nq.set_categories(
        Reference(ws, min_col=1, min_row=nq_start + 1, max_row=nq_start + 2)
    )
    pie_nq.dataLabels = DataLabelList()
    pie_nq.dataLabels.showPercent = True
    pie_nq.dataLabels.showVal = True
    pie_nq.dataLabels.numFmt = "#,##0"
    pie_nq.width, pie_nq.height = 14, 8.5
    _colorir_fatias_pizza(pie_nq, [COR_OK, COR_ATENCAO])
    _colocar_grafico(pie_nq, "necessaria_questionavel")

    # --- NOVO: Top 10 — maiores valores parados em ruptura (R$) ---
    # Zerados_x_Agenda já vem ordenado por "Valor em Ruptura (R$)" decrescente.
    ws_zag = wb["Zerados_x_Agenda"]
    if ws_zag.max_row >= 2:
        header_zag = [ws_zag.cell(row=1, column=c).value for c in range(1, ws_zag.max_column + 1)]
        col_idx_zag = {n: i + 1 for i, n in enumerate(header_zag) if n}
        top_n_zag = min(10, ws_zag.max_row - 1)
        if "Valor em Ruptura (R$)" in col_idx_zag and "Cód. + Descrição" in col_idx_zag:
            chart_valor = BarChart()
            chart_valor.type = "bar"
            chart_valor.style = 2
            chart_valor.title = "Top 10 — maiores valores parados em ruptura (R$)"
            chart_valor.y_axis.title = "Valor em ruptura (R$)"
            chart_valor.add_data(
                Reference(
                    ws_zag,
                    min_col=col_idx_zag["Valor em Ruptura (R$)"],
                    min_row=1,
                    max_row=top_n_zag + 1,
                ),
                titles_from_data=True,
            )
            chart_valor.set_categories(
                Reference(
                    ws_zag,
                    min_col=col_idx_zag["Cód. + Descrição"],
                    min_row=2,
                    max_row=top_n_zag + 1,
                )
            )
            chart_valor.width, chart_valor.height, chart_valor.legend = 24, 12, None
            _sem_grade_vertical(chart_valor)
            _rotular_dados(chart_valor, formato="#,##0")
            _colorir_series(chart_valor, [COR_CRITICO])
            _colocar_grafico(chart_valor, "top10_valor_ruptura")

    # --- NOVO: Top 10 — Estoque atual x Qtd. Agendada (maiores estoques),
    # sem filtro de giro — visão geral dos maiores volumes parados que ainda
    # assim têm mais carga chegando. Tabela de apoio fica fora da área
    # visível principal (colunas AZ em diante) só para alimentar o gráfico.
    top10_est = ctx["top10_maiores_estoques"]
    col_apoio = 55  # coluna "BC" em diante — fora da área de leitura normal
    ws.cell(row=2, column=col_apoio, value="Cód. + Descrição")
    ws.cell(row=2, column=col_apoio + 1, value="Estoque Atual")
    ws.cell(row=2, column=col_apoio + 2, value="Qtd. Agendada")
    for i, (_, row_) in enumerate(top10_est.iterrows()):
        r = 3 + i
        ws.cell(
            row=r,
            column=col_apoio,
            value=row_.get("Cód. + Descrição")
            or _rotulo_cod_desc(row_["Cód. Produto"], row_["Descrição"]),
        )
        ws.cell(row=r, column=col_apoio + 1, value=row_["Estoque"])
        ws.cell(row=r, column=col_apoio + 2, value=row_["Qtd. Agendada"])
    n_top10_est = len(top10_est)
    if n_top10_est:
        chart_estoques = BarChart()
        chart_estoques.type = "bar"
        chart_estoques.style = 2
        chart_estoques.title = "Top 10 — Estoque atual x Qtd. Agendada (maiores estoques)"
        chart_estoques.add_data(
            Reference(
                ws,
                min_col=col_apoio + 1,
                max_col=col_apoio + 2,
                min_row=2,
                max_row=n_top10_est + 2,
            ),
            titles_from_data=True,
        )
        chart_estoques.set_categories(
            Reference(ws, min_col=col_apoio, min_row=3, max_row=n_top10_est + 2)
        )
        chart_estoques.width, chart_estoques.height = 24, 12
        _sem_grade_vertical(chart_estoques)
        _rotular_dados(chart_estoques, formato="#,##0")
        _colorir_series(chart_estoques, [COR_NEUTRO, COR_SECUNDARIA])
        _colocar_grafico(chart_estoques, "top10_maiores_estoques")

    # --- Top 10 de Compra_Desnecessaria (já ordenada por Estoque Atual
    # decrescente), unificado num ÚNICO gráfico de barras horizontais: a
    # barra é o valor estimado da compra extra (métrica mais acionável, em
    # R$), e o rótulo de cada item traz os demais detalhes (estoque, agenda
    # e dias de cobertura) — assim tudo aparece em um só lugar sem misturar
    # unidades diferentes (R$, unidades e dias) nas mesmas barras.
    if "Compra_Desnecessaria" in wb.sheetnames:
        ws_cd = wb["Compra_Desnecessaria"]
        n_cd_top = min(10, ctx["n_compra_desnecessaria"])
        if n_cd_top > 0:
            header_cd2 = [
                ws_cd.cell(row=1, column=c).value for c in range(1, ws_cd.max_column + 1)
            ]
            col_idx_cd2 = {n: i + 1 for i, n in enumerate(header_cd2) if n}

            # Área de staging para o gráfico de 3 séries (colunas fora da área visível).
            # Estrutura: [Rótulo Simples] [Estoque Atual] [Qtd. Agendada] [Valor Est.]
            # O rótulo exibe apenas: Nº do item — Descrição | X dias de cobertura
            # (Estoque e Agenda aparecem como barras, não mais no rótulo).
            col_rotulo = ws_cd.max_column + 2          # label simplificado
            col_est    = col_rotulo + 1                # Estoque Atual (un)
            col_age    = col_rotulo + 2                # Qtd. Agendada (un)
            col_val    = col_rotulo + 3                # Valor Estimado (R$)

            # Cabeçalhos do staging (linha 1 → serão usados como nomes de série)
            ws_cd.cell(row=1, column=col_rotulo, value="Item — Cobertura")
            ws_cd.cell(row=1, column=col_est,    value="Estoque Atual (un)")
            ws_cd.cell(row=1, column=col_age,    value="Qtd. Agendada (un)")
            ws_cd.cell(row=1, column=col_val,    value="Valor Est. Compra Extra (R$)")

            for r in range(2, n_cd_top + 2):
                cod_desc = ws_cd.cell(row=r, column=col_idx_cd2["Cód. + Descrição"]).value or ""
                dias_v   = ws_cd.cell(row=r, column=col_idx_cd2["Dias de Cobertura Atual (giro)"]).value or 0
                estoque_v = ws_cd.cell(row=r, column=col_idx_cd2["Estoque Atual"]).value or 0
                agenda_v  = ws_cd.cell(row=r, column=col_idx_cd2["Qtd. Agendada"]).value or 0
                valor_v   = ws_cd.cell(row=r, column=col_idx_cd2["Valor Estimado da Compra Extra (R$)"]).value or 0

                # Rótulo simplificado: só Nº item — Descrição | dias de cobertura
                # (evita poluição visual — estoque e agenda já aparecem como barras)
                rotulo_simples = (
                    f"{cod_desc}  |  {float(dias_v):,.0f} dias"
                ).replace(",", ".")
                ws_cd.cell(row=r, column=col_rotulo, value=rotulo_simples)
                ws_cd.cell(row=r, column=col_est,    value=estoque_v)
                ws_cd.cell(row=r, column=col_age,    value=agenda_v)
                ws_cd.cell(row=r, column=col_val,    value=valor_v)

            # Gráfico combo — colunas verticais + linha em eixo secundário:
            #   ▮ Estoque Atual   (cinza DPSP — quanto já está no CD)
            #   ▮ Qtd. Agendada   (azul DPSP  — quanto está a caminho)
            #   ── Valor Est. R$  (vermelho atenuado — impacto financeiro estimado)
            # Estoque/Agenda ficam em unidades no eixo primário; Valor (R$) tem
            # escala muito diferente, por isso vai como linha no eixo secundário
            # (mesmo padrão usado no gráfico "curva" de Projecao_30_60_Dias).
            cats_ref_cd = Reference(
                ws_cd,
                min_col=col_rotulo,
                min_row=2,
                max_row=n_cd_top + 1,
            )

            chart_cd = BarChart()
            chart_cd.type     = "col"        # colunas verticais
            chart_cd.grouping = "clustered"  # paralelas (não empilhadas)
            chart_cd.style    = 2
            chart_cd.title    = (
                "Top 10 — Compra provavelmente desnecessária: "
                "Estoque | Agenda | Valor estimado  (por item)"
            )
            chart_cd.y_axis.title = "Qtde (un)  /  Valor estimado (R$)"
            chart_cd.add_data(
                Reference(
                    ws_cd,
                    min_col=col_est,
                    max_col=col_age,
                    min_row=1,
                    max_row=n_cd_top + 1,
                ),
                titles_from_data=True,
            )
            chart_cd.set_categories(cats_ref_cd)
            # Largura/altura mantidas; legenda embaixo para não sumir as barras
            chart_cd.width  = 24
            chart_cd.height = 14
            chart_cd.legend = Legend(legendPos="b")
            _sem_grade_vertical(chart_cd)
            _rotular_dados(chart_cd, formato="#,##0")
            # Cores: cinza (estoque) | azul (agenda)
            _colorir_series(chart_cd, [COR_NEUTRO, COR_OK])

            # Série "Valor Est. Compra Extra (R$)" como linha no eixo secundário
            line_valor_cd = LineChart()
            line_valor_cd.add_data(
                Reference(
                    ws_cd,
                    min_col=col_val,
                    max_col=col_val,
                    min_row=1,
                    max_row=n_cd_top + 1,
                ),
                titles_from_data=True,
            )
            line_valor_cd.set_categories(cats_ref_cd)
            line_valor_cd.y_axis.axId = 200
            line_valor_cd.y_axis.crosses = "max"
            for s in line_valor_cd.series:
                s.smooth = True
                s.marker.symbol = "none"
                s.graphicalProperties.line.solidFill = COR_ATENCAO
                s.graphicalProperties.line.width = 22000

            chart_cd += line_valor_cd
            _colocar_grafico(chart_cd, "top10_compra_desnecessaria_unificado")

    # --- NOVO: Top 10 itens da Compra_Desnecessaria cruzados com o arquivo
    # 420_NORMA — mostra quais itens têm mais estoque parado ocupando espaço
    # físico no armazém (quantidade real + número de posições ocupadas).
    if "Posicoes_Estoque_420Norma" in wb.sheetnames and ctx["total_itens_cruzados_420"] > 0:
        ws_pos = wb["Posicoes_Estoque_420Norma"]
        _estilizar_tabela(ws_pos, moeda_cols=[], data_cols=["Validade"])
        df_cruzado = ctx["compra_desnecessaria_com_posicao"].head(10)
        col_apoio2 = 60
        ws_pos.cell(row=2, column=col_apoio2, value="Item (posições ocupadas)")
        ws_pos.cell(row=2, column=col_apoio2 + 1, value="Qtde total no 420_NORMA")
        for i, (_, row_) in enumerate(df_cruzado.iterrows()):
            r = 3 + i
            rotulo_item = (
                f"{_rotulo_cod_desc(row_['Cód. Produto'], row_['Descrição'])} "
                f"({int(row_['N_Posicoes'])} posições)"
            )
            ws_pos.cell(row=r, column=col_apoio2, value=rotulo_item)
            ws_pos.cell(row=r, column=col_apoio2 + 1, value=row_["Qtde_Total_420"])
        n_cruzado = len(df_cruzado)

        chart_pos = BarChart()
        chart_pos.type = "bar"
        chart_pos.style = 2
        chart_pos.title = (
            "Top 10 — Itens de compra desnecessária com maior ocupação física "
            "no armazém (cruzado com 420_NORMA)"
        )
        chart_pos.y_axis.title = "Quantidade total ocupando o armazém (todas as posições)"
        chart_pos.add_data(
            Reference(
                ws_pos,
                min_col=col_apoio2 + 1,
                min_row=2,
                max_row=n_cruzado + 2,
            ),
            titles_from_data=True,
        )
        chart_pos.set_categories(
            Reference(ws_pos, min_col=col_apoio2, min_row=3, max_row=n_cruzado + 2)
        )
        chart_pos.legend = None
        _sem_grade_vertical(chart_pos)
        _rotular_dados(chart_pos, formato="#,##0")
        _colorir_series(chart_pos, [COR_CRITICO])
        _colocar_grafico(chart_pos, "posicoes_420norma")

    if "Notas_Metodológicas" in wb.sheetnames:
        del wb["Notas_Metodológicas"]
    notas = wb.create_sheet("Notas_Metodológicas")
    notas.sheet_view.showGridLines = False
    notas.column_dimensions["A"].width = 100
    linhas = [
        (
            "TITULO",
            "Notas metodológicas — Análise de Ruptura x Agenda de Recebimento (CD-SP)",
        ),
        ("", ""),
        ("SUB", "1. Cruzamento realizado"),
        (
            "",
            "Chave: 'Material' (Agenda_CD_SP.xlsx) = 'Cód. Produto' (Ruptura_qlikview.xlsx). Um mesmo Material com várias cargas agendadas é consolidado (quantidade somada, datas e fornecedores agrupados).",
        ),
        ("", ""),
        ("SUB", "2. Definição de 'zerado' usada nas abas principais"),
        (
            "",
            "Estoque (coluna M do Ruptura) igual a 0 — conforme solicitado originalmente.",
        ),
        ("", ""),
        ("SUB", "3. Alerta — achado que muda o diagnóstico"),
        (
            "",
            "O QlikView já calcula uma flag oficial de ruptura ('Ruptura': C/Ruptura x S/Ruptura) baseada em dias de cobertura (Estoque ÷ média de saída diária), não apenas em Estoque=0.",
        ),
        (
            "",
            f"{ctx['total_alerta_ampliado']:,}".replace(",", ".")
            + " itens estão sinalizados como 'C/ Ruptura' pelo sistema, contra "
            + f"{ctx['total_zerados']:,}".replace(",", ".")
            + " com Estoque literalmente zero. "
            + f"{ctx['total_alerta_estoque_positivo']:,}".replace(",", ".")
            + " itens têm estoque > 0 mas cobrem menos de 1 dia de venda — já em ruptura operacional.",
        ),
        (
            "",
            "Recomendação: usar a flag 'Ruptura' oficial como critério primário de monitoramento diário, não apenas Estoque=0.",
        ),
        (
            "",
            "Os números acima são recalculados a cada execução deste script a partir dos arquivos de origem do dia — não são valores fixos de uma análise pontual.",
        ),
        ("", ""),
        ("SUB", "4. Reprodutibilidade diária"),
        (
            "",
            "Basta substituir os dois arquivos de origem (mesmos nomes e colunas) e rodar 'gerar_dashboard_diario.py' novamente.",
        ),
        ("", ""),
        ("SUB", "5. Projeção de ruptura em 30 e 60 dias (novo)"),
        (
            "",
            "Cálculo: Dias até zerar = Estoque disponível ÷ Média de saída diária (coluna 'Média Qtd' do Ruptura). 'Estoque disponível' SEM Agenda = Estoque atual. COM Agenda = Estoque atual + Qtd. Agendada Total.",
        ),
        (
            "",
            "A quantidade agendada é somada como já disponível, sem considerar a data exata de chegada — decisão do usuário, já que o arquivo Agenda_CD_SP.xlsx cobre apenas ~1 dia de horizonte.",
        ),
        (
            "",
            "19 itens aparecem na Agenda mas não têm nenhum registro no relatório Ruptura. Para esses, a Média de saída foi ESTIMADA usando a mediana da Média Qtd de toda a carteira.",
        ),
        (
            "",
            "Itens sem nenhum histórico de saída (Média Qtd = 0) e com estoque > 0 são tratados como 'sem risco projetável'.",
        ),
        (
            "",
            "Limitação importante: a Média Qtd usada é uma média histórica. Trate a projeção como uma estimativa direcional, não como previsão exata.",
        ),
        ("", ""),
        ("SUB", "6. Por que a Agenda muda pouco entre 30 e 60 dias"),
        (
            "",
            "O arquivo Agenda_CD_SP.xlsx normalmente traz só a programação de carga de 1 dia. Isso significa que a 'Qtd. Agendada' usada na projeção reflete o que está confirmado para chegar em breve.",
        ),
        ("", ""),
        ("SUB", "7. Códigos de produto duplicados no arquivo de Ruptura"),
        (
            "",
            (
                f"{ctx['n_codigos_duplicados']} código(s) de produto ({ctx['n_linhas_duplicadas']} linhas) aparecem mais de uma vez no arquivo de Ruptura desta execução. Veja a aba 'Diagnostico_Duplicados'."
                if ctx["n_codigos_duplicados"]
                else "Nenhum código de produto duplicado foi encontrado no arquivo de Ruptura desta execução."
            ),
        ),
        (
            "",
            "Quando um código se repete, cada linha tem seu próprio Estoque e Média — e 'Dias Restantes de Estoque' é calculado linha a linha. Ao filtrar o código no Excel, é fácil comparar por engano o Estoque de uma linha com o resultado calculado a partir de outra linha do mesmo código, o que parece um erro de fórmula mas é, na verdade, dado duplicado na origem.",
        ),
        (
            "",
            "O script não soma nem descarta automaticamente essas linhas — a decisão de tratá-las como o mesmo SKU (somando Estoque) ou como registros legitimamente distintos deve ser validada com quem gera o export do QlikView antes de qualquer ajuste no cálculo.",
        ),
        ("", ""),
        ("SUB", "8. Aba 'Compra_Desnecessaria' — critério usado"),
        (
            "",
            f"Um item entra nesta aba quando, ao mesmo tempo: (1) Estoque atual > 0; (2) existe carga agendada (Qtd. Agendada > 0); e (3) o giro é baixo, definido como cobertura atual ≥ {LIMIAR_DIAS_GIRO_BAIXO} dias de venda, OU o item não tem nenhum histórico de saída (Média Qtd = 0) mesmo tendo estoque.",
        ),
        (
            "",
            "A lógica é: se o item tem MUITO estoque mas o giro é ALTO (poucos dias de cobertura), a compra faz sentido e ele NÃO entra na lista — o risco real é estoque alto + giro baixo + comprando mais mesmo assim.",
        ),
        (
            "",
            f"Nesta execução: {ctx['n_compra_desnecessaria']} item(ns) se enquadram no critério, somando um valor estimado de "
            + f"R$ {ctx['valor_compra_desnecessaria']:,.2f}".replace(",", "#").replace(".", ",").replace("#", ".")
            + " em compra questionável.",
        ),
        (
            "",
            "'Valor Estimado da Compra Extra (R$)' é uma ESTIMATIVA: preço unitário = Média Valor de saída (R$/dia) ÷ Média Qtd de saída (un/dia), multiplicado pela Qtd. Agendada. Não é o custo de compra oficial — use como ordem de grandeza para priorizar, não como valor contábil.",
        ),
        (
            "",
            f"O limiar de {LIMIAR_DIAS_GIRO_BAIXO} dias é um parâmetro do script (constante LIMIAR_DIAS_GIRO_BAIXO) e pode ser ajustado conforme a política de estoque do CD.",
        ),
        ("", ""),
        ("SUB", "9. Cruzamento com posições físicas no armazém (420_NORMA)"),
        (
            "",
            (
                f"O arquivo 420_NORMA foi carregado e cruzado com a aba Compra_Desnecessaria: {ctx['total_itens_cruzados_420']} item(ns) em comum encontrados nesta execução."
                if ctx["norma_arquivo_ok"]
                else "O arquivo 420_NORMA não foi encontrado/carregado nesta execução — a aba e o gráfico de posições no armazém não foram gerados. Coloque o arquivo '420_NORMA*.csv' na mesma pasta do script (ou passe como 4º argumento) para ativar essa análise."
            ),
        ),
        (
            "",
            "Quantidade por posição = coluna 'UM' × coluna 'UGE1/UM' do arquivo 420_NORMA (unidades por posição × unidades por embalagem). Um mesmo código pode aparecer em várias posições — a aba Posicoes_Estoque_420Norma lista cada posição individualmente; o gráfico soma o total por item.",
        ),
        (
            "",
            "Só entram nesse cruzamento os itens que aparecem NOS DOIS arquivos ao mesmo tempo (Compra_Desnecessaria E 420_NORMA). Um item da Compra_Desnecessaria que não aparece aqui simplesmente não estava no arquivo 420_NORMA desta execução — não significa que não tenha estoque físico.",
        ),
    ]
    r = 1
    for tipo, texto in linhas:
        cell = notas.cell(row=r, column=1, value=texto)
        if tipo == "TITULO":
            cell.font = Font(name=FONT_NAME, bold=True, size=13)
        elif tipo == "SUB":
            cell.font = Font(name=FONT_NAME, bold=True, size=11, color=NAVY)
            cell.fill = PatternFill("solid", fgColor=GREY_FILL)
        else:
            cell.font = Font(name=FONT_NAME, size=10)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        notas.row_dimensions[r].height = 30 if texto and len(texto) > 90 else 18
        r += 1

    order = [
        "GRAFICOS",
        "Dashboard",
        "Zerados_x_Agenda",
        "Criticos_Sem_Agenda",
        "Zerados_Com_Agenda",
        "Alerta_Ampliado_Cobertura",
        "Estoque_x_Agenda",
        "Compra_Desnecessaria",
        "Posicoes_Estoque_420Norma",
        "Projecao_30_60_Dias",
        "Detalhe_Projecao_60d",
        "Diagnostico_Duplicados",
        "Notas_Metodológicas",
    ]
    order = [name for name in order if name in wb.sheetnames]
    wb._sheets = [wb[name] for name in order]
    wb.active = 0
    wb.save(path)
    _forcar_numfmt_rotulos_grafico(path)
    print(f"\nPronto! Dashboard salvo em: {path}")


def _forcar_numfmt_rotulos_grafico(path):
    """Corrige uma limitação do openpyxl: ao definir um formato numérico
    customizado para rótulos de dados (ex.: '0' para esconder casas
    decimais), ele grava <numFmt formatCode="0"/> no XML SEM o atributo
    sourceLinked="0". Sem esse atributo, o Excel/LibreOffice ignora o
    formatCode e usa o formato "geral" da célula de origem — por isso os
    rótulos apareciam com várias casas decimais (ex. "1,110") mesmo com o
    formato "0" definido, sobrepondo os nomes dos itens nos gráficos de
    barra horizontal. Este pós-processamento reabre o .xlsx já salvo e
    injeta sourceLinked="0" em todo numFmt de rótulo de dados que não o
    tiver, sem precisar reprocessar nenhum dado."""
    import re
    import shutil
    import zipfile

    tmp_path = path + ".tmp"
    padrao = re.compile(r'<numFmt formatCode="([^"]*)"\s*/>')
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(
        tmp_path, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            conteudo = zin.read(item.filename)
            if item.filename.startswith("xl/charts/chart") and item.filename.endswith(
                ".xml"
            ):
                texto = conteudo.decode("utf-8")
                texto = padrao.sub(
                    r'<numFmt formatCode="\1" sourceLinked="0"/>', texto
                )
                conteudo = texto.encode("utf-8")
            zout.writestr(item, conteudo)
    shutil.move(tmp_path, path)


if __name__ == "__main__":
    main()
