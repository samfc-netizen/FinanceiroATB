# DRE.py
# Streamlit multi-página: DRE + DFC
# Modelo conforme especificação do Samuel:
#  - Filtros laterais: Ano e Meses
#  - DRE:
#       RECEITA (aba RECEITA: mês, ano, receita)
#       CMV (aba CMV: mês, ano, CMV)
#       MARGEM DE CONTRIBUIÇÃO = Receita - CMV
#       DESPESA C/ PESSOAL / FINANCEIRA / IMPOSTOS / OPERACIONAL (aba BASE DE DADOS: TIPO, Valor total, DATA)
#       RESULTADO OPERACIONAL = Margem - (despesas)
#       Tabela: JAN | JAN% | FEV | FEV% ... + ACUMULADO + %ACUMULADO (percentual sobre receita)
#       Drill: Linha -> Plano de contas (sintetizado) -> Descrição + Pizza + Cards (média e % sobre Receita)
#  - DFC:
#       RECEBIMENTOS (aba RECEBIMENTO: mês, ano, recebimento)
#       Despesas (BASE DE DADOS: mesmos campos) + EMPRÉSTIMOS + FORNECEDOR
#       RESULTADO CAIXA = Recebimentos - (todas as despesas abaixo, incluindo Empréstimos)
#       RESULTADO ANTES DOS EMPRÉSTIMOS = Resultado Caixa + Empréstimos
#       Drill igual ao DRE (denominador = Recebimentos)
#
# Requisitos:
#   python -m pip install streamlit pandas openpyxl plotly
#
# Rodar (na pasta onde está o Excel):
#   python -m streamlit run DRE.py
#
# Observação importante:
# - Este app procura um Excel na mesma pasta. Preferência:
#     1) "DRE E FLUXO DE CAIXA.xlsx"
#     2) "DRE E DFC GERAL.xlsx"
#     3) o .xlsx mais recente da pasta
#
# - Para evitar erro por nomes com espaços/acentos, o app resolve abas e colunas por normalização.
#
from __future__ import annotations

import glob
import os
import re
import unicodedata
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st


# ----------------------------
# Constantes / formatação
# ----------------------------

MESES_PT = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
MES_NUM_TO_PT = {1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
                 7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"}
MES_PT_TO_NUM = {v: k for k, v in MES_NUM_TO_PT.items()}


def _norm_txt(s: object) -> str:
    """Normaliza texto (remove acentos, lower, normaliza espaços)."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    t = str(s)
    t = t.replace("\u00a0", " ").replace("–", "-").replace("—", "-")
    t = unicodedata.normalize("NFKD", t)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    t = re.sub(r"\s+", " ", t).strip().lower()
    return t


def _norm_sheet_name(s: str) -> str:
    """Normaliza nome de aba ignorando espaços/acentos."""
    return _norm_txt(s).replace(" ", "")


def to_num(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip()
    if s == "":
        return 0.0
    s = s.replace("\u00a0", " ").replace("R$", "").strip()
    # pt-BR 1.234,56 -> 1234.56
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def format_brl(x) -> str:
    try:
        return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"


def fmt_pct(x) -> str:
    try:
        return f"{float(x):,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00%"


def parse_mes(v) -> Optional[int]:
    """Aceita 1..12, '01', 'JAN', 'Janeiro'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip().upper()
    if s.isdigit():
        m = int(s)
        return m if 1 <= m <= 12 else None
    mapa = {
        "JANEIRO": 1, "JAN": 1,
        "FEVEREIRO": 2, "FEV": 2,
        "MARCO": 3, "MARÇO": 3, "MAR": 3,
        "ABRIL": 4, "ABR": 4,
        "MAIO": 5, "MAI": 5,
        "JUNHO": 6, "JUN": 6,
        "JULHO": 7, "JUL": 7,
        "AGOSTO": 8, "AGO": 8,
        "SETEMBRO": 9, "SET": 9,
        "OUTUBRO": 10, "OUT": 10,
        "NOVEMBRO": 11, "NOV": 11,
        "DEZEMBRO": 12, "DEZ": 12,
    }
    return mapa.get(s)


def sintetizar_plano_contas(nome: object) -> str:
    """
    Sintetiza 'Plano de contas' removendo sufixos comuns e normalizando espaços.
    """
    if nome is None or (isinstance(nome, float) and pd.isna(nome)):
        return "—"
    s = str(nome).strip()
    # remove sufixo do tipo "(3 - DESPESAS)" se houver
    s = re.sub(r"\s*\(\s*\d+\s*-\s*despesas\s*\)\s*$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s if s else "—"


# ----------------------------
# Localização dos dois arquivos + cache
# ----------------------------

PLANO_FILE_HINTS = [
    "PLANO DE CONTAS.xlsx", "PLANO DE CONTAS.xlsm", "PLANO DE CONTAS.xls",
]
DADOS_EXTERNOS_FILE_HINTS = [
    "DADOS EXTERNOS.xlsx", "DADOS EXTERNOS.xlsm", "DADOS EXTERNOS.xls",
]


def _excel_files() -> List[str]:
    files: List[str] = []
    for pat in ["*.xlsx", "*.xlsm", "*.xls"]:
        files.extend(glob.glob(pat))
    return [f for f in files if os.path.isfile(f) and not os.path.basename(f).startswith("~$")]


def _find_named_excel(hints: List[str], required_tokens: List[str]) -> Optional[str]:
    for fn in hints:
        if os.path.exists(fn):
            return fn

    candidates = []
    for f in _excel_files():
        norm = _norm_txt(os.path.splitext(os.path.basename(f))[0])
        if all(tok in norm for tok in required_tokens):
            candidates.append(f)
    if not candidates:
        return None
    candidates.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return candidates[0]


def find_plano_contas_file() -> Optional[str]:
    return _find_named_excel(PLANO_FILE_HINTS, ["plano", "contas"])


def find_dados_externos_file() -> Optional[str]:
    return _find_named_excel(DADOS_EXTERNOS_FILE_HINTS, ["dados", "externos"])


def excel_signature(path: str) -> Tuple[int, int]:
    stt = os.stat(path)
    return (stt.st_mtime_ns, stt.st_size)


@st.cache_resource(show_spinner=False)
def get_excel_file(excel_path: str, sig: Tuple[int, int]) -> pd.ExcelFile:
    return pd.ExcelFile(excel_path)


def resolve_sheet(xls: pd.ExcelFile, desired: str) -> Optional[str]:
    want = _norm_sheet_name(desired)
    mapping = {_norm_sheet_name(s): s for s in xls.sheet_names}
    if want in mapping:
        return mapping[want]
    for k, real in mapping.items():
        if want in k:
            return real
    return None


def _detect_header_row(excel_path: str, sheet_name: str, required_groups: List[List[str]], max_rows: int = 40) -> int:
    raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=max_rows)
    for idx, row in raw.iterrows():
        values = {_norm_txt(v) for v in row.tolist() if _norm_txt(v)}
        ok = True
        for group in required_groups:
            if not any(_norm_txt(c) in values for c in group):
                ok = False
                break
        if ok:
            return int(idx)
    return 0


@st.cache_data(show_spinner=False)
def read_sheet(excel_path: str, desired_sheet: str, sig: Tuple[int, int], auto_header: bool = False) -> Optional[pd.DataFrame]:
    xls = get_excel_file(excel_path, sig)
    real = resolve_sheet(xls, desired_sheet)
    if real is None:
        return None
    try:
        header = 0
        if auto_header:
            header = _detect_header_row(
                excel_path,
                real,
                [["DATA", "Data de vencimento", "Data de competência", "Data de confirmação"],
                 ["Plano de contas"],
                 ["Valor total", "Valor"]],
            )
        df = pd.read_excel(excel_path, sheet_name=real, header=header)
    except Exception:
        return None
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df


def _find_col(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    # match direto
    for cand in candidates:
        if cand in df.columns:
            return cand
    # match normalizado
    norm_map = {_norm_txt(c): c for c in df.columns}
    for cand in candidates:
        cc = norm_map.get(_norm_txt(cand))
        if cc is not None:
            return cc
    return None


@st.cache_data(show_spinner=False)
def prep_base(
    plano_path: str,
    plano_sig: Tuple[int, int],
    dados_path: str,
    dados_sig: Tuple[int, int],
) -> Optional[pd.DataFrame]:
    """Lê os lançamentos do arquivo PLANO DE CONTAS e acrescenta TIPO via aba TIPO."""
    xls_plano = get_excel_file(plano_path, plano_sig)
    # Usa a primeira aba; o nome do arquivo é a referência principal.
    real_sheet = xls_plano.sheet_names[0] if xls_plano.sheet_names else None
    if real_sheet is None:
        return None

    try:
        header = _detect_header_row(
            plano_path,
            real_sheet,
            [["Data de confirmação"],
             ["Descrição"],
             ["Plano de contas"],
             ["Valor total"]],
        )
        df = pd.read_excel(plano_path, sheet_name=real_sheet, header=header)
    except Exception:
        return None

    df.columns = [str(c).strip() for c in df.columns]
    base = df.dropna(how="all").copy()

    col_val = _find_col(base, ["Valor total", "VALOR TOTAL", "Valor Total", "VALOR", "Valor"])
    # Regra definida pelo usuário: usar exclusivamente Data de confirmação como data do lançamento.
    col_dt = _find_col(base, ["Data de confirmação", "Data de Confirmação", "DATA DE CONFIRMAÇÃO", "Data confirmacao", "DATA CONFIRMACAO"])
    col_plano = _find_col(base, ["Plano de contas", "PLANO DE CONTAS", "Plano de Contas"])
    col_desc = _find_col(base, ["Descrição", "DESCRIÇÃO", "Descricao", "DESCRICAO"])

    if col_val is None or col_dt is None or col_plano is None:
        return None

    # A aba TIPO pode vir com cabeçalho ou sem cabeçalho.
    # No modelo real enviado pelo usuário, ela não possui cabeçalho:
    # coluna A = Plano de contas | coluna B = TIPO.
    try:
        tipo_real_sheet = resolve_sheet(get_excel_file(dados_path, dados_sig), "TIPO")
        if tipo_real_sheet is None:
            return None

        tipo_df = pd.read_excel(dados_path, sheet_name=tipo_real_sheet)
        tipo_df.columns = [str(c).strip() for c in tipo_df.columns]

        tipo_plano_col = _find_col(tipo_df, ["Plano de contas", "PLANO DE CONTAS", "Plano de Contas"])
        tipo_col = _find_col(tipo_df, ["TIPO", "Tipo"])

        # Fallback para aba sem cabeçalho: relê com header=None e usa as 2 primeiras colunas.
        if tipo_plano_col is None or tipo_col is None:
            tipo_df = pd.read_excel(dados_path, sheet_name=tipo_real_sheet, header=None)
            tipo_df = tipo_df.dropna(how="all").copy()
            if tipo_df.shape[1] < 2:
                return None
            tipo_df = tipo_df.iloc[:, :2].copy()
            tipo_df.columns = ["PLANO DE CONTAS", "TIPO"]
            tipo_plano_col = "PLANO DE CONTAS"
            tipo_col = "TIPO"
    except Exception:
        return None

    mapa_tipo = tipo_df[[tipo_plano_col, tipo_col]].dropna(subset=[tipo_plano_col]).copy()
    mapa_tipo["_plano_key"] = mapa_tipo[tipo_plano_col].apply(_norm_txt)
    mapa_tipo["_tipo_map"] = mapa_tipo[tipo_col].astype(str).str.strip()
    mapa_tipo = mapa_tipo[mapa_tipo["_plano_key"] != ""].drop_duplicates("_plano_key", keep="last")

    base["_dt"] = pd.to_datetime(base[col_dt], errors="coerce", dayfirst=True)
    base["_ano"] = base["_dt"].dt.year
    base["_mes"] = base["_dt"].dt.month
    base["_v"] = base[col_val].apply(to_num)
    base["_plano"] = base[col_plano].fillna("—").astype(str).str.strip()
    base["_plano_sint"] = base["_plano"].apply(sintetizar_plano_contas)
    base["_plano_key"] = base["_plano"].apply(_norm_txt)
    base["_desc"] = base[col_desc].fillna("—").astype(str).str.strip() if col_desc is not None else "—"

    base = base.merge(mapa_tipo[["_plano_key", "_tipo_map"]], on="_plano_key", how="left")
    base["_tipo"] = base["_tipo_map"].fillna("").astype(str).str.strip()
    base["_tipo_norm"] = base["_tipo"].apply(_norm_txt)
    base["_tipo_pendente"] = base["_tipo_norm"] == ""
    return base


def planos_sem_tipo(base: pd.DataFrame) -> pd.DataFrame:
    if base is None or base.empty or "_tipo_pendente" not in base.columns:
        return pd.DataFrame(columns=["PLANO DE CONTAS", "TIPO"])
    pend = base[base["_tipo_pendente"] & (base["_plano_key"] != "")].copy()
    if pend.empty:
        return pd.DataFrame(columns=["PLANO DE CONTAS", "TIPO"])
    out = pend[["_plano"]].drop_duplicates().sort_values("_plano")
    out = out.rename(columns={"_plano": "PLANO DE CONTAS"})
    out["TIPO"] = ""
    return out.reset_index(drop=True)


def agg_mes_ano_val(df: pd.DataFrame, value_candidates: List[str], ano_ref: int) -> Optional[Dict[int, float]]:
    """Agrega por mês usando colunas mês/ano + valor."""
    col_ano = _find_col(df, ["ANO", "Ano"])
    col_mes = _find_col(df, ["MÊS", "MES", "MÊS ", "MES "])
    col_val = _find_col(df, value_candidates)

    if col_mes is None or col_val is None:
        return None

    tmp = df.copy()
    if col_ano is not None:
        tmp["_ano"] = pd.to_numeric(tmp[col_ano], errors="coerce").astype("Int64")
        tmp = tmp[tmp["_ano"] == int(ano_ref)]

    tmp["_mes"] = tmp[col_mes].apply(parse_mes)
    tmp = tmp[tmp["_mes"].notna()].copy()

    tmp["_v"] = tmp[col_val].apply(to_num)
    grp = tmp.groupby("_mes")["_v"].sum()
    return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}


def sum_base_by_month(base: pd.DataFrame, ano: int, tipo_label: str) -> Dict[int, float]:
    """Soma BASE DE DADOS por mês filtrando pelo TIPO."""
    key_norm = _norm_txt(tipo_label)
    tmp = base[base["_ano"] == int(ano)].copy()
    mask = (tmp["_tipo_norm"] == key_norm) | (tmp["_tipo_norm"].str.contains(key_norm, na=False))
    tmp = tmp[mask].copy()
    grp = tmp.groupby("_mes")["_v"].sum()
    return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}


def filter_base(base: pd.DataFrame, ano: int, meses: List[int], tipo_label: str) -> pd.DataFrame:
    key_norm = _norm_txt(tipo_label)
    tmp = base[(base["_ano"] == int(ano)) & (base["_mes"].isin(meses))].copy()
    mask = (tmp["_tipo_norm"] == key_norm) | (tmp["_tipo_norm"].str.contains(key_norm, na=False))
    return tmp[mask].copy()


def build_table(linhas: List[Tuple[str, Dict[int, float]]], meses_sel: List[int], denom_by_month: Dict[int, float]) -> pd.DataFrame:
    """Monta tabela: JAN | JAN% | ... + ACUMULADO + %ACUMULADO."""
    rows = []
    for nome, by_month in linhas:
        row = {"LINHA": nome}
        for m in meses_sel:
            val = float(by_month.get(m, 0.0))
            denom = float(denom_by_month.get(m, 0.0))
            pct = (val / denom * 100.0) if denom != 0 else 0.0
            mes_pt = MES_NUM_TO_PT[m]
            row[mes_pt] = val
            row[f"{mes_pt}%"] = pct
        rows.append(row)

    df = pd.DataFrame(rows)

    if meses_sel:
        df["ACUMULADO"] = df[[MES_NUM_TO_PT[m] for m in meses_sel]].sum(axis=1, skipna=True)
    else:
        df["ACUMULADO"] = 0.0

    denom_acum = float(sum(float(denom_by_month.get(m, 0.0)) for m in meses_sel))
    df["%ACUMULADO"] = (df["ACUMULADO"] / denom_acum * 100.0) if denom_acum != 0 else 0.0
    return df


def style_result_row(df: pd.DataFrame, result_label: str) -> pd.io.formats.style.Styler:
    """Destaca a linha de resultado."""
    def _style(row):
        styles = [""] * len(row)
        if str(row.get("LINHA", "")) == result_label:
            for j, col in enumerate(row.index):
                if col == "LINHA":
                    styles[j] = "font-weight: 900;"
                    continue
                if ("%" in col):
                    styles[j] = "font-weight: 800;"
                    continue
                try:
                    val = float(row[col])
                except Exception:
                    continue
                styles[j] = "color: #c00000; font-weight: 900;" if val < 0 else "color: #1f4e79; font-weight: 900;"
        return styles

    fmt = {}
    for c in df.columns:
        if c == "LINHA":
            continue
        fmt[c] = (lambda x: fmt_pct(x)) if ("%" in c) else (lambda x: f"R$ {format_brl(x)}")
    return df.style.format(fmt).apply(_style, axis=1).hide(axis="index")

def style_result_rows(df: pd.DataFrame, labels: List[str]) -> pd.io.formats.style.Styler:
    """Destaca múltiplas linhas de resultado."""
    labels_set = set(labels or [])
    def _style(row):
        styles = ["" ] * len(row)
        if str(row.get("LINHA", "")) in labels_set:
            for j, col in enumerate(row.index):
                if col == "LINHA":
                    styles[j] = "font-weight: 900;"
                    continue
                if ("%" in col):
                    styles[j] = "font-weight: 800;"
                    continue
                try:
                    val = float(row[col])
                except Exception:
                    continue
                styles[j] = "color: #c00000; font-weight: 900;" if val < 0 else "color: #1f4e79; font-weight: 900;"
        return styles

    fmt = {}
    for c in df.columns:
        if c == "LINHA":
            continue
        fmt[c] = (lambda x: fmt_pct(x)) if ("%" in c) else (lambda x: f"R$ {format_brl(x)}")
    return df.style.format(fmt).apply(_style, axis=1).hide(axis="index")


# ----------------------------
# Drill (comum DRE / DFC)
# ----------------------------

def render_drill(
    base: pd.DataFrame,
    ano_ref: int,
    meses_sel: List[int],
    meses_pt_sel: List[str],
    linha_tipo: str,
    linha_by_month: Dict[int, float],
    denom_label: str,
    denom_by_month: Dict[int, float],
    ui_key_prefix: str,
):
    """
    Drill:
      - Sempre mostra o TOTAL da linha no período (mesmo sem detalhamento na BASE).
      - Se houver lançamentos, detalha por Plano de contas e Descrição.
      - KPIs:
          * Total do período
          * Média mensal do período
          * % de representatividade no denominador (Receita ou Recebimentos)
      - No detalhamento por Plano:
          * % no denominador
          * % do plano sobre a própria linha (representatividade na despesa/linha)
    """

    c1, c2 = st.columns([2, 1])
    with c1:
        mes_opt = ["TODOS"] + (meses_pt_sel if meses_pt_sel else MESES_PT)
        mes_sel = st.selectbox("Mês", options=mes_opt, index=0, key=f"{ui_key_prefix}_mes")
    with c2:
        st.caption(f"TIPO: {linha_tipo}")

    meses_drill = meses_sel if mes_sel == "TODOS" else [MES_PT_TO_NUM[mes_sel]]

    # Totais de referência (vindos da tabela/linha), independente da BASE
    total_linha_periodo = float(sum(float(linha_by_month.get(m, 0.0)) for m in meses_drill))
    media_linha_periodo = total_linha_periodo / max(len(meses_drill), 1)

    denom_periodo = float(sum(float(denom_by_month.get(m, 0.0)) for m in meses_drill))
    pct_linha_denom = (total_linha_periodo / denom_periodo * 100.0) if denom_periodo != 0 else 0.0

    k1, k2, k3 = st.columns([1, 1, 1])
    k1.metric("Total no período", f"R$ {format_brl(total_linha_periodo)}")
    k2.metric("Média mensal no período", f"R$ {format_brl(media_linha_periodo)}")
    k3.metric(f"Representatividade em {denom_label}", fmt_pct(pct_linha_denom))

    # Filtra BASE (pode estar vazia)
    base_f = filter_base(base, ano_ref, meses_drill, linha_tipo)

    # Se não houver detalhamento, ainda assim mostramos a linha como "SEM DETALHAMENTO"
    if base_f.empty:
        st.info("Sem detalhamento na BASE DE DADOS para a linha/período selecionado. Exibindo apenas o total do período.")
        plano_agg = pd.DataFrame(
            [{
                "PLANO DE CONTAS": "SEM DETALHAMENTO NA BASE",
                "VALOR": total_linha_periodo,
                f"% {denom_label}": pct_linha_denom,
                "% DA LINHA": 100.0 if total_linha_periodo != 0 else 0.0,
            }]
        )
        st.dataframe(
            plano_agg.style.format(
                {"VALOR": lambda x: f"R$ {format_brl(x)}",
                 f"% {denom_label}": lambda x: fmt_pct(x),
                 "% DA LINHA": lambda x: fmt_pct(x)}
            ).hide(axis="index"),
            use_container_width=True,
        )
        return

    # Agrega planos existentes na BASE
    plano_agg = (base_f.groupby("_plano_sint", dropna=False)["_v"].sum().reset_index()
                 .rename(columns={"_plano_sint": "PLANO DE CONTAS", "_v": "VALOR"}))

    # Ajuste: garante que o total do drill respeite o TOTAL da linha (tabela),
    # adicionando diferença como "SEM CLASSIFICAÇÃO NA BASE" se necessário.
    total_base = float(plano_agg["VALOR"].sum())
    diff = float(total_linha_periodo - total_base)
    if abs(diff) > 0.005:  # tolerância para arredondamentos
        plano_agg = pd.concat(
            [plano_agg,
             pd.DataFrame([{"PLANO DE CONTAS": "SEM CLASSIFICAÇÃO NA BASE", "VALOR": diff}])],
            ignore_index=True
        )

    plano_agg[f"% {denom_label}"] = (plano_agg["VALOR"] / denom_periodo * 100.0) if denom_periodo != 0 else 0.0
    plano_agg["% DA LINHA"] = (plano_agg["VALOR"] / total_linha_periodo * 100.0) if total_linha_periodo != 0 else 0.0
    plano_agg = plano_agg.sort_values("VALOR", ascending=False)

    left, right = st.columns([1.2, 1])
    with left:
        fig = px.pie(plano_agg.head(25), names="PLANO DE CONTAS", values="VALOR",
                     title="Representatividade (Plano de contas) — Top 25")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        st.metric("Total (linha)", f"R$ {format_brl(total_linha_periodo)}", fmt_pct(pct_linha_denom))

    st.dataframe(
        plano_agg.head(120).style.format(
            {"VALOR": lambda x: f"R$ {format_brl(x)}",
             f"% {denom_label}": lambda x: fmt_pct(x),
             "% DA LINHA": lambda x: fmt_pct(x)}
        ).hide(axis="index"),
        use_container_width=True,
    )

    plano_sel = st.selectbox("Selecione um Plano de contas", options=plano_agg["PLANO DE CONTAS"].tolist(), key=f"{ui_key_prefix}_plano")
    base_p = base_f[base_f["_plano_sint"] == plano_sel].copy()

    # Plano selecionado pode ser a linha de "SEM CLASSIFICAÇÃO NA BASE"
    if base_p.empty:
        # se for o item criado (diff), não existe base_p; mostramos apenas KPIs do plano
        plano_val = float(plano_agg.loc[plano_agg["PLANO DE CONTAS"] == plano_sel, "VALOR"].sum())
        pct_plano_denom = (plano_val / denom_periodo * 100.0) if denom_periodo != 0 else 0.0
        pct_plano_linha = (plano_val / total_linha_periodo * 100.0) if total_linha_periodo != 0 else 0.0
        st.info("Este item não possui lançamentos detalhados na BASE DE DADOS.")
        p1, p2, p3 = st.columns(3)
        p1.metric("Total do plano (período)", f"R$ {format_brl(plano_val)}")
        p2.metric(f"% em {denom_label}", fmt_pct(pct_plano_denom))
        p3.metric("% do plano na linha", fmt_pct(pct_plano_linha))
        return

    desc_agg = (base_p.groupby("_desc", dropna=False)["_v"].sum().reset_index()
                .rename(columns={"_desc": "DESCRIÇÃO", "_v": "VALOR"})).sort_values("VALOR", ascending=False)

    # KPIs do plano
    total_plano = float(base_p["_v"].sum())
    pct_plano_denom = (total_plano / denom_periodo * 100.0) if denom_periodo != 0 else 0.0
    pct_plano_linha = (total_plano / total_linha_periodo * 100.0) if total_linha_periodo != 0 else 0.0
    media_plano = total_plano / max(len(meses_drill), 1)

    st.divider()
    st.subheader("KPIs — plano selecionado")
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Plano selecionado", plano_sel)
    p2.metric("Total (período)", f"R$ {format_brl(total_plano)}")
    p3.metric("Média mensal (período)", f"R$ {format_brl(media_plano)}")
    p4.metric("% do plano na linha", fmt_pct(pct_plano_linha))

    st.caption(f"% do plano em {denom_label}: {fmt_pct(pct_plano_denom)}")

    st.subheader("Detalhamento por descrição")
    st.dataframe(
        desc_agg.head(200).style.format({"VALOR": lambda x: f"R$ {format_brl(x)}"}).hide(axis="index"),
        use_container_width=True,
    )

    st.subheader("Lançamentos (BASE)")
    cols_show = [c for c in ["_mes", "_ano", "_tipo", "_plano", "_desc", "_v"] if c in base_p.columns]
    if cols_show:
        view = base_p[cols_show].copy()
        if "_mes" in view.columns:
            view["_mes"] = view["_mes"].map(lambda x: MES_NUM_TO_PT.get(int(x), x))
        view = view.rename(columns={"_mes": "MÊS", "_ano": "ANO", "_tipo": "TIPO", "_plano": "PLANO", "_desc": "DESCRIÇÃO", "_v": "VALOR"})
        st.dataframe(view.style.format({"VALOR": lambda x: f"R$ {format_brl(x)}"}).hide(axis="index"), use_container_width=True)


# ----------------------------
# Páginas
# ----------------------------

def page_dre(plano_path: str, plano_sig: Tuple[int, int], dados_path: str, dados_sig: Tuple[int, int], ano_ref: int, meses_pt_sel: List[str]):
    st.title("DRE")

    meses_pt = meses_pt_sel if meses_pt_sel else MESES_PT
    meses_sel = [MES_PT_TO_NUM[m] for m in meses_pt]

    base_periodo_check = prep_base(plano_path, plano_sig, dados_path, dados_sig)
    if base_periodo_check is not None:
        qtd_periodo = len(base_periodo_check[(base_periodo_check["_ano"] == int(ano_ref)) & (base_periodo_check["_mes"].isin(meses_sel))])
        if qtd_periodo == 0:
            meses_com_dados = sorted(
                base_periodo_check.loc[base_periodo_check["_ano"] == int(ano_ref), "_mes"]
                .dropna().astype(int).unique().tolist()
            )
            nomes = ", ".join(MES_NUM_TO_PT[m] for m in meses_com_dados) if meses_com_dados else "nenhum mês"
            st.warning(
                f"O arquivo PLANO DE CONTAS não possui lançamentos nos meses selecionados. "
                f"Para {ano_ref}, há lançamentos em: {nomes}."
            )

    xls = get_excel_file(dados_path, dados_sig)

    df_receita = read_sheet(dados_path, "RECEITA", dados_sig)
    df_cmv = read_sheet(dados_path, "CMV", dados_sig)
    base = prep_base(plano_path, plano_sig, dados_path, dados_sig)

    missing = []
    if df_receita is None: missing.append("RECEITA")
    if df_cmv is None: missing.append("CMV")
    if base is None: missing.append("PLANO DE CONTAS / aba TIPO")

    if missing:
        st.error(f"Não consegui localizar abas essenciais: {', '.join(missing)}")
        st.write("Abas detectadas no Excel:", xls.sheet_names)
        st.stop()

    receita_by_month = agg_mes_ano_val(df_receita, ["receita", "RECEITA", "Receita"], ano_ref)
    cmv_by_month = agg_mes_ano_val(df_cmv, ["CMV", "cmv"], ano_ref)

    if receita_by_month is None:
        st.error("Aba RECEITA: preciso de colunas (mês, ano, receita).")
        st.write("Colunas encontradas:", list(df_receita.columns))
        st.stop()

    if cmv_by_month is None:
        st.error("Aba CMV: preciso de colunas (mês, ano, CMV).")
        st.write("Colunas encontradas:", list(df_cmv.columns))
        st.stop()

    margem_by_month = {m: float(receita_by_month.get(m, 0.0)) - float(cmv_by_month.get(m, 0.0)) for m in range(1, 13)}

    pessoal = sum_base_by_month(base, ano_ref, "DESPESA C/ PESSOAL")
    financeira = sum_base_by_month(base, ano_ref, "FINANCEIRA")
    impostos = sum_base_by_month(base, ano_ref, "IMPOSTOS")
    operacional = sum_base_by_month(base, ano_ref, "OPERACIONAL")

    despesas_total = {m: float(pessoal.get(m, 0.0)) + float(financeira.get(m, 0.0)) + float(impostos.get(m, 0.0)) + float(operacional.get(m, 0.0)) for m in range(1, 13)}
    resultado_operacional = {m: float(margem_by_month.get(m, 0.0)) - float(despesas_total.get(m, 0.0)) for m in range(1, 13)}


    # Cards topo
    receita_total_periodo = float(sum(float(receita_by_month.get(m, 0.0)) for m in meses_sel))
    res_op_total_periodo = float(sum(float(resultado_operacional.get(m, 0.0)) for m in meses_sel))
    pct_res_op = (res_op_total_periodo / receita_total_periodo * 100.0) if receita_total_periodo != 0 else 0.0

    c1, c2 = st.columns(2)
    c1.metric("Receita total (período)", f"R$ {format_brl(receita_total_periodo)}")
    c2.metric("Resultado operacional (período)", f"R$ {format_brl(res_op_total_periodo)}", fmt_pct(pct_res_op))

    linhas = [
        ("RECEITA", receita_by_month),
        ("CMV", cmv_by_month),
        ("MARGEM DE CONTRIBUIÇÃO", margem_by_month),
        ("DESPESA C/ PESSOAL", pessoal),
        ("FINANCEIRA", financeira),
        ("IMPOSTOS", impostos),
        ("OPERACIONAL", operacional),
        ("RESULTADO OPERACIONAL", resultado_operacional),
    ]

    st.subheader("Tabela (JAN | JAN% | ...)")
    dre_tbl = build_table(linhas, meses_sel, denom_by_month=receita_by_month)
    st.dataframe(style_result_row(dre_tbl, "RESULTADO OPERACIONAL"), use_container_width=True)

    st.divider()
    st.subheader("Drill — detalhamento (BASE DE DADOS)")

    linhas_map = {nome: by_month for (nome, by_month) in linhas}
    linhas_opt = list(linhas_map.keys())
    default_idx = linhas_opt.index("OPERACIONAL") if "OPERACIONAL" in linhas_opt else 0

    linha_sel = st.selectbox(
        "Selecione a linha para detalhar",
        linhas_opt,
        index=default_idx,
        key="dre_linha_sel",
    )

    render_drill(
        base=base,
        ano_ref=ano_ref,
        meses_sel=meses_sel,
        meses_pt_sel=meses_pt,
        linha_tipo=linha_sel,
        linha_by_month=linhas_map.get(linha_sel, {}),
        denom_label="RECEITA",
        denom_by_month=receita_by_month,
        ui_key_prefix="dre",
    )


def page_dfc(plano_path: str, plano_sig: Tuple[int, int], dados_path: str, dados_sig: Tuple[int, int], ano_ref: int, meses_pt_sel: List[str]):
    st.title("DFC")

    meses_pt = meses_pt_sel if meses_pt_sel else MESES_PT
    meses_sel = [MES_PT_TO_NUM[m] for m in meses_pt]

    base_periodo_check = prep_base(plano_path, plano_sig, dados_path, dados_sig)
    if base_periodo_check is not None:
        qtd_periodo = len(base_periodo_check[(base_periodo_check["_ano"] == int(ano_ref)) & (base_periodo_check["_mes"].isin(meses_sel))])
        if qtd_periodo == 0:
            meses_com_dados = sorted(
                base_periodo_check.loc[base_periodo_check["_ano"] == int(ano_ref), "_mes"]
                .dropna().astype(int).unique().tolist()
            )
            nomes = ", ".join(MES_NUM_TO_PT[m] for m in meses_com_dados) if meses_com_dados else "nenhum mês"
            st.warning(
                f"O arquivo PLANO DE CONTAS não possui lançamentos nos meses selecionados. "
                f"Para {ano_ref}, há lançamentos em: {nomes}."
            )

    xls = get_excel_file(dados_path, dados_sig)

    df_receb = read_sheet(dados_path, "RECEBIMENTO", dados_sig)
    base = prep_base(plano_path, plano_sig, dados_path, dados_sig)

    missing = []
    if df_receb is None: missing.append("RECEBIMENTO")
    if base is None: missing.append("PLANO DE CONTAS / aba TIPO")

    if missing:
        st.error(f"Não consegui localizar abas essenciais: {', '.join(missing)}")
        st.write("Abas detectadas no Excel:", xls.sheet_names)
        st.stop()

    receb_by_month = agg_mes_ano_val(df_receb, ["recebimento", "RECEBIMENTO", "Recebimento"], ano_ref)
    if receb_by_month is None:
        st.error("Aba RECEBIMENTO: preciso de colunas (mês, ano, recebimento).")
        st.write("Colunas encontradas:", list(df_receb.columns))
        st.stop()

    pessoal = sum_base_by_month(base, ano_ref, "DESPESA C/ PESSOAL")
    financeira = sum_base_by_month(base, ano_ref, "FINANCEIRA")
    impostos = sum_base_by_month(base, ano_ref, "IMPOSTOS")
    operacional = sum_base_by_month(base, ano_ref, "OPERACIONAL")
    emprestimos = sum_base_by_month(base, ano_ref, "EMPRÉSTIMOS")
    fornecedor = sum_base_by_month(base, ano_ref, "FORNECEDOR")

    # Resultado Caixa (inclui empréstimos nas despesas, conforme pedido)
    saidas_total = {m: float(pessoal.get(m, 0.0)) + float(financeira.get(m, 0.0)) + float(impostos.get(m, 0.0)) + float(operacional.get(m, 0.0)) + float(emprestimos.get(m, 0.0)) + float(fornecedor.get(m, 0.0)) for m in range(1, 13)}
    resultado_caixa = {m: float(receb_by_month.get(m, 0.0)) - float(saidas_total.get(m, 0.0)) for m in range(1, 13)}
    resultado_antes_emprest = {m: float(resultado_caixa.get(m, 0.0)) + float(emprestimos.get(m, 0.0)) for m in range(1, 13)}


    # Cards topo
    receb_total_periodo = float(sum(float(receb_by_month.get(m, 0.0)) for m in meses_sel))
    res_caixa_total_periodo = float(sum(float(resultado_caixa.get(m, 0.0)) for m in meses_sel))
    res_antes_emp_total_periodo = float(sum(float(resultado_antes_emprest.get(m, 0.0)) for m in meses_sel))

    pct_res_caixa = (res_caixa_total_periodo / receb_total_periodo * 100.0) if receb_total_periodo != 0 else 0.0
    pct_res_antes = (res_antes_emp_total_periodo / receb_total_periodo * 100.0) if receb_total_periodo != 0 else 0.0

    c1, c2, c3 = st.columns(3)
    c1.metric("Recebimentos (período)", f"R$ {format_brl(receb_total_periodo)}")
    c2.metric("Resultado Caixa (período)", f"R$ {format_brl(res_caixa_total_periodo)}", fmt_pct(pct_res_caixa))
    c3.metric("Resultado antes dos empréstimos (período)", f"R$ {format_brl(res_antes_emp_total_periodo)}", fmt_pct(pct_res_antes))

    linhas = [
        ("RECEBIMENTOS", receb_by_month),
        ("DESPESA C/ PESSOAL", pessoal),
        ("FINANCEIRA", financeira),
        ("IMPOSTOS", impostos),
        ("OPERACIONAL", operacional),
        ("FORNECEDOR", fornecedor),
        ("EMPRÉSTIMOS", emprestimos),
        ("RESULTADO ANTES DOS EMPRÉSTIMOS", resultado_antes_emprest),
        ("RESULTADO CAIXA", resultado_caixa),
    ]

    st.subheader("Tabela (JAN | JAN% | ...)")
    dfc_tbl = build_table(linhas, meses_sel, denom_by_month=receb_by_month)
    st.dataframe(style_result_rows(dfc_tbl, ["RESULTADO ANTES DOS EMPRÉSTIMOS", "RESULTADO CAIXA"]), use_container_width=True)

    st.divider()
    st.subheader("Drill — detalhamento (BASE DE DADOS)")

    linhas_map = {nome: by_month for (nome, by_month) in linhas}
    linhas_opt = list(linhas_map.keys())
    default_idx = linhas_opt.index("OPERACIONAL") if "OPERACIONAL" in linhas_opt else 0

    linha_sel = st.selectbox(
        "Selecione a linha para detalhar",
        linhas_opt,
        index=default_idx,
        key="dfc_linha_sel",
    )

    render_drill(
        base=base,
        ano_ref=ano_ref,
        meses_sel=meses_sel,
        meses_pt_sel=meses_pt,
        linha_tipo=linha_sel,
        linha_by_month=linhas_map.get(linha_sel, {}),
        denom_label="RECEBIMENTOS",
        denom_by_month=receb_by_month,
        ui_key_prefix="dfc",
    )


# ----------------------------
# App / Sidebar
# ----------------------------

st.set_page_config(page_title="DRE & DFC — Financeiro", layout="wide")
st.sidebar.title("Menu")

plano_path = find_plano_contas_file()
dados_path = find_dados_externos_file()

if not plano_path:
    st.sidebar.error('Não encontrei o arquivo "PLANO DE CONTAS" (.xlsx/.xlsm/.xls) na pasta do app.')
if not dados_path:
    st.sidebar.error('Não encontrei o arquivo "DADOS EXTERNOS" (.xlsx/.xlsm/.xls) na pasta do app.')
if not plano_path or not dados_path:
    st.info('No GitHub, mantenha os dois arquivos na mesma pasta do PY e use nomes contendo "PLANO DE CONTAS" e "DADOS EXTERNOS".')
    st.stop()

plano_sig = excel_signature(plano_path)
dados_sig = excel_signature(dados_path)
xls_plano = get_excel_file(plano_path, plano_sig)
xls_dados = get_excel_file(dados_path, dados_sig)

st.sidebar.caption(f"Lançamentos: **{os.path.basename(plano_path)}**")
st.sidebar.caption(f"Dados externos: **{os.path.basename(dados_path)}**")
st.sidebar.success("Arquivos carregados")

with st.sidebar.expander("Diagnóstico dos arquivos"):
    st.write("Abas — PLANO DE CONTAS:", xls_plano.sheet_names)
    st.write("Abas — DADOS EXTERNOS:", xls_dados.sheet_names)

base_tmp = prep_base(plano_path, plano_sig, dados_path, dados_sig)
if base_tmp is None:
    st.error("Não foi possível montar a base. O sistema aceita a aba TIPO com ou sem cabeçalho. Confira se PLANO DE CONTAS contém Descrição, Plano de contas, Valor total e uma coluna de data válida.")
    st.stop()

pendentes = planos_sem_tipo(base_tmp)
if not pendentes.empty:
    st.warning(f"Foram encontrados {len(pendentes)} plano(s) de contas sem TIPO vinculado na aba TIPO de DADOS EXTERNOS.")
    st.dataframe(pendentes, use_container_width=True, hide_index=True)
    st.download_button(
        "Baixar planos sem TIPO (CSV)",
        data=pendentes.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig"),
        file_name="planos_de_contas_sem_tipo.csv",
        mime="text/csv",
        use_container_width=True,
    )
    st.caption("Preencha a coluna TIPO, copie as linhas para a aba TIPO de DADOS EXTERNOS e atualize o arquivo no repositório.")

anos = set(base_tmp["_ano"].dropna().astype(int).unique().tolist())
for sheet in ["RECEITA", "CMV", "RECEBIMENTO"]:
    df_tmp = read_sheet(dados_path, sheet, dados_sig)
    if df_tmp is not None:
        col_ano = _find_col(df_tmp, ["ANO", "Ano"])
        if col_ano is not None:
            anos |= set(pd.to_numeric(df_tmp[col_ano], errors="coerce").dropna().astype(int).unique().tolist())

anos = sorted(anos)
if not anos:
    st.sidebar.error("Não encontrei nenhum ANO válido nos arquivos.")
    st.stop()

ano_ref = st.sidebar.selectbox("Ano", options=anos, index=len(anos) - 1)

# Por padrão, abre nos meses realmente existentes no arquivo PLANO DE CONTAS.
# Isso evita mostrar janeiro a junho zerados quando o relatório anexado contém apenas julho.
meses_base_ano = sorted(
    base_tmp.loc[base_tmp["_ano"] == int(ano_ref), "_mes"]
    .dropna().astype(int).unique().tolist()
)
meses_default = [MES_NUM_TO_PT[m] for m in meses_base_ano if m in MES_NUM_TO_PT]
if not meses_default:
    meses_default = MESES_PT

# A chave inclui a assinatura dos arquivos. Assim, ao substituir uma planilha,
# o filtro inicia nos meses existentes na nova base; depois disso, a seleção
# do usuário permanece livre e não é sobrescrita pelo código.
meses_key = (
    f"meses_{ano_ref}_"
    f"{plano_sig[0]}_{plano_sig[1]}_"
    f"{dados_sig[0]}_{dados_sig[1]}"
)

meses_pt_sel = st.sidebar.multiselect(
    "Meses",
    options=MESES_PT,
    default=meses_default,
    key=meses_key,
    help=(
        "Selecione livremente um ou mais meses. Meses sem lançamentos "
        "confirmados serão exibidos com valor zero."
    ),
)

# Não altera a seleção do usuário. Apenas informa quando o período escolhido
# não contém lançamentos confirmados no arquivo PLANO DE CONTAS.
meses_com_lancamentos = set(meses_default) if meses_default != MESES_PT else set()
if meses_pt_sel and meses_com_lancamentos:
    meses_escolhidos_com_dados = set(meses_pt_sel).intersection(meses_com_lancamentos)
    if not meses_escolhidos_com_dados:
        st.sidebar.warning(
            "Os meses selecionados não possuem lançamentos com "
            "'Data de confirmação'. A tabela será exibida com valores zerados "
            "para as contas do PLANO DE CONTAS."
        )

if not meses_pt_sel:
    st.sidebar.info(
        "Nenhum mês selecionado. Selecione pelo menos um mês para montar as tabelas."
    )

# Diagnóstico objetivo da base carregada.
base_ano = base_tmp[base_tmp["_ano"] == int(ano_ref)].copy()
if not base_ano.empty:
    dt_min = base_ano["_dt"].min()
    dt_max = base_ano["_dt"].max()
    st.sidebar.info(
        f"PLANO DE CONTAS: {len(base_ano):,} lançamentos em {ano_ref} "
        f"({dt_min.strftime('%d/%m/%Y')} a {dt_max.strftime('%d/%m/%Y')})."
        .replace(',', '.')
    )

pagina = st.sidebar.radio("Página", ["DRE", "DFC"])

if pagina == "DRE":
    page_dre(plano_path, plano_sig, dados_path, dados_sig, ano_ref, meses_pt_sel)
else:
    page_dfc(plano_path, plano_sig, dados_path, dados_sig, ano_ref, meses_pt_sel)
