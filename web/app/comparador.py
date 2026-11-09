"""
Core de comparação de planilhas do RH.
Portado e adaptado de comparador_rh_v2.py para uso web.

Modelo atual da planilha do RH (aba "Ativos"):
  linha 1  → título ("Efetivo Aliseo CLT e PJ")
  linha 2  → cabeçalho: Código | Nome | Função | Departamento | Data de Admissão
  linha 3+ → colaboradores

O carregamento localiza o cabeçalho sozinho (não depende da linha) e aceita os
nomes antigos das colunas (Matrícula, Cargo, Admissão), de modo que snapshots
gravados no modelo anterior continuam comparáveis.
"""

import re
import shutil
import json
import unicodedata
from datetime import datetime
from pathlib import Path
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Configuração de colunas ────────────────────────────────────────────────
CONFIG = {
    "col_codigo":       "CÓDIGO",
    "col_nome":         "NOME",
    "col_funcao":       "FUNÇÃO",
    "col_departamento": "DEPARTAMENTO",
    "col_admissao":     "DATA DE ADMISSÃO",
}

CAMPOS_MONITORADOS = ["FUNÇÃO", "DEPARTAMENTO"]

# Mapeamento para chaves limpas no JSON (sem acentos).
# Atenção: "matricula", "cargo" e "admissao" mantêm os nomes do modelo antigo de
# propósito — os relatórios já gravados em data/relatorios/ usam essas chaves e
# precisam continuar abrindo normalmente. Só os rótulos exibidos mudaram.
CAMPO_MAP = {
    "TIPO":             "tipo",
    "CÓDIGO":           "matricula",
    "NOME":             "nome",
    "FUNÇÃO":           "cargo",
    "DEPARTAMENTO":     "departamento",
    "DATA DE ADMISSÃO": "admissao",
}

# Nomes aceitos no cabeçalho (sem acento, em maiúsculas) → coluna canônica.
_ALIASES = {
    "CODIGO":           "CÓDIGO",
    "MATRICULA":        "CÓDIGO",
    "NOME":             "NOME",
    "FUNCAO":           "FUNÇÃO",
    "CARGO":            "FUNÇÃO",
    "DEPARTAMENTO":     "DEPARTAMENTO",
    "DATA DE ADMISSAO": "DATA DE ADMISSÃO",
    "ADMISSAO":         "DATA DE ADMISSÃO",
}
_COLUNAS_OBRIGATORIAS = ["CÓDIGO", "NOME", "FUNÇÃO", "DEPARTAMENTO"]
_ABA_EFETIVO = "ATIVOS"       # a planilha também traz a aba "Demitidos", que é ignorada
_MAX_LINHAS_CABECALHO = 20    # quantas linhas do topo são vasculhadas atrás do cabeçalho

# Snapshots são gravados com a extensão original (.xls ou .xlsx)
GLOB_SNAPSHOTS = "snapshot_*.xls*"

# ── Estilos Excel ──────────────────────────────────────────────────────────
FONT_TITLE   = Font(name="Arial", size=13, bold=True, color="FFFFFF")
FONT_SECTION = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FONT_HEADER  = Font(name="Arial", size=10, bold=True, color="FFFFFF")
FONT_BODY    = Font(name="Arial", size=10)
FONT_BOLD    = Font(name="Arial", size=10, bold=True)

FILL_TITLE  = PatternFill("solid", start_color="263238")
FILL_ADD    = PatternFill("solid", start_color="2E7D32")
FILL_REM    = PatternFill("solid", start_color="C62828")
FILL_ALT    = PatternFill("solid", start_color="EF6C00")
FILL_HEADER = PatternFill("solid", start_color="455A64")
FILL_ZEBRA  = PatternFill("solid", start_color="ECEFF1")
FILL_DIFF   = PatternFill("solid", start_color="FFF59D")

THIN   = Side(border_style="thin", color="B0BEC5")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)


# ── Carga e normalização ───────────────────────────────────────────────────

def _sem_acento(texto) -> str:
    t = unicodedata.normalize("NFD", str(texto))
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


def _coluna_canonica(valor) -> str | None:
    """Traduz o texto de uma célula de cabeçalho para a coluna canônica (ou None)."""
    if valor is None or pd.isna(valor):
        return None
    chave = re.sub(r"\s+", " ", _sem_acento(valor)).strip().upper()
    return _ALIASES.get(chave)


def _ler_aba_efetivo(caminho: Path) -> pd.DataFrame:
    """Lê a aba do efetivo sem interpretar cabeçalho (todas as células como texto)."""
    with pd.ExcelFile(caminho) as xls:
        # Nomes de aba vêm com espaços sobrando ("Ativos "), por isso o strip().
        aba = next(
            (n for n in xls.sheet_names if n.strip().upper() == _ABA_EFETIVO),
            xls.sheet_names[0],  # modelo antigo: aba única com nome qualquer
        )
        return xls.parse(aba, header=None, dtype=str)


def _localizar_cabecalho(bruto: pd.DataFrame) -> tuple[int, list]:
    """Acha a linha de cabeçalho (a que tem Código/Matrícula e Nome)."""
    for i in range(min(len(bruto), _MAX_LINHAS_CABECALHO)):
        canonicas = [_coluna_canonica(v) for v in bruto.iloc[i]]
        if "CÓDIGO" in canonicas and "NOME" in canonicas:
            return i, canonicas
    raise ValueError(
        "Cabeçalho não encontrado. A planilha precisa ter as colunas "
        "Código, Nome, Função, Departamento e Data de Admissão "
        f"(procurei nas {_MAX_LINHAS_CABECALHO} primeiras linhas da aba 'Ativos')."
    )


def _limpar_codigo(v) -> str:
    """'610.0' → '610' (o pandas lê número como float)."""
    v = str(v).strip()
    if v.endswith(".0") and v[:-2].replace("-", "").isdigit():
        return v[:-2]
    return v


_RE_DATA_ISO = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T]\d{2}:\d{2}:\d{2})?$")


def _formatar_data(v: str) -> str:
    """'2026-07-07 00:00:00' (data do Excel lida como texto) → '07/07/2026'."""
    m = _RE_DATA_ISO.match(v)
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else v


def carregar_planilha(caminho: Path) -> pd.DataFrame:
    """Carrega, normaliza e indexa uma planilha do RH."""
    bruto = _ler_aba_efetivo(caminho)
    linha_cab, canonicas = _localizar_cabecalho(bruto)

    # posição da coluna na planilha → nome canônico (a 1ª ocorrência vence)
    colunas: dict[int, str] = {}
    for pos, nome in enumerate(canonicas):
        if nome and nome not in colunas.values():
            colunas[pos] = nome

    faltando = [c for c in _COLUNAS_OBRIGATORIAS if c not in colunas.values()]
    if faltando:
        raise ValueError(
            "Colunas não encontradas na planilha: " + ", ".join(faltando) + ". "
            "Esperado: Código, Nome, Função, Departamento e Data de Admissão."
        )

    df = bruto.iloc[linha_cab + 1:, list(colunas)].copy()
    df.columns = list(colunas.values())
    df = df.fillna("").astype(str).apply(lambda s: s.str.strip()).reset_index(drop=True)

    col_c = CONFIG["col_codigo"]
    col_n = CONFIG["col_nome"]
    col_a = CONFIG["col_admissao"]
    if col_a not in df.columns:      # opcional: só é exibida
        df[col_a] = ""

    df[col_c] = df[col_c].apply(_limpar_codigo)
    df[col_a] = df[col_a].apply(_formatar_data)

    # Descarta linhas sem informação útil: sem código E sem nome (linhas em branco,
    # rodapé) ou PJ sem nome (não há como identificar).
    pj_sem_nome = (df[col_c].str.upper().isin(["", "PJ"])) & (df[col_n] == "")
    df = df[~pj_sem_nome].copy()

    if df.empty:
        raise ValueError(
            "Nenhum colaborador encontrado na aba 'Ativos'. "
            "A planilha contém apenas o cabeçalho?"
        )

    def gerar_chave(row):
        codigo = row[col_c].upper()
        if codigo in ("", "PJ"):        # PJ: sem código numérico → identificado pelo nome
            return f"PJ::{row[col_n].upper()}"
        return f"CLT::{codigo}"

    df["_CHAVE"] = df.apply(gerar_chave, axis=1)

    # Chave duplicada (erro do RH): mantém a última ocorrência
    df = df.drop_duplicates(subset="_CHAVE", keep="last")

    return df.set_index("_CHAVE")


def comparar(df_antigo: pd.DataFrame, df_novo: pd.DataFrame):
    """Retorna (adicoes, remocoes, alteracoes) como DataFrames."""
    k_ant, k_nov = set(df_antigo.index), set(df_novo.index)

    adicoes  = df_novo.loc[list(k_nov - k_ant)].copy()
    remocoes = df_antigo.loc[list(k_ant - k_nov)].copy()

    linhas_alt = []
    for chave in k_ant & k_nov:
        antigo = df_antigo.loc[chave]
        novo   = df_novo.loc[chave]
        diffs  = {}
        for campo in CAMPOS_MONITORADOS:
            v_ant = str(antigo.get(campo, "")).strip()
            v_nov = str(novo.get(campo, "")).strip()
            if v_ant != v_nov:
                diffs[campo] = (v_ant, v_nov)
        if diffs:
            linha = {
                "TIPO":   "PJ" if chave.startswith("PJ::") else "CLT",
                "CÓDIGO": novo.get(CONFIG["col_codigo"], ""),
                "NOME":   novo.get(CONFIG["col_nome"], ""),
            }
            for campo in CAMPOS_MONITORADOS:
                if campo in diffs:
                    linha[f"{campo}_ANTES"]  = diffs[campo][0]
                    linha[f"{campo}_DEPOIS"] = diffs[campo][1]
                    linha[f"{campo}_MUDOU"]  = True
                else:
                    v = novo.get(campo, "")
                    linha[f"{campo}_ANTES"]  = v
                    linha[f"{campo}_DEPOIS"] = v
                    linha[f"{campo}_MUDOU"]  = False
            linhas_alt.append(linha)

    alteracoes = pd.DataFrame(linhas_alt)

    for df in (adicoes, remocoes):
        if not df.empty:
            df.insert(0, "TIPO",
                      ["PJ" if k.startswith("PJ::") else "CLT" for k in df.index])

    return adicoes, remocoes, alteracoes


def _df_to_list(df: pd.DataFrame, campos: list) -> list:
    """Converte DataFrame para lista de dicts com chaves limpas (sem acentos)."""
    result = []
    for _, row in df.iterrows():
        item = {}
        for campo in campos:
            valor = row.get(campo, "")
            clean = CAMPO_MAP.get(campo, campo.lower())
            item[clean] = str(valor) if pd.notna(valor) else ""
        result.append(item)
    return result


def extrair_data(arquivo: Path) -> str:
    # Usa apenas "." ou "-" como separadores (não "_") para evitar capturar
    # o timestamp prefixado no nome do arquivo (ex: 20260527_123626_efetivo-26.05.xlsx
    # teria "27_12" como falso-positivo se "_" fosse separador aceito).
    m = re.search(r"(\d{2})[.\-](\d{2})(?:[.\-](\d{2,4}))?", arquivo.stem)
    if not m:
        return datetime.now().strftime("%d-%m-%Y")
    dd, mm, yy = m.group(1), m.group(2), m.group(3)
    return f"{dd}-{mm}-{yy}" if yy else f"{dd}-{mm}"


# ── Processamento principal ────────────────────────────────────────────────

def processar_planilha(
    arquivo_novo: Path,
    pasta_snapshots: Path,
    pasta_relatorios: Path,
    usuario: str = "sistema",
) -> dict:
    """
    Processa uma nova planilha do RH:
    1. Carrega e normaliza o arquivo
    2. Compara com o snapshot mais recente (se existir)
    3. Salva resultado em JSON
    4. Salva snapshot do arquivo para próxima comparação

    Retorna dict com todos os dados da comparação.
    """
    pasta_snapshots.mkdir(parents=True, exist_ok=True)
    pasta_relatorios.mkdir(parents=True, exist_ok=True)

    df_novo = carregar_planilha(arquivo_novo)
    n_pj  = sum(1 for k in df_novo.index if k.startswith("PJ::"))
    n_clt = sum(1 for k in df_novo.index if k.startswith("CLT::"))

    snapshots = sorted(pasta_snapshots.glob(GLOB_SNAPSHOTS))
    snapshot_anterior = snapshots[-1] if snapshots else None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    data_str  = extrair_data(arquivo_novo)

    resultado: dict = {
        "id":        timestamp,
        "data":      data_str,
        "arquivo":   arquivo_novo.name,
        "timestamp": datetime.now().isoformat(),
        "usuario":   usuario,
        "total":     len(df_novo),
        "total_clt": n_clt,
        "total_pj":  n_pj,
        "is_baseline": snapshot_anterior is None,
        "n_adicoes":    0,
        "n_remocoes":   0,
        "n_alteracoes": 0,
        "adicoes":    [],
        "remocoes":   [],
        "alteracoes": [],
    }

    if snapshot_anterior is not None:
        df_antigo = carregar_planilha(snapshot_anterior)
        adicoes, remocoes, alteracoes = comparar(df_antigo, df_novo)

        campos_adic = ["TIPO", "CÓDIGO", "NOME", "FUNÇÃO", "DEPARTAMENTO", "DATA DE ADMISSÃO"]
        campos_rem  = ["TIPO", "CÓDIGO", "NOME", "FUNÇÃO", "DEPARTAMENTO"]

        resultado["n_adicoes"]  = len(adicoes)
        resultado["n_remocoes"] = len(remocoes)
        resultado["n_alteracoes"] = len(alteracoes)

        resultado["adicoes"]  = _df_to_list(adicoes,  campos_adic) if not adicoes.empty else []
        resultado["remocoes"] = _df_to_list(remocoes, campos_rem)  if not remocoes.empty else []

        alt_list = []
        for _, row in alteracoes.iterrows():
            item = {
                "tipo":      str(row.get("TIPO", "")),
                "matricula": str(row.get("CÓDIGO", "")),
                "nome":      str(row.get("NOME", "")),
            }
            for campo in CAMPOS_MONITORADOS:
                k = CAMPO_MAP[campo]
                item[f"{k}_antes"]  = str(row.get(f"{campo}_ANTES", ""))
                item[f"{k}_depois"] = str(row.get(f"{campo}_DEPOIS", ""))
                item[f"{k}_mudou"]  = bool(row.get(f"{campo}_MUDOU", False))
            alt_list.append(item)
        resultado["alteracoes"] = alt_list

        # Alerta se mais de 50% removidos (possível problema de chave)
        if len(df_antigo) > 0 and len(remocoes) / len(df_antigo) > 0.5:
            resultado["alerta_chave"] = (
                f"Mais de 50% dos colaboradores marcados como removidos "
                f"({len(remocoes)}/{len(df_antigo)}). "
                "Possível problema de formato na planilha."
            )

    # Salva resultado em JSON
    json_path = pasta_relatorios / f"resultado_{timestamp}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    # Salva snapshot para próxima comparação (mantém a extensão original: .xls/.xlsx)
    snapshot_path = (
        pasta_snapshots / f"snapshot_{timestamp}_{data_str}{arquivo_novo.suffix.lower()}"
    )
    shutil.copy2(arquivo_novo, snapshot_path)

    return resultado


# ── Geração de relatório Excel para download ───────────────────────────────

def gerar_excel(resultado: dict) -> bytes:
    """Gera relatório Excel formatado a partir dos dados JSON. Retorna bytes."""
    wb = Workbook()
    ws = wb.active
    ws.title = f"Alterações {resultado['data']}"

    data_str   = resultado["data"]
    arquivo    = resultado["arquivo"]
    adicoes    = resultado.get("adicoes", [])
    remocoes   = resultado.get("remocoes", [])
    alteracoes = resultado.get("alteracoes", [])

    N_COLS = 7  # maior tabela (alterações): Tipo, Código, Nome + 2×(antes/depois)

    # ── Cabeçalho ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=N_COLS)
    c = ws["A1"]
    c.value = f"Relatório de Alterações — {data_str}"
    c.font, c.fill, c.alignment = FONT_TITLE, FILL_TITLE, CENTER
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=N_COLS)
    ws["A2"].value = f"Arquivo: {arquivo}"
    ws["A2"].font  = Font(name="Arial", size=9, italic=True, color="546E7A")

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=N_COLS)
    ws["A3"].value = (
        f"Resumo: {resultado['n_adicoes']} admissões  |  "
        f"{resultado['n_remocoes']} desligamentos  |  "
        f"{resultado['n_alteracoes']} alterações  |  "
        f"Total: {resultado['total']} colaboradores"
    )
    ws["A3"].font = FONT_BOLD

    linha = 5

    def _escrever_secao(ws, linha, titulo, fill, dados, colunas_display, colunas_chave):
        ws.merge_cells(
            start_row=linha, start_column=1,
            end_row=linha,   end_column=len(colunas_display),
        )
        c = ws.cell(row=linha, column=1, value=titulo)
        c.font, c.fill, c.alignment = FONT_SECTION, fill, LEFT
        ws.row_dimensions[linha].height = 20
        linha += 1

        if not dados:
            c = ws.cell(row=linha, column=1, value="Nenhuma ocorrência.")
            c.font = Font(name="Arial", size=10, italic=True, color="78909C")
            ws.merge_cells(
                start_row=linha, start_column=1,
                end_row=linha, end_column=len(colunas_display),
            )
            return linha + 2

        for idx, col in enumerate(colunas_display, 1):
            c = ws.cell(row=linha, column=idx, value=col)
            c.font, c.fill, c.alignment, c.border = FONT_HEADER, FILL_HEADER, CENTER, BORDER
        linha += 1

        for i, item in enumerate(dados):
            for j, chave in enumerate(colunas_chave, 1):
                c = ws.cell(row=linha, column=j, value=item.get(chave, ""))
                c.font = FONT_BODY
                c.alignment = CENTER if j <= 2 else LEFT
                c.border = BORDER
                if i % 2 == 1:
                    c.fill = FILL_ZEBRA
            linha += 1
        return linha + 1

    linha = _escrever_secao(
        ws, linha,
        f"➕  ADMISSÕES ({len(adicoes)})", FILL_ADD,
        adicoes,
        ["Tipo", "Código", "Nome", "Função", "Departamento", "Admissão"],
        ["tipo", "matricula", "nome", "cargo", "departamento", "admissao"],
    )

    linha = _escrever_secao(
        ws, linha,
        f"➖  DESLIGAMENTOS ({len(remocoes)})", FILL_REM,
        remocoes,
        ["Tipo", "Código", "Nome", "Função", "Departamento"],
        ["tipo", "matricula", "nome", "cargo", "departamento"],
    )

    # Seção de alterações
    ws.merge_cells(
        start_row=linha, start_column=1,
        end_row=linha, end_column=N_COLS,
    )
    c = ws.cell(row=linha, column=1, value=f"✏️   ALTERAÇÕES ({len(alteracoes)})")
    c.font, c.fill, c.alignment = FONT_SECTION, FILL_ALT, LEFT
    ws.row_dimensions[linha].height = 20
    linha += 1

    if not alteracoes:
        c = ws.cell(row=linha, column=1, value="Nenhuma ocorrência.")
        c.font = Font(name="Arial", size=10, italic=True, color="78909C")
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=N_COLS)
    else:
        cabecalhos = [
            "Tipo", "Código", "Nome",
            "Função Antes", "Função Depois",
            "Depto Antes",  "Depto Depois",
        ]
        for idx, h in enumerate(cabecalhos, 1):
            c = ws.cell(row=linha, column=idx, value=h)
            c.font, c.fill, c.alignment, c.border = FONT_HEADER, FILL_HEADER, CENTER, BORDER
        linha += 1

        campos_alt_keys = [
            "tipo", "matricula", "nome",
            "cargo_antes",        "cargo_depois",
            "departamento_antes", "departamento_depois",
        ]
        # coluna (1-based) → flag que indica se o campo mudou
        mudou_map = {
            4: "cargo_mudou",        5: "cargo_mudou",
            6: "departamento_mudou", 7: "departamento_mudou",
        }
        for i, item in enumerate(alteracoes):
            for j, chave in enumerate(campos_alt_keys, 1):
                c = ws.cell(row=linha, column=j, value=item.get(chave, ""))
                c.font = FONT_BODY
                c.alignment = CENTER if j <= 2 else LEFT
                c.border = BORDER
                if i % 2 == 1:
                    c.fill = FILL_ZEBRA
                if j in mudou_map and item.get(mudou_map[j]):
                    c.fill = FILL_DIFF
                    c.font = FONT_BOLD
            linha += 1

    larguras = {"A": 7, "B": 11, "C": 30, "D": 28, "E": 28, "F": 26, "G": 26}
    for col, w in larguras.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A4"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
