"""
Camada de dados de Escalas & Embarques do Dashboard Qualitech.

Lê o arquivo "Relatório Mensal de Eventos*.xlsx" (abas: Base Dashboard,
Base Workers, Análise_Dobras, Análise_AFI) e devolve um EmbarquesData com
os indicadores de escala/embarque prontos para os cards/gráficos do
app.py — em paralelo aos indicadores operacionais (data.py) e de RH
(workforce.py).

Conceitos da planilha, mapeados para a linguagem de escala/embarque:
- "Diárias Regulares" + "Dobras"  -> dias efetivamente embarcados/trabalhados
  (uma "Dobra" é um turno extra / embarque além do programado).
- "AFI" (Folga Indenizada)        -> folga.
- "Ociosidade" / STB              -> standby (aguardando embarque, sem OS).

A aba "Base Dashboard" traz a série mensal (um mês por linha); a aba
"Base Workers" traz o consolidado por funcionário no período coberto pela
base. Segue o mesmo estilo de leitura tolerante das demais camadas: cada
bloco é lido dentro de um try/except e, se a aba não existir ou vier vazia,
o indicador correspondente fica com um DataFrame vazio em vez de quebrar
o app inteiro.
"""
from __future__ import annotations

import glob
import os
from dataclasses import dataclass, field

import pandas as pd

# --------------------------------------------------------------------------- #
# Descoberta de arquivo
# --------------------------------------------------------------------------- #

_KEYWORDS = ("eventos", "embarque", "escala", "stb")


def find_embarques_workbook(base_dir: str) -> str | None:
    candidates = list_available_embarques_workbooks(base_dir)
    return candidates[0] if candidates else None


def list_available_embarques_workbooks(base_dir: str) -> list[str]:
    if not base_dir or not os.path.isdir(base_dir):
        return []
    all_xlsx = glob.glob(os.path.join(base_dir, "**", "*.xlsx"), recursive=True)
    all_xlsx = [c for c in all_xlsx if not os.path.basename(c).startswith("~$")]
    candidates = [c for c in all_xlsx if any(k in os.path.basename(c).lower() for k in _KEYWORDS)]
    candidates.sort(key=os.path.getmtime, reverse=True)
    return candidates


# --------------------------------------------------------------------------- #
# Estrutura de dados devolvida
# --------------------------------------------------------------------------- #

@dataclass
class EmbarquesData:
    path: str
    mtime: float
    periodo_label: str

    monthly: pd.DataFrame = field(default_factory=pd.DataFrame)
    workers: pd.DataFrame = field(default_factory=pd.DataFrame)
    por_familia: pd.DataFrame = field(default_factory=pd.DataFrame)
    observacoes: pd.DataFrame = field(default_factory=pd.DataFrame)

    total_funcionarios: int = 0
    funcoes_n: int = 0
    familias_n: int = 0
    dobras_total: int = 0
    afi_total: int = 0
    ociosidade_total: int = 0
    pct_com_dobra: float = 0.0
    pct_com_afi: float = 0.0
    pct_com_stb: float = 0.0

    @property
    def latest(self) -> dict | None:
        """Último mês com dados lançados na série mensal (Base Dashboard)."""
        if self.monthly.empty:
            return None
        return self.monthly.iloc[-1].to_dict()

    @property
    def mes_ref_label(self) -> str:
        last = self.latest
        return str(last["mes"]) if last else "—"

    def _top(self, col: str, n: int = 15) -> pd.DataFrame:
        if self.workers.empty:
            return pd.DataFrame()
        d = self.workers[self.workers[col] > 0].sort_values(col, ascending=False).head(n)
        return d

    @property
    def top_dobras(self) -> pd.DataFrame:
        return self._top("dobras")

    @property
    def top_afi(self) -> pd.DataFrame:
        return self._top("afi")

    @property
    def top_ociosidade(self) -> pd.DataFrame:
        return self._top("ociosidade")


# --------------------------------------------------------------------------- #
# Helpers de leitura
# --------------------------------------------------------------------------- #

def _read_observacoes(path: str, sheet_names: set, sheet: str, qty_col: int, evento_label: str) -> pd.DataFrame:
    if sheet not in sheet_names:
        return pd.DataFrame()
    try:
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
    except Exception:
        return pd.DataFrame()
    rows = []
    for _, r in raw.iloc[2:].iterrows():
        nome = r[1] if len(r) > 1 else None
        funcao = r[2] if len(r) > 2 else None
        qty = r[qty_col] if len(r) > qty_col else None
        just = r[9] if len(r) > 9 else None
        if isinstance(nome, str) and nome.strip() and isinstance(just, str) and just.strip():
            rows.append(dict(nome=nome.strip(), funcao=funcao, evento=evento_label,
                              quantidade=qty, comentario=just.strip()))
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# Loader principal
# --------------------------------------------------------------------------- #

def load_embarques_data(path: str) -> EmbarquesData:
    mtime = os.path.getmtime(path)
    periodo_label = os.path.splitext(os.path.basename(path))[0]
    try:
        sheet_names = set(pd.ExcelFile(path).sheet_names)
    except Exception:
        sheet_names = set()

    # ---- Base Dashboard: série mensal (efetivo, HD, regulares, dobras, AFI, ociosidade) ----
    monthly_df = pd.DataFrame()
    if "Base Dashboard" in sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name="Base Dashboard", header=0)
            raw = raw.iloc[:, :16].copy()
            raw.columns = [
                "mes_data", "mes", "efetivo", "hd", "capacidade", "regulares", "afi",
                "dobras", "total_diarias", "ociosidade", "pct_reg_cap", "pct_afi_total",
                "pct_dobras_total", "pct_ociosidade", "pct_dobras_hd", "pct_afi_hd",
            ]
            raw["mes"] = raw["mes"].astype(str)
            monthly_df = raw.dropna(subset=["efetivo"]).reset_index(drop=True)
        except Exception:
            monthly_df = pd.DataFrame()

    # ---- Base Workers: consolidado por funcionário no período ----
    workers_df = pd.DataFrame()
    if "Base Workers" in sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name="Base Workers", header=0, usecols=range(9))
            raw.columns = ["nome", "funcao", "familia", "dobras", "afi", "ociosidade",
                           "pct_dobras_hd", "pct_afi_hd", "pct_ociosidade"]
            workers_df = raw.dropna(subset=["nome"]).reset_index(drop=True)
            for c in ("dobras", "afi", "ociosidade"):
                workers_df[c] = pd.to_numeric(workers_df[c], errors="coerce").fillna(0).astype(int)
        except Exception:
            workers_df = pd.DataFrame()

    total_funcionarios = len(workers_df)
    funcoes_n = int(workers_df["funcao"].nunique()) if not workers_df.empty else 0
    familias_n = int(workers_df["familia"].nunique()) if not workers_df.empty else 0
    dobras_total = int(workers_df["dobras"].sum()) if not workers_df.empty else 0
    afi_total = int(workers_df["afi"].sum()) if not workers_df.empty else 0
    ociosidade_total = int(workers_df["ociosidade"].sum()) if not workers_df.empty else 0
    pct_com_dobra = float((workers_df["dobras"] > 0).mean()) if not workers_df.empty else 0.0
    pct_com_afi = float((workers_df["afi"] > 0).mean()) if not workers_df.empty else 0.0
    pct_com_stb = float((workers_df["ociosidade"] > 0).mean()) if not workers_df.empty else 0.0

    # ---- Composição por família de função (para o gráfico de barras) ----
    por_familia_df = pd.DataFrame()
    if not workers_df.empty:
        agg = workers_df.groupby("familia")[["dobras", "afi", "ociosidade"]].sum()
        agg["total"] = agg["dobras"] + agg["afi"] + agg["ociosidade"]
        agg = agg.sort_values("total", ascending=False).head(10).reset_index()
        por_familia_df = agg

    # ---- Observações da liderança (Análise_Dobras / Análise_AFI) ----
    obs_dobras = _read_observacoes(path, sheet_names, "Análise_Dobras", 4, "Dobra extra")
    obs_afi = _read_observacoes(path, sheet_names, "Análise_AFI", 5, "Folga indenizada (AFI)")
    observacoes_df = pd.concat([obs_dobras, obs_afi], ignore_index=True) if (not obs_dobras.empty or not obs_afi.empty) else pd.DataFrame()

    return EmbarquesData(
        path=path, mtime=mtime, periodo_label=periodo_label,
        monthly=monthly_df, workers=workers_df, por_familia=por_familia_df,
        observacoes=observacoes_df,
        total_funcionarios=total_funcionarios, funcoes_n=funcoes_n, familias_n=familias_n,
        dobras_total=dobras_total, afi_total=afi_total, ociosidade_total=ociosidade_total,
        pct_com_dobra=pct_com_dobra, pct_com_afi=pct_com_afi, pct_com_stb=pct_com_stb,
    )
