"""
Esquemas de dados, validação de cabeçalhos e funções utilitárias de normalização (v0.3).
"""

from typing import Dict, Iterable, List, Optional, Set, Tuple
import pandas as pd

from .config import Config


class MissingColumnError(Exception):
    """Exceção levantada quando colunas essenciais esperadas estão ausentes em um arquivo."""

    def __init__(self, filename: str, missing_columns: List[str]):
        self.filename = filename
        self.missing_columns = missing_columns
        msg = (
            f"O arquivo/aba '{filename}' não contém as colunas obrigatórias: "
            f"{', '.join(missing_columns)}. Verifique o cabeçalho da planilha."
        )
        super().__init__(msg)


# Colunas padrão da aba Equipe
EQUIPE_COLUMNS: List[str] = [
    "Disciplina",
    "Matrícula",
    "Função",
    "Projeto",
    "Experiência",
    "Nome",
    "E-mail",
]

# Colunas padrão dos eventos de planejamento
PLANEJAMENTO_COLUMNS: List[str] = [
    "Nome",
    "Matrícula",
    "Início",
    "Término",
    "Tipo",
    "Detalhamento",
    "Plataforma",
    "Experiência",
]

# Tipos de atividade válidos reconhecidos pela aplicação
VALID_ACTIVITY_TYPES: Set[str] = {
    "Estaleiro",
    "Embarque",
    "Folga",
    "Férias",
    "Treinamento",
    "Workshop",
    "Visita Técnica",
    "Atestado",
    "Licença",
    "Emprestado",
}

# De-para de siglas operacionais para tipos canônicos de atividade
STATUS_TO_TIPO_MAP: Dict[str, str] = {
    "EMB": "Embarque",
    "E": "Embarque",
    "FOL": "Folga",
    "FÉR": "Férias",
    "FÉRIAS": "Férias",
    "BASE": "Estaleiro",
    "TREIN": "Treinamento",
    "TREINAMENTO": "Treinamento",
    "PMB": "Treinamento",
    "LIC": "Licença",
    "LICENÇA": "Licença",
    "LICENCA": "Licença",
    "EMP": "Emprestado",
    "EMPRESTADO": "Emprestado",
}


def apply_column_aliases(df: pd.DataFrame, aliases: Optional[Dict[str, str]] = None) -> pd.DataFrame:
    """
    Aplica .str.strip() em todos os cabeçalhos e mapeia variações conhecidas (aliases)
    para o padrão canônico interno (ex.: 'Início', 'Término', 'Matrícula', 'Nome', etc.).
    """
    if len(df.columns) == 0:
        return df

    alias_map = aliases or Config.COLUMN_ALIASES
    renamed: Dict[str, str] = {}

    for col in df.columns:
        col_clean = str(col).strip()
        col_key = col_clean.lower()
        if col_key in alias_map:
            renamed[col] = alias_map[col_key]
        else:
            renamed[col] = col_clean

    return df.rename(columns=renamed)


def validate_required_columns(df: pd.DataFrame, required: List[str], filename: str) -> None:
    """
    Validação 'fail-fast' amigável: verifica a presença das colunas obrigatórias.
    Lança MissingColumnError detalhado caso alguma coluna esteja ausente.
    """
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise MissingColumnError(filename=filename, missing_columns=missing)


def parse_date_series(series: pd.Series) -> pd.Series:
    """
    Converte uma série para pd.Timestamp tratando tanto formato ISO (YYYY-MM-DD)
    quanto formato brasileiro (DD/MM/AAAA) via dayfirst=True com coerção segura.
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    # Tenta formato ISO primeiro para evitar conflitos de dia/mês
    res = pd.to_datetime(series, format="%Y-%m-%d", errors="coerce")
    mask = res.isna() & series.notna()
    if mask.any():
        # Para datas brasileiras e outros formatos, utiliza dayfirst=True
        res.loc[mask] = pd.to_datetime(series.loc[mask], dayfirst=True, errors="coerce")
    return res


def sanitize_dates(
    df: pd.DataFrame,
    start_col: str = "Início",
    end_col: str = "Término",
    filename: str = "Planilha",
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Converte colunas de data usando dayfirst=True e errors='coerce'.
    Separa e audita linhas onde:
    - Início ou Término sejam NaT.
    - Término seja estritamente anterior a Início (Término < Início).

    Retorna tupla: (df_valido, df_inconsistencias).
    """
    if df.empty or start_col not in df.columns or end_col not in df.columns:
        empty_inconsistencias = pd.DataFrame(
            columns=["Arquivo", "Linha", "Nome", "Matrícula", "Início_Original", "Término_Original", "Motivo"]
        )
        return df.copy(), empty_inconsistencias

    work = df.copy()
    raw_start = work[start_col].copy()
    raw_end = work[end_col].copy()

    work[start_col] = parse_date_series(work[start_col])
    work[end_col] = parse_date_series(work[end_col])

    mask_nat = work[start_col].isna() | work[end_col].isna()
    mask_inverso = (~mask_nat) & (work[end_col] < work[start_col])

    invalid_mask = mask_nat | mask_inverso

    if not invalid_mask.any():
        empty_inconsistencias = pd.DataFrame(
            columns=["Arquivo", "Linha", "Nome", "Matrícula", "Início_Original", "Término_Original", "Motivo"]
        )
        return work, empty_inconsistencias

    inval_df = work[invalid_mask].copy()
    reasons = []
    for idx, is_nat, is_inv in zip(inval_df.index, mask_nat[invalid_mask], mask_inverso[invalid_mask]):
        if is_nat:
            reasons.append("Data Inválida ou Ausente (NaT)")
        elif is_inv:
            reasons.append("Término anterior ao Início")
        else:
            reasons.append("Inconsistência de Data")

    inconsistencias = pd.DataFrame({
        "Arquivo": filename,
        "Linha": inval_df.index + 2,  # Ajuste aproximado para numeração do Excel (1-based + header)
        "Nome": inval_df["Nome"] if "Nome" in inval_df.columns else pd.NA,
        "Matrícula": inval_df["Matrícula"] if "Matrícula" in inval_df.columns else pd.NA,
        "Início_Original": raw_start.loc[inval_df.index],
        "Término_Original": raw_end.loc[inval_df.index],
        "Motivo": reasons,
    })

    valid_df = work[~invalid_mask].copy().reset_index(drop=True)
    return valid_df, inconsistencias.reset_index(drop=True)


def normalize_dates(df: pd.DataFrame, cols: Iterable[str]) -> pd.DataFrame:
    """Converte colunas para pd.Timestamp/datetime com dayfirst=True e coerção de erros."""
    for col in cols:
        if col in df.columns:
            df[col] = parse_date_series(df[col])
    return df


def to_matricula(series: pd.Series) -> pd.Series:
    """Normaliza coluna de matrículas para strings numéricas inteiras sem decimais."""
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.apply(lambda x: str(int(x)) if pd.notna(x) else pd.NA)


def normalize_text(series: pd.Series, uppercase: bool = True) -> pd.Series:
    """Limpa espaços em branco e ajusta a caixa do texto (maiúsculo ou strip simples)."""
    s = series.astype("string").str.strip()
    return s.str.upper() if uppercase else s


def normalizar_projeto(p: Optional[str]) -> str:
    """Padroniza nomenclaturas de projetos/plataformas (ex: 'P80' -> 'P-80')."""
    if not p:
        return "P-80"
    p_limpo = str(p).strip().upper().replace(" ", "").replace("-", "")
    if "80" in p_limpo:
        return "P-80"
    if "82" in p_limpo:
        return "P-82"
    if "83" in p_limpo:
        return "P-83"
    return str(p).strip()


def validate_equipe_schema(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    """Valida se o DataFrame possui as colunas essenciais para o cadastro de equipe."""
    required = ["Nome", "Matrícula"]
    missing = [c for c in required if c not in df.columns]
    return (len(missing) == 0, missing)


def validate_planejamento_schema(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    """Valida se o DataFrame possui as colunas essenciais de planejamento temporal."""
    required = ["Início", "Término"]
    missing = [c for c in required if c not in df.columns]
    return (len(missing) == 0, missing)
