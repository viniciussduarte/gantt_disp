"""
Carregamento, download, validação fail-fast e sanitização de dados com cache (v0.3).
"""

import os
from typing import Any, List, Optional, Tuple
import pandas as pd
import requests
import streamlit as st

from .config import Config
from .schemas import (
    STATUS_TO_TIPO_MAP,
    VALID_ACTIVITY_TYPES,
    MissingColumnError,
    apply_column_aliases,
    normalize_dates,
    normalize_text,
    sanitize_dates,
    to_matricula,
    validate_required_columns,
)


class DataLoader:
    """Carregamento seguro, com validação de esquema, auditoria de datas e cache de planilhas."""

    _inconsistencias_list: List[pd.DataFrame] = []

    @classmethod
    def clear_inconsistencias(cls) -> None:
        """Limpa o registro de inconsistências acumuladas."""
        cls._inconsistencias_list = []

    @classmethod
    def record_inconsistencias(cls, df: pd.DataFrame) -> None:
        """Registra inconsistências de datas detectadas para exibição na interface."""
        if df is not None and not df.empty:
            cls._inconsistencias_list.append(df)

    @classmethod
    def get_inconsistencias(cls) -> pd.DataFrame:
        """Retorna todas as inconsistências de datas acumuladas na sessão atual."""
        if not cls._inconsistencias_list:
            return pd.DataFrame(
                columns=["Arquivo", "Linha", "Nome", "Matrícula", "Início_Original", "Término_Original", "Motivo"]
            )
        return pd.concat(cls._inconsistencias_list, ignore_index=True)

    @staticmethod
    def _normalize_dates(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
        """Encapsula a normalização de datas delegando ao módulo de schemas."""
        return normalize_dates(df, cols)

    @staticmethod
    def _to_matricula(series: pd.Series) -> pd.Series:
        """Encapsula a normalização de matrículas delegando ao módulo de schemas."""
        return to_matricula(series)

    @staticmethod
    def _download_from_drive(file_id: str, dest_path: str) -> bool:
        """Realiza o download de arquivo compartilhado do Google Drive com verificação de integridade."""
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        try:
            resp = requests.get(url, timeout=60)
            if resp.status_code == 200 and len(resp.content) > 1000:
                with open(dest_path, "wb") as f:
                    f.write(resp.content)
                if dest_path.endswith(".xlsx") and not DataLoader._is_valid_xlsx(dest_path):
                    return False
                return True
            return False
        except Exception as e:
            try:
                st.error(f"Erro ao baixar do Google Drive ({file_id}): {e}")
            except Exception:
                pass
            return False

    @staticmethod
    def _is_valid_xlsx(path: str) -> bool:
        """Verifica se o cabeçalho do arquivo binário corresponde a um arquivo zip/xlsx válido."""
        try:
            with open(path, "rb") as f:
                return f.read(4)[:2] == b"PK"
        except Exception:
            return False

    @staticmethod
    def _ensure_drive_file(filename: str, local_path: str) -> bool:
        """Garante a existência do arquivo local, tentando baixá-lo do Drive se configurado."""
        drive_id = Config.GOOGLE_DRIVE_IDS.get(filename)
        if not drive_id:
            return os.path.exists(local_path)
        if DataLoader._download_from_drive(drive_id, local_path):
            return True
        return os.path.exists(local_path)

    @staticmethod
    @st.cache_data(show_spinner="Carregando planilhas...")
    def load_estaleiro_data(source: Optional[Any] = None) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        """
        Carrega os dados de Equipe e Planejamento IED da planilha de Estaleiro.
        Suporta tanto caminho em disco quanto UploadedFile do Streamlit.
        Lança MissingColumnError em caso de colunas obrigatórias ausentes.
        """
        file_to_read = source or Config.FILE_PATH_ESTALEIRO

        if source is None:
            if not DataLoader._ensure_drive_file("Planejamento Estaleiro.xlsx", Config.FILE_PATH_ESTALEIRO):
                st.error("Não foi possível obter o Planejamento Estaleiro (nem do Drive nem local).")
                return None, None

        try:
            # 1. Carregamento da aba Equipe
            equipe_df = pd.read_excel(file_to_read, sheet_name="Equipe")
            equipe_df.columns = equipe_df.columns.astype(str).str.strip()
            equipe_df = apply_column_aliases(equipe_df)

            validate_required_columns(
                equipe_df,
                ["Nome", "Matrícula"],
                filename="Planejamento Estaleiro.xlsx (Aba Equipe)",
            )

            keep = ["Disciplina", "Matrícula", "Função", "Projeto", "Experiência", "Nome", "E-mail"]
            keep = [c for c in keep if c in equipe_df.columns]
            equipe_df = equipe_df[keep].copy()

            equipe_df = equipe_df.dropna(subset=["Nome", "Matrícula"])
            equipe_df["Matrícula"] = to_matricula(equipe_df["Matrícula"])
            equipe_df = equipe_df.dropna(subset=["Matrícula"])

            for c in ["Disciplina", "Função", "Projeto", "Experiência", "E-mail"]:
                if c in equipe_df.columns:
                    equipe_df[c] = equipe_df[c].astype("string")
            equipe_df["Nome"] = normalize_text(equipe_df["Nome"], uppercase=True)
            for c in ["Disciplina", "Função", "Projeto", "Experiência"]:
                if c in equipe_df.columns:
                    equipe_df[c] = normalize_text(equipe_df[c], uppercase=True)

            # 2. Carregamento da aba Planejamento IED
            plan_df = pd.read_excel(
                file_to_read,
                sheet_name="Planejamento IED",
                skiprows=8,
                usecols="C:E",
            )
            plan_df.columns = ["Nome", "Início", "Término"]
            plan_df.columns = plan_df.columns.astype(str).str.strip()
            plan_df = apply_column_aliases(plan_df)

            validate_required_columns(
                plan_df,
                ["Nome", "Início", "Término"],
                filename="Planejamento Estaleiro.xlsx (Aba Planejamento IED)",
            )

            plan_df = plan_df.dropna(subset=["Nome"])
            plan_df["Nome"] = normalize_text(plan_df["Nome"], uppercase=True)

            plan_df, inc = sanitize_dates(plan_df, "Início", "Término", filename="Planejamento Estaleiro.xlsx")
            DataLoader.record_inconsistencias(inc)

            plan_df = plan_df.merge(
                equipe_df[["Nome", "Matrícula", "Disciplina", "Função", "Projeto"]],
                on="Nome",
                how="left",
            )
            plan_df["Matrícula"] = to_matricula(plan_df["Matrícula"])
            plan_df["Tipo"] = "Estaleiro"

            return equipe_df, plan_df
        except MissingColumnError:
            raise
        except Exception as e:
            st.error(f"Erro ao carregar o Planejamento Estaleiro: {e}")
            return None, None

    @staticmethod
    @st.cache_data(show_spinner="Carregando planilhas...")
    def load_ferias_data(source: Optional[Any] = None) -> Optional[pd.DataFrame]:
        """
        Carrega e transforma os períodos parcelados de férias em linhas individuais.
        Aplica sanitização com dayfirst=True e audita datas inconsistentes.
        """
        file_to_read = source or Config.FILE_PATH_FERIAS
        try:
            df = pd.read_excel(file_to_read)
            df.columns = df.columns.astype(str).str.strip()
            df = apply_column_aliases(df)

            validate_required_columns(df, ["Matrícula"], filename="Férias.xlsx")

            pairs = [
                ("Primeira Parcela", "Termino Primeira Parcela"),
                ("Segunda Parcela", "Termino Segunda Parcela"),
                ("Terceira Parcela", "Termino Terceira Parcela"),
            ]
            frames = []
            for ini_col, fim_col in pairs:
                if ini_col not in df.columns or fim_col not in df.columns:
                    continue
                tmp = df[["Matrícula", ini_col, fim_col]].copy()
                tmp.columns = ["Matrícula", "Início", "Término"]
                tmp = tmp.dropna(subset=["Início"])

                tmp, inc = sanitize_dates(tmp, "Início", "Término", filename="Férias.xlsx")
                DataLoader.record_inconsistencias(inc)
                frames.append(tmp)

            if not frames:
                ferias = pd.DataFrame(columns=["Matrícula", "Início", "Término", "Tipo"])
            else:
                ferias = pd.concat(frames, ignore_index=True)
                ferias["Matrícula"] = to_matricula(ferias["Matrícula"])
                ferias = ferias.dropna(subset=["Matrícula"])
                ferias["Tipo"] = "Férias"
            return ferias
        except MissingColumnError:
            raise
        except Exception as e:
            st.error(f"Erro ao carregar o arquivo de Férias: {e}")
            return None

    @staticmethod
    @st.cache_data(show_spinner="Carregando planilhas...")
    def load_geral_equipe_data(source: Optional[Any] = None) -> Optional[pd.DataFrame]:
        """Carrega a aba Equipe do Planejamento Geral como base cadastral corporativa."""
        file_to_read = source or Config.FILE_PATH_GERAL
        if source is None:
            if not DataLoader._ensure_drive_file("Planejamento Geral.xlsx", Config.FILE_PATH_GERAL):
                st.error("Não foi possível obter o Planejamento Geral (nem do Drive nem local).")
                return None
        try:
            equipe_df = pd.read_excel(file_to_read, sheet_name="Equipe")
            equipe_df.columns = equipe_df.columns.astype(str).str.strip()
            equipe_df = apply_column_aliases(equipe_df)

            validate_required_columns(
                equipe_df,
                ["Nome", "Matrícula"],
                filename="Planejamento Geral.xlsx (Aba Equipe)",
            )

            keep = ["Disciplina", "Matrícula", "Função", "Projeto", "Experiência", "Nome", "E-mail"]
            keep = [c for c in keep if c in equipe_df.columns]
            equipe_df = equipe_df[keep].copy()

            equipe_df = equipe_df.dropna(subset=["Nome", "Matrícula"])
            equipe_df["Matrícula"] = to_matricula(equipe_df["Matrícula"])
            equipe_df = equipe_df.dropna(subset=["Matrícula"])

            for c in ["Disciplina", "Função", "Projeto", "Experiência", "E-mail"]:
                if c in equipe_df.columns:
                    equipe_df[c] = equipe_df[c].astype("string")
            equipe_df["Nome"] = normalize_text(equipe_df["Nome"], uppercase=True)
            for c in ["Disciplina", "Função", "Projeto", "Experiência"]:
                if c in equipe_df.columns:
                    equipe_df[c] = normalize_text(equipe_df[c], uppercase=True)

            return equipe_df
        except MissingColumnError:
            raise
        except Exception as e:
            st.error(f"Erro ao carregar a aba Equipe do Planejamento Geral: {e}")
            return None

    @staticmethod
    @st.cache_data(show_spinner="Carregando planilhas...")
    def load_planejamento_geral(source: Optional[Any] = None) -> Optional[pd.DataFrame]:
        """Carrega eventos globais da primeira aba do Planejamento Geral com sanitização de datas."""
        file_to_read = source or Config.FILE_PATH_GERAL
        if source is None:
            if not DataLoader._ensure_drive_file("Planejamento Geral.xlsx", Config.FILE_PATH_GERAL):
                st.error("Não foi possível obter o Planejamento Geral (nem do Drive nem local).")
                return None
        try:
            df = pd.read_excel(file_to_read)
            df.columns = df.columns.astype(str).str.strip()
            df = apply_column_aliases(df)

            validate_required_columns(
                df,
                ["Nome", "Matrícula", "Início", "Término"],
                filename="Planejamento Geral.xlsx",
            )

            df["Nome"] = normalize_text(df["Nome"], uppercase=True)
            df["Matrícula"] = to_matricula(df["Matrícula"])
            df = df.dropna(subset=["Matrícula", "Início"])

            df, inc = sanitize_dates(df, "Início", "Término", filename="Planejamento Geral.xlsx")
            DataLoader.record_inconsistencias(inc)

            if "Tipo" not in df.columns:
                df["Tipo"] = "Planejamento"
            df["Tipo"] = df["Tipo"].astype("string").str.strip().str.title()
            valid_types = VALID_ACTIVITY_TYPES
            df = df[df["Tipo"].isin(valid_types)].reset_index(drop=True)
            return df
        except MissingColumnError:
            raise
        except Exception as e:
            st.error(f"Erro ao carregar o Planejamento Geral: {e}")
            return None

    @staticmethod
    def load_escalas_plataformas_data(equipe_df: Optional[pd.DataFrame] = None) -> Optional[pd.DataFrame]:
        """Carrega e converte períodos de escalas das plataformas (SQLite) em eventos de cronograma."""
        try:
            from escala_p80 import EscalaP80DB
            periodos, _ = EscalaP80DB.get_periodos_agrupados("TODOS")
            if not periodos:
                return pd.DataFrame(
                    columns=["Nome", "Matrícula", "Início", "Término", "Tipo", "Plataforma", "Detalhamento"]
                )

            rows = []
            for p in periodos:
                st_tipo = STATUS_TO_TIPO_MAP.get(str(p.get("status", "")).strip().upper(), "Embarque")
                rows.append({
                    "Nome": p["nome"],
                    "Início": pd.Timestamp(p["inicio"]),
                    "Término": pd.Timestamp(p["fim"]),
                    "Tipo": st_tipo,
                    "Plataforma": p.get("projeto", "P-80"),
                    "Detalhamento": f"Escala {p.get('projeto', 'P-80')}",
                })

            df = pd.DataFrame(rows)
            df, inc = sanitize_dates(df, "Início", "Término", filename="Escalas Plataformas (SQLite)")
            DataLoader.record_inconsistencias(inc)

            if equipe_df is not None and not equipe_df.empty:
                df = df.merge(
                    equipe_df[["Nome", "Matrícula"]].drop_duplicates(subset=["Nome"]),
                    on="Nome",
                    how="left",
                )
            else:
                df["Matrícula"] = pd.NA
            return df
        except Exception:
            return None
