"""
Processamento de dados, cruzamento de fontes, cálculo de disponibilidade e motor vetorizado de conflitos (v0.3).
"""

from typing import Dict, List, Optional, Set, Tuple
import numpy as np
import pandas as pd

from .config import Config
from .schemas import VALID_ACTIVITY_TYPES, normalizar_projeto


def classify_conflict_severity(tipo_a: str, tipo_b: str) -> str:
    """
    Classifica a severidade do conflito:
    - 'Crítico': Férias ou Licença sobreposta com atividade presencial/embarque/estaleiro.
    - 'Operacional': Duas atividades simultâneas de trabalho/planejamento.
    """
    rest_types = {"Férias", "Licença"}
    active_types = {"Estaleiro", "Embarque", "Treinamento", "Workshop", "Visita Técnica"}

    t_a = str(tipo_a).strip().title()
    t_b = str(tipo_b).strip().title()

    if (t_a in rest_types and t_b in active_types) or (t_b in rest_types and t_a in active_types):
        return Config.SEVERITY_CRITICO
    if t_a in rest_types and t_b in rest_types:
        return Config.SEVERITY_CRITICO
    return Config.SEVERITY_OPERACIONAL


class DataProcessor:
    """Motor de processamento de alocação de equipe e detecção vetorizada de conflitos."""

    @staticmethod
    def _mat_set(series: pd.Series) -> Set[str]:
        """Extrai um conjunto de matrículas normalizadas como strings."""
        return set(series.dropna().astype(str))

    @staticmethod
    def prepare_combined_data(
        equipe_df: pd.DataFrame,
        dfs_eventos: List[Optional[pd.DataFrame]],
    ) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Consolida eventos de múltiplas origens (estaleiro, férias, geral, plataformas)
        e cruza com a equipe filtrada, garantindo integridade e ordenação por disciplina.
        """
        valid_dfs = []
        matriculas_validas = DataProcessor._mat_set(equipe_df["Matrícula"])

        for df in dfs_eventos:
            if df is None or df.empty:
                continue
            df_filtered = df[df["Matrícula"].astype(str).isin(matriculas_validas)].copy()
            cols = ["Matrícula", "Início", "Término", "Tipo"]
            for extra in ["Nome", "Detalhamento", "Plataforma"]:
                if extra in df_filtered.columns:
                    cols.append(extra)
            valid_dfs.append(df_filtered[cols])

        if not valid_dfs:
            combined = pd.DataFrame(columns=["Matrícula", "Nome", "Início", "Término", "Disciplina", "Tipo"])
        else:
            combined = pd.concat(valid_dfs, ignore_index=True)

        for c in ["Detalhamento", "Plataforma"]:
            if c not in combined.columns:
                combined[c] = pd.NA

        combined = combined.dropna(subset=["Início", "Término"])
        combined["Tipo"] = combined["Tipo"].astype("string").str.strip().str.title()
        valid_types = VALID_ACTIVITY_TYPES
        combined = combined[combined["Tipo"].isin(valid_types)].reset_index(drop=True)

        combined = combined.merge(
            equipe_df[["Matrícula", "Nome", "Disciplina", "Função", "Projeto", "Experiência"]],
            on="Matrícula",
            how="left",
            suffixes=("", "_eq"),
        )
        for c in ["Nome", "Disciplina", "Função", "Projeto", "Experiência"]:
            eq = f"{c}_eq"
            if eq in combined.columns:
                combined[c] = combined[eq].where(combined[eq].notna(), combined[c])
                combined = combined.drop(columns=[eq])

        # Inicializa colunas de status de conflito no DataFrame consolidado
        combined["Tem_Conflito"] = False
        combined["Detalhe_Conflito"] = pd.NA

        disc_rank = {d: i for i, d in enumerate(Config.DISCIPLINA_ORDER)}
        unique_members = equipe_df.copy()
        unique_members["_rank"] = unique_members["Disciplina"].map(disc_rank).fillna(999)
        unique_members = unique_members.sort_values(["_rank", "Função", "Projeto", "Nome"]).drop(columns="_rank")
        unique_members = unique_members.reset_index(drop=True)

        return combined, unique_members

    @staticmethod
    def available_in_period(
        unique_members: pd.DataFrame,
        combined_df: pd.DataFrame,
        start_date: pd.Timestamp,
        end_date: pd.Timestamp,
    ) -> Set[str]:
        """Identifica o conjunto de matrículas que NÃO possuem nenhum evento no período."""
        if combined_df.empty:
            return DataProcessor._mat_set(unique_members["Matrícula"])
        mask = (combined_df["Início"] <= end_date) & (combined_df["Término"] >= start_date)
        occupied = DataProcessor._mat_set(combined_df.loc[mask, "Matrícula"])
        return DataProcessor._mat_set(unique_members["Matrícula"]) - occupied

    @staticmethod
    def occupied_in_period(
        combined_df: pd.DataFrame,
        start_date: pd.Timestamp,
        end_date: pd.Timestamp,
    ) -> Set[str]:
        """Identifica o conjunto de matrículas que possuem pelo menos um evento ativo no período."""
        if combined_df.empty:
            return set()
        mask = (combined_df["Início"] <= end_date) & (combined_df["Término"] >= start_date)
        return DataProcessor._mat_set(combined_df.loc[mask, "Matrícula"])

    @staticmethod
    def events_in_period(
        combined_df: pd.DataFrame,
        start_date: pd.Timestamp,
        end_date: pd.Timestamp,
    ) -> pd.DataFrame:
        """Filtra eventos que tenham sobreposição com o intervalo de datas especificado."""
        if combined_df.empty:
            return combined_df.copy()
        mask = (combined_df["Início"] <= end_date) & (combined_df["Término"] >= start_date)
        return combined_df.loc[mask].copy()

    @staticmethod
    def detect_conflicts(combined_df: pd.DataFrame) -> Tuple[pd.DataFrame, Set[Tuple[int, pd.Timestamp, str]]]:
        """
        Detecção de Conflitos Vetorizada ordenada por ['Matrícula', 'Início', 'Término'].
        Dois intervalos [A_inicio, A_fim] e [B_inicio, B_fim] do mesmo colaborador conflitam
        se e somente se A_inicio <= B_fim e B_inicio <= A_fim.

        Gera:
        - Coluna booleana 'Tem_Conflito'
        - Coluna descritiva 'Detalhe_Conflito' indicando atividade e dias de interseção
        - Coluna 'Severidade' ('Crítico' vs 'Operacional')
        """
        cols = [
            "Nome", "Matrícula", "Projeto", "Severidade", "Tem_Conflito", "Detalhe_Conflito",
            "Atividade A", "Período A", "Início A", "Término A",
            "Atividade B", "Período B", "Início B", "Término B",
            "Início Conflito", "Término Conflito", "Período Conflito", "Sobreposição (dias)",
        ]
        if combined_df.empty:
            return pd.DataFrame(columns=cols), set()

        df = combined_df.dropna(subset=["Início", "Término"]).copy()
        if df.empty:
            return pd.DataFrame(columns=cols), set()

        # Ordenação canônica obrigatória
        df = df.sort_values(by=["Matrícula", "Início", "Término"]).reset_index(drop=True)

        rows = []
        keys: Set[Tuple[int, pd.Timestamp, str]] = set()

        # Dicionário auxiliar para preencher combined_df com os detalhes do conflito
        conflict_descriptions_by_idx: Dict[int, List[str]] = {}

        for mat, group in df.groupby("Matrícula", sort=False):
            if len(group) < 2:
                continue

            n = len(group)
            group_indices = group.index.to_numpy()
            starts = group["Início"].to_numpy()
            ends = group["Término"].to_numpy()

            # Broadcasting vetorizado com índices superiores
            idx_i, idx_j = np.triu_indices(n, k=1)

            # Condição necessária e suficiente: A_inicio <= B_fim e B_inicio <= A_fim
            overlap_mask = (starts[idx_i] <= ends[idx_j]) & (starts[idx_j] <= ends[idx_i])
            if not np.any(overlap_mask):
                continue

            matching_i = idx_i[overlap_mask]
            matching_j = idx_j[overlap_mask]

            records = group.to_dict("records")
            try:
                mat_key = int(mat)
            except (ValueError, TypeError):
                mat_key = mat

            for i, j in zip(matching_i, matching_j):
                orig_idx_i = group_indices[i]
                orig_idx_j = group_indices[j]

                a = records[i]
                b = records[j]

                overlap_start = max(a["Início"], b["Início"])
                overlap_end = min(a["Término"], b["Término"])
                overlap_days = (overlap_end - overlap_start).days + 1

                # Classificação de Severidade (Crítico vs Operacional)
                severidade = classify_conflict_severity(a["Tipo"], b["Tipo"])

                # Descrição detalhada da colisão
                detalhe = (
                    f"Sobreposição de {overlap_days} dia(s) entre {a['Tipo']} "
                    f"({a['Início']:%d/%m/%Y} a {a['Término']:%d/%m/%Y}) e "
                    f"{b['Tipo']} ({b['Início']:%d/%m/%Y} a {b['Término']:%d/%m/%Y})"
                )

                conflict_descriptions_by_idx.setdefault(orig_idx_i, []).append(detalhe)
                conflict_descriptions_by_idx.setdefault(orig_idx_j, []).append(detalhe)

                # Resolução do projeto normalizado
                proj_val = None
                for item in (a, b):
                    for col_name in ("Projeto", "Plataforma"):
                        v = item.get(col_name)
                        if pd.notna(v) and str(v).strip() != "" and str(v).strip().lower() != "nan":
                            proj_val = str(v).strip()
                            break
                    if proj_val:
                        break
                proj_norm = normalizar_projeto(proj_val) if proj_val else "P-80"

                rows.append({
                    "Nome": a["Nome"],
                    "Matrícula": mat,
                    "Projeto": proj_norm,
                    "Severidade": severidade,
                    "Tem_Conflito": True,
                    "Detalhe_Conflito": detalhe,
                    "Atividade A": a["Tipo"],
                    "Período A": f"{a['Início']:%d/%m/%Y} – {a['Término']:%d/%m/%Y}",
                    "Início A": a["Início"],
                    "Término A": a["Término"],
                    "Atividade B": b["Tipo"],
                    "Período B": f"{b['Início']:%d/%m/%Y} – {b['Término']:%d/%m/%Y}",
                    "Início B": b["Início"],
                    "Término B": b["Término"],
                    "Início Conflito": overlap_start,
                    "Término Conflito": overlap_end,
                    "Período Conflito": f"{overlap_start:%d/%m/%Y} – {overlap_end:%d/%m/%Y}",
                    "Sobreposição (dias)": overlap_days,
                })
                keys.add((mat_key, a["Início"], a["Tipo"]))
                keys.add((mat_key, b["Início"], b["Tipo"]))

        # Atualiza colunas Tem_Conflito e Detalhe_Conflito em combined_df
        if conflict_descriptions_by_idx:
            for idx, descs in conflict_descriptions_by_idx.items():
                if idx in combined_df.index:
                    combined_df.at[idx, "Tem_Conflito"] = True
                    combined_df.at[idx, "Detalhe_Conflito"] = " | ".join(descs)

        conflicts = pd.DataFrame(rows, columns=cols)
        return conflicts, keys

    @staticmethod
    def filter_future_conflicts(
        conflicts_full: pd.DataFrame,
        reference_date: Optional[pd.Timestamp] = None,
    ) -> pd.DataFrame:
        """Filtra apenas conflitos cujo término da sobreposição seja estritamente posterior à data de referência."""
        if conflicts_full.empty or "Término Conflito" not in conflicts_full.columns:
            return pd.DataFrame(columns=conflicts_full.columns)
        ref = reference_date or pd.Timestamp.today().normalize()
        return conflicts_full[conflicts_full["Término Conflito"] > ref].copy().reset_index(drop=True)

    @staticmethod
    def build_conflict_details(conflicts_full: pd.DataFrame) -> Dict[Tuple[int, pd.Timestamp, str], List[str]]:
        """Gera um dicionário mapeando chave de evento para descrições textuais das sobreposições."""
        conflict_details: Dict[Tuple[int, pd.Timestamp, str], List[str]] = {}
        if conflicts_full.empty:
            return conflict_details

        for _, row in conflicts_full.iterrows():
            try:
                mat_key = int(row["Matrícula"])
            except (ValueError, TypeError):
                mat_key = row["Matrícula"]

            key_a = (mat_key, pd.Timestamp(row["Início A"]), row["Atividade A"])
            key_b = (mat_key, pd.Timestamp(row["Início B"]), row["Atividade B"])
            sev_badge = f"[{row.get('Severidade', 'Operacional')}] " if "Severidade" in row else ""
            desc = f"{sev_badge}{row['Atividade A']} ({row['Período A']}) x {row['Atividade B']} ({row['Período B']}) · {row.get('Sobreposição (dias)', 1)}d"
            conflict_details.setdefault(key_a, []).append(desc)
            conflict_details.setdefault(key_b, []).append(desc)

        return conflict_details
