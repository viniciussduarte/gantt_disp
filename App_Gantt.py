"""
Aplicação Streamlit para visualização de alocação de equipe, férias e planejamento.
"""
import os
import warnings
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# --- CONFIGURAÇÃO -----------------------------------------------------------
class Config:
    """Configurações e constantes da aplicação."""

    FILE_PATH_ESTALEIRO = "Planejamento Estaleiro.xlsx"
    FILE_PATH_FERIAS = "Férias.xlsx"
    FILE_PATH_GERAL = "Planejamento Geral.xlsx"

    GOOGLE_DRIVE_IDS = {
        "Planejamento Estaleiro.xlsx": "1VYvNJV9V4vUYeQgCA7DjfZY0ggYOpbP3",
        "Planejamento Geral.xlsx": "1NSP2p69F33kqE_FLvsbfugLjDc2v-fS5",
    }

    DISCIPLINA_ORDER = ["ELET", "INST", "MEC"]

    DATE_PRESETS = {
        "Próximos 30 dias": 30,
        "Próximos 90 dias": 90,
        "Próximos 6 meses": 180,
        "Próximo ano": 365,
        "Personalizado": None,
    }

    PALETTE = {
        "Estaleiro": "#2563EB",
        "Embarque": "#2563EB",
        "Treinamento": "#2563EB",
        "Folga": "#6B7280",
        "Férias": "#6B7280",
        "Workshop": "#14B8A6",
        "Visita Técnica": "#EC4899",
    }

    COLOR_TODAY = "#DC2626"
    COLOR_CONFLICT = "#DC2626"
    COLOR_WEEKEND = "rgba(100,116,139,0.10)"
    COLOR_PERIOD = "rgba(59,130,246,0.06)"
    COLOR_GRID = "rgba(148,163,184,0.35)"
    COLOR_GRID_MINOR = "rgba(226,232,240,0.55)"
    COLOR_DIVIDER = "rgba(15,23,42,0.35)"

    HOVER_TEMPLATE = (
        "<b>%{customdata[3]}</b><br>"
        "Disciplina: %{customdata[0]}<br>"
        "Projeto: %{customdata[1]}<br>"
        "Detalhe: %{customdata[2]}<br>"
        "%{customdata[4]} → %{customdata[5]} · %{customdata[6]} dias"
        "<extra></extra>"
    )

    @classmethod
    def palette_for(cls, tipo: str) -> str:
        return cls.PALETTE.get(tipo, "#94A3B8")


# --- CARREGAMENTO DE DADOS --------------------------------------------------
class DataLoader:
    """Carregamento e normalização de dados a partir das planilhas Excel."""

    @staticmethod
    def _normalize_dates(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
        for col in cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
        return df

    @staticmethod
    def _to_matricula(series: pd.Series) -> pd.Series:
        numeric = pd.to_numeric(series, errors="coerce")
        return numeric.apply(lambda x: str(int(x)) if pd.notna(x) else pd.NA)

    @staticmethod
    def _download_from_drive(file_id: str, dest_path: str) -> bool:
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
            st.error(f"Erro ao baixar do Google Drive ({file_id}): {e}")
            return False

    @staticmethod
    def _is_valid_xlsx(path: str) -> bool:
        try:
            with open(path, "rb") as f:
                return f.read(4)[:2] == b"PK"
        except Exception:
            return False

    @staticmethod
    def _ensure_drive_file(filename: str, local_path: str) -> bool:
        drive_id = Config.GOOGLE_DRIVE_IDS.get(filename)
        if not drive_id:
            return os.path.exists(local_path)
        if DataLoader._download_from_drive(drive_id, local_path):
            return True
        return os.path.exists(local_path)


    @staticmethod
    @st.cache_data(ttl=3600, show_spinner="Baixando Planejamento Estaleiro do Google Drive...")
    def load_estaleiro_data() -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        if not DataLoader._ensure_drive_file("Planejamento Estaleiro.xlsx", Config.FILE_PATH_ESTALEIRO):
            st.error("Não foi possível obter o Planejamento Estaleiro (nem do Drive nem local).")
            return None, None
        try:
            equipe_df = pd.read_excel(Config.FILE_PATH_ESTALEIRO, sheet_name="Equipe")
            keep = ["Disciplina", "Matrícula", "Função", "Projeto", "Experiência", "Nome", "E-mail"]
            keep = [c for c in keep if c in equipe_df.columns]
            equipe_df = equipe_df[keep].copy()

            equipe_df = equipe_df.dropna(subset=["Nome", "Matrícula"])
            equipe_df["Matrícula"] = DataLoader._to_matricula(equipe_df["Matrícula"])
            equipe_df = equipe_df.dropna(subset=["Matrícula"])

            for c in ["Disciplina", "Função", "Projeto", "Experiência", "E-mail"]:
                if c in equipe_df.columns:
                    equipe_df[c] = equipe_df[c].astype("string")
            equipe_df["Nome"] = equipe_df["Nome"].astype("string").str.strip().str.upper()
            for c in ["Disciplina", "Função", "Projeto", "Experiência"]:
                if c in equipe_df.columns:
                    equipe_df[c] = equipe_df[c].str.strip().str.upper()

            plan_df = pd.read_excel(
                Config.FILE_PATH_ESTALEIRO,
                sheet_name="Planejamento IED",
                skiprows=8,
                usecols="C:E",
            )
            plan_df.columns = ["Nome", "Início", "Término"]
            plan_df = plan_df.dropna(subset=["Nome"])
            plan_df["Nome"] = plan_df["Nome"].astype("string").str.strip().str.upper()
            plan_df = DataLoader._normalize_dates(plan_df, ["Início", "Término"])

            plan_df = plan_df.merge(
                equipe_df[["Nome", "Matrícula", "Disciplina", "Função", "Projeto"]],
                on="Nome",
                how="left",
            )
            plan_df["Matrícula"] = DataLoader._to_matricula(plan_df["Matrícula"])
            plan_df["Tipo"] = "Estaleiro"

            return equipe_df, plan_df
        except Exception as e:
            st.error(f"Erro ao carregar o Planejamento Estaleiro: {e}")
            return None, None

    @staticmethod
    @st.cache_data(ttl=3600, show_spinner="Carregando Férias...")
    def load_ferias_data() -> Optional[pd.DataFrame]:
        try:
            df = pd.read_excel(Config.FILE_PATH_FERIAS)
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
                frames.append(tmp)

            if not frames:
                ferias = pd.DataFrame(columns=["Matrícula", "Início", "Término", "Tipo"])
            else:
                ferias = pd.concat(frames, ignore_index=True)
                ferias["Matrícula"] = DataLoader._to_matricula(ferias["Matrícula"])
                ferias = ferias.dropna(subset=["Matrícula"])
                ferias = DataLoader._normalize_dates(ferias, ["Início", "Término"])
                ferias["Tipo"] = "Férias"
            return ferias
        except Exception as e:
            st.error(f"Erro ao carregar o arquivo de Férias: {e}")
            return None

    @staticmethod
    @st.cache_data(ttl=3600, show_spinner="Baixando Planejamento Geral do Google Drive...")
    def load_geral_equipe_data() -> Optional[pd.DataFrame]:
        if not DataLoader._ensure_drive_file("Planejamento Geral.xlsx", Config.FILE_PATH_GERAL):
            st.error("Não foi possível obter o Planejamento Geral (nem do Drive nem local).")
            return None
        try:
            equipe_df = pd.read_excel(Config.FILE_PATH_GERAL, sheet_name="Equipe")
            keep = ["Disciplina", "Matrícula", "Função", "Projeto", "Experiência", "Nome", "E-mail"]
            keep = [c for c in keep if c in equipe_df.columns]
            equipe_df = equipe_df[keep].copy()

            equipe_df = equipe_df.dropna(subset=["Nome", "Matrícula"])
            equipe_df["Matrícula"] = DataLoader._to_matricula(equipe_df["Matrícula"])
            equipe_df = equipe_df.dropna(subset=["Matrícula"])

            for c in ["Disciplina", "Função", "Projeto", "Experiência", "E-mail"]:
                if c in equipe_df.columns:
                    equipe_df[c] = equipe_df[c].astype("string")
            equipe_df["Nome"] = equipe_df["Nome"].astype("string").str.strip().str.upper()
            for c in ["Disciplina", "Função", "Projeto", "Experiência"]:
                if c in equipe_df.columns:
                    equipe_df[c] = equipe_df[c].str.strip().str.upper()

            return equipe_df
        except Exception as e:
            st.error(f"Erro ao carregar a aba Equipe do Planejamento Geral: {e}")
            return None

    @staticmethod
    @st.cache_data(ttl=3600, show_spinner="Baixando Planejamento Geral do Google Drive...")
    def load_planejamento_geral() -> Optional[pd.DataFrame]:
        if not DataLoader._ensure_drive_file("Planejamento Geral.xlsx", Config.FILE_PATH_GERAL):
            st.error("Não foi possível obter o Planejamento Geral (nem do Drive nem local).")
            return None
        try:
            cols = ["Nome", "Matrícula", "Início", "Término", "Atividade", "Detalhamento", "Plataforma", "Experiência"]
            df = pd.read_excel(Config.FILE_PATH_GERAL, usecols=cols)
            df = df.rename(columns={"Atividade": "Tipo"})
            df["Nome"] = df["Nome"].astype("string").str.strip().str.upper()
            df["Matrícula"] = DataLoader._to_matricula(df["Matrícula"])
            df = df.dropna(subset=["Matrícula", "Início"])
            df = DataLoader._normalize_dates(df, ["Início", "Término"])
            df["Tipo"] = df["Tipo"].astype("string").str.strip().str.title()
            valid_types = {"Estaleiro", "Embarque", "Folga", "Férias", "Treinamento", "Workshop", "Visita Técnica", "Atestado"}
            df = df[df["Tipo"].isin(valid_types)].reset_index(drop=True)
            return df
        except Exception as e:
            st.error(f"Erro ao carregar o Planejamento Geral: {e}")
            return None


# --- PROCESSAMENTO ----------------------------------------------------------
class DataProcessor:
    """Combinação de fontes, cálculo de disponibilidade e detecção de conflitos."""

    @staticmethod
    def _mat_set(series: pd.Series) -> Set:
        return set(series.dropna().astype(str))

    @staticmethod
    def prepare_combined_data(
        equipe_df: pd.DataFrame,
        dfs_eventos: List[Optional[pd.DataFrame]],
    ) -> Tuple[pd.DataFrame, pd.DataFrame]:
        valid_dfs = []
        matriculas_validas = DataProcessor._mat_set(equipe_df["Matrícula"])

        for df in dfs_eventos:
            if df is None or df.empty:
                continue
            df = df[df["Matrícula"].astype(str).isin(matriculas_validas)].copy()
            cols = ["Matrícula", "Início", "Término", "Tipo"]
            for extra in ["Nome", "Detalhamento", "Plataforma"]:
                if extra in df.columns:
                    cols.append(extra)
            valid_dfs.append(df[cols])

        if not valid_dfs:
            combined = pd.DataFrame(columns=["Matrícula", "Nome", "Início", "Término", "Disciplina", "Tipo"])
        else:
            combined = pd.concat(valid_dfs, ignore_index=True)

        for c in ["Detalhamento", "Plataforma"]:
            if c not in combined.columns:
                combined[c] = pd.NA

        combined = combined.dropna(subset=["Início", "Término"])
        combined["Tipo"] = combined["Tipo"].astype("string").str.strip().str.title()
        valid_types = {"Estaleiro", "Embarque", "Folga", "Férias", "Treinamento", "Workshop", "Visita Técnica", "Atestado"}
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
    ) -> Set:
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
    ) -> Set:
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
        if combined_df.empty:
            return combined_df
        mask = (combined_df["Início"] <= end_date) & (combined_df["Término"] >= start_date)
        return combined_df.loc[mask].copy()

    @staticmethod
    def detect_conflicts(combined_df: pd.DataFrame) -> Tuple[pd.DataFrame, Set]:
        if combined_df.empty:
            return pd.DataFrame(), set()

        df = combined_df.dropna(subset=["Início", "Término"]).copy()
        rows = []
        keys: Set = set()

        for mat, group in df.groupby("Matrícula"):
            group = group.sort_values("Início").reset_index(drop=True)
            values = group.to_dict("records")
            for i in range(len(values)):
                a = values[i]
                for j in range(i + 1, len(values)):
                    b = values[j]
                    if a["Início"] <= b["Término"] and b["Início"] <= a["Término"]:
                        overlap = min(a["Término"], b["Término"]) - max(a["Início"], b["Início"])
                        overlap_days = overlap.days + 1
                        rows.append({
                            "Nome": a["Nome"],
                            "Matrícula": mat,
                            "Atividade A": a["Tipo"],
                            "Período A": f"{a['Início']:%d/%m/%Y} – {a['Término']:%d/%m/%Y}",
                            "Início A": a["Início"],
                            "Atividade B": b["Tipo"],
                            "Período B": f"{b['Início']:%d/%m/%Y} – {b['Término']:%d/%m/%Y}",
                            "Início B": b["Início"],
                            "Sobreposição (dias)": overlap_days,
                        })
                        keys.add((int(mat), a["Início"], a["Tipo"]))
                        keys.add((int(mat), b["Início"], b["Tipo"]))
                    elif b["Início"] > a["Término"]:
                        break

        conflicts = pd.DataFrame(
            rows,
            columns=[
                "Nome", "Matrícula", "Atividade A", "Período A", "Início A",
                "Atividade B", "Período B", "Início B", "Sobreposição (dias)",
            ],
        )
        display_cols = ["Nome", "Atividade A", "Período A", "Atividade B", "Período B", "Sobreposição (dias)"]
        return conflicts, keys


# --- VISUALIZAÇÃO -----------------------------------------------------------
class Visualizer:
    @staticmethod
    def create_gantt_chart(
        chart_df: pd.DataFrame,
        unique_members: pd.DataFrame,
        start_date: pd.Timestamp,
        end_date: pd.Timestamp,
        conflict_keys: Set,
        conflict_details: Dict,
        y_ticktext: List[str],
    ) -> go.Figure:
        names = unique_members["Nome"].astype(str).tolist()
        n = len(names)
        fig = go.Figure()

        if not chart_df.empty:
            chart_df = chart_df.dropna(subset=["Início", "Término"]).copy()
            det = chart_df["Detalhamento"] if "Detalhamento" in chart_df.columns else pd.Series(pd.NA, index=chart_df.index)
            if "Plataforma" in chart_df.columns:
                det = det.where(det.notna(), chart_df["Plataforma"])
            det = det.fillna("—").astype(str)

            for tipo in Config.PALETTE:
                g = chart_df[chart_df["Tipo"] == tipo]
                if g.empty:
                    continue
                g = g.sort_values("Nome")
                starts = g["Início"]
                ends = g["Término"]
                start_ms = (starts.astype("int64") // 1_000_000).astype("int64")
                end_ms = ((ends + pd.Timedelta(days=1)).astype("int64") // 1_000_000).astype("int64")
                x_ms = (end_ms - start_ms)
                dias = (ends - starts).dt.days + 1

                custom = pd.DataFrame({
                    "Disciplina": g["Disciplina"].fillna("—").astype(str),
                    "Projeto": g["Projeto"].fillna("—").astype(str),
                    "Detalhe": det.loc[g.index],
                    "Tipo": [tipo] * len(g),
                    "ini": starts.dt.strftime("%d/%m/%Y"),
                    "fim": ends.dt.strftime("%d/%m/%Y"),
                    "dias": dias.astype(int).astype(str),
                })

                fig.add_trace(go.Bar(
                    y=g["Nome"].astype(str).tolist(),
                    x=x_ms.tolist(),
                    base=start_ms.tolist(),
                    orientation="h",
                    name=tipo,
                    marker=dict(
                        color=Config.palette_for(tipo),
                        line=dict(width=1, color="rgba(15,23,42,0.35)"),
                    ),
                    customdata=custom.values.tolist(),
                    hovertemplate=Config.HOVER_TEMPLATE,
                    showlegend=True,
                ))

        if not chart_df.empty and conflict_keys:
            conflict_mask = []
            for _, r in chart_df.iterrows():
                key = (int(r["Matrícula"]) if pd.notna(r["Matrícula"]) else None, pd.Timestamp(r["Início"]), r["Tipo"])
                conflict_mask.append(key in conflict_keys)
            chart_df_conf = chart_df[conflict_mask].copy()
            if not chart_df_conf.empty:
                starts_c = chart_df_conf["Início"]
                ends_c = chart_df_conf["Término"]
                start_ms_c = (starts_c.astype("int64") // 1_000_000).astype("int64")
                end_ms_c = ((ends_c + pd.Timedelta(days=1)).astype("int64") // 1_000_000).astype("int64")
                x_ms_c = (end_ms_c - start_ms_c)
                custom_conf = []
                for _, r in chart_df_conf.iterrows():
                    key = (int(r["Matrícula"]) if pd.notna(r["Matrícula"]) else None, pd.Timestamp(r["Início"]), r["Tipo"])
                    details = conflict_details.get(key, ["Conflito"])
                    custom_conf.append(["\n".join(details)])
                fig.add_trace(go.Bar(
                    y=chart_df_conf["Nome"].astype(str).tolist(),
                    x=x_ms_c.tolist(),
                    base=start_ms_c.tolist(),
                    orientation="h",
                    name="Conflito",
                    marker=dict(color="rgba(220,38,38,0.25)", line=dict(width=0)),
                    customdata=custom_conf,
                    hovertemplate="<b>Conflito(s)</b><br>%{customdata[0]}<extra></extra>",
                    showlegend=True,
                ))

        if len(fig.data) == 0:
            fig.add_trace(go.Scatter(
                y=list(range(len(names))),
                x=[0] * len(names),
                mode="markers",
                marker=dict(opacity=0, size=0),
                showlegend=False,
                hoverinfo="skip",
                hovertemplate=None,
            ))

        x0 = start_date - pd.Timedelta(days=1)
        x1 = end_date + pd.Timedelta(days=4)

        fig.add_shape(
            type="rect", xref="x", yref="paper",
            x0=start_date, x1=end_date + pd.Timedelta(days=1),
            y0=0, y1=1, fillcolor=Config.COLOR_PERIOD, line_width=0, layer="below",
        )

        d = (x0.normalize()).to_pydatetime()
        end_d = x1.to_pydatetime()
        while d <= end_d:
            if d.weekday() == 5:
                fig.add_shape(
                    type="rect", xref="x", yref="paper",
                    x0=d, x1=d + timedelta(days=2),
                    y0=0, y1=1, fillcolor=Config.COLOR_WEEKEND, line_width=0, layer="below",
                )
                d += timedelta(days=2)
            else:
                d += timedelta(days=1)

        hoje = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        fig.add_vline(x=hoje, line=dict(color=Config.COLOR_TODAY, width=2, dash="dash"))
        fig.add_annotation(
            x=hoje, y=1.0, xref="x", yref="paper",
            text="Hoje", showarrow=False, yanchor="bottom",
            font=dict(color=Config.COLOR_TODAY, size=11),
        )

        discs = unique_members["Disciplina"].astype("object").tolist()
        prev = None
        first_idx = 0
        for i, disc in enumerate(discs):
            if prev is not None and disc != prev:
                fig.add_hline(y=i - 0.5, line=dict(color=Config.COLOR_DIVIDER, width=1, dash="dot"))
                center = (first_idx + i - 1) / 2
                fig.add_annotation(
                    x=1, y=center, xref="paper", yref="y", text=f"<b>{prev}</b>",
                    showarrow=False, xanchor="right", font=dict(size=11, color="#334155"),
                )
                first_idx = i
            if prev is None:
                first_idx = i
            prev = disc
        if prev is not None:
            center = (first_idx + len(discs) - 1) / 2
            fig.add_annotation(
                x=1, y=center, xref="paper", yref="y", text=f"<b>{prev}</b>",
                showarrow=False, xanchor="right", font=dict(size=11, color="#334155"),
            )

        cur = start_date.normalize()
        end_norm = end_date.normalize()
        while cur <= end_norm:
            if cur.weekday() == 0:
                fig.add_vline(
                    x=cur,
                    line=dict(color="rgba(148,163,184,0.25)", width=1, dash="dot"),
                    layer="below",
                )
                fig.add_annotation(
                    x=cur, y=1.0, xref="x", yref="paper",
                    text=cur.strftime("%d/%m/%Y"),
                    showarrow=False,
                    xanchor="center", yanchor="bottom",
                    textangle=-45,
                    font=dict(size=10, color="#475569"),
                    yshift=4,
                )
            cur += timedelta(days=1)

        fig.update_layout(
            height=max(650, n * 32),
            barmode="overlay",
            bargap=0.25,
            bargroupgap=0,
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(l=8, r=8, t=60, b=50),
            title=dict(text="Cronograma de Alocação", x=0.01, xanchor="left", font=dict(size=16, color="#0f172a")),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title="Atividade"),
            hovermode="closest",
            hoverlabel=dict(bgcolor="white", font_size=12, bordercolor="#cbd5e1"),
            xaxis=dict(
                type="date",
                range=[x0, x1],
                rangeslider=dict(visible=True, thickness=0.05),
                dtick="M1",
                tickformat="%b\n%Y",
                showgrid=True,
                gridcolor=Config.COLOR_GRID,
                minor=dict(showgrid=True, dtick=86400000, gridcolor=Config.COLOR_GRID_MINOR),
                zeroline=False,
            ),
            yaxis=dict(
                categoryorder="array",
                categoryarray=names,
                range=[len(names) - 0.5, -0.5],
                tickmode="array",
                tickvals=list(range(len(names))),
                ticktext=y_ticktext,
                dtick=1,
                showgrid=True,
                gridcolor="rgba(226,232,240,0.8)",
                zeroline=False,
                tickfont=dict(size=11),
            ),
        )
        return fig


# --- APLICAÇÃO PRINCIPAL ----------------------------------------------------
class App:
    def __init__(self):
        st.set_page_config(
            layout="wide",
            page_title="Gestão de Alocação",
            page_icon="📊",
            initial_sidebar_state="expanded",
        )
        self.hoje = datetime.today().date()

    def run(self):
        st.title("Painel de Disponibilidade de Equipe")
        st.caption("Visualização integrada de Estaleiro, Férias e Planejamento Geral.")

        with st.sidebar:
            st.header("Filtros")
            if st.button("🔄 Atualizar dados", use_container_width=True, help="Recarrega as planilhas ignorando o cache"):
                st.cache_data.clear()
                st.rerun()

            equipe_df = DataLoader.load_geral_equipe_data()
            _, plan_df = DataLoader.load_estaleiro_data()
            ferias_df = DataLoader.load_ferias_data()
            geral_df = DataLoader.load_planejamento_geral()

            if equipe_df is None:
                st.error("Falha ao carregar a aba Equipe do Planejamento Geral.")
                return

            all_discs = sorted([x for x in equipe_df["Disciplina"].dropna().unique()])
            sel_discs = st.multiselect("Disciplina", all_discs, default=all_discs)
            all_projs = sorted([x for x in equipe_df["Projeto"].dropna().unique()])
            sel_projs = st.multiselect("Projeto", all_projs, default=all_projs)

            funcao_map = {"SUPERVISOR": "LIDERANÇA", "COORDENADOR": "LIDERANÇA"}
            equipe_df["FunçãoAgrupada"] = equipe_df["Função"].map(
                lambda f: funcao_map.get(f, f) if pd.notna(f) else f
            )
            all_funcoes = sorted([x for x in equipe_df["FunçãoAgrupada"].dropna().unique()])
            sel_funcoes = st.multiselect("Função", all_funcoes, default=all_funcoes)

            preset = st.selectbox("Período", list(Config.DATE_PRESETS.keys()), index=1)
            dias = Config.DATE_PRESETS[preset]
            if dias is not None:
                d_inicio = pd.Timestamp(self.hoje - timedelta(days=7))
                d_fim = pd.Timestamp(self.hoje + timedelta(days=dias))
                st.caption(f"{d_inicio:%d/%m/%Y} a {d_fim:%d/%m/%Y}")
            else:
                col1, col2 = st.columns(2)
                d_inicio = pd.Timestamp(col1.date_input("Início", self.hoje - timedelta(days=7)))
                d_fim = pd.Timestamp(col2.date_input("Fim", self.hoje + timedelta(days=90)))

            if d_inicio > d_fim:
                st.sidebar.warning("Data de início posterior à data de fim. Invertendo...")
                d_inicio, d_fim = d_fim, d_inicio

            only_available = st.checkbox(
                "Apenas disponíveis no período",
                help="Mostra apenas quem não tem nenhum evento agendado dentro do período selecionado",
            )

        equipe_filtered = equipe_df[
            equipe_df["Disciplina"].isin(sel_discs)
            & equipe_df["Projeto"].isin(sel_projs)
            & equipe_df["FunçãoAgrupada"].isin(sel_funcoes)
        ].copy()

        if equipe_filtered.empty:
            st.warning("Nenhum colaborador encontrado com os filtros atuais.")
            return

        combined_df, unique_members = DataProcessor.prepare_combined_data(
            equipe_filtered, [plan_df, ferias_df, geral_df]
        )

        occupied_ids = DataProcessor.occupied_in_period(combined_df, d_inicio, d_fim)
        familiar_available = DataProcessor.available_in_period(unique_members, combined_df, d_inicio, d_fim)
        events_in_window = DataProcessor.events_in_period(combined_df, d_inicio, d_fim)
        conflicts_full, conflict_keys = DataProcessor.detect_conflicts(combined_df)
        conflicts_df = conflicts_full[["Nome", "Atividade A", "Período A", "Atividade B", "Período B", "Sobreposição (dias)"]] if not conflicts_full.empty else conflicts_full

        conflict_details = {}
        if not conflicts_full.empty:
            for _, row in conflicts_full.iterrows():
                key_a = (int(row["Matrícula"]), pd.Timestamp(row["Início A"]), row["Atividade A"])
                key_b = (int(row["Matrícula"]), pd.Timestamp(row["Início B"]), row["Atividade B"])
                desc = f"{row['Atividade A']} ({row['Período A']}) x {row['Atividade B']} ({row['Período B']})"
                conflict_details.setdefault(key_a, []).append(desc)
                conflict_details.setdefault(key_b, []).append(desc)

        n_total = len(unique_members)
        n_available = len(familiar_available)
        n_occupied = n_total - n_available

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Equipe no filtro", n_total)
        k2.metric("Disponíveis", n_available)
        k3.metric("Ocupados no período", n_occupied)
        k4.metric("Eventos no período", len(events_in_window))
        k5.metric("Conflitos", len(conflicts_df))

        display_members = unique_members.copy()
        if only_available:
            display_members = display_members[
                display_members["Matrícula"].astype(str).isin(familiar_available)
            ].reset_index(drop=True)

        if display_members.empty:
            st.info("Nenhum colaborador disponível para os critérios selecionados.")
            return

        chart_df = combined_df[combined_df["Matrícula"].isin(set(display_members["Matrícula"]))].copy()
        chart_df = DataProcessor.events_in_period(chart_df, d_inicio, d_fim) if not chart_df.empty else chart_df

        tab_crono, tab_conf, tab_equipe, tab_export = st.tabs(
            ["Cronograma", "Conflitos", "Equipe", "Exportar"]
        )

        with tab_crono:
            y_ticktext = []
            for _, row in display_members.iterrows():
                nome = str(row["Nome"])
                mat_str = str(int(row["Matrícula"])) if pd.notna(row["Matrícula"]) else None
                if mat_str is not None and mat_str in familiar_available:
                    y_ticktext.append(f'<span style="color:#10B981;font-weight:bold">{nome}</span>')
                else:
                    y_ticktext.append(f'<span style="color:#EF4444;font-weight:bold">{nome}</span>')
            fig = Visualizer.create_gantt_chart(chart_df, display_members, d_inicio, d_fim, conflict_keys, conflict_details, y_ticktext)
            st.plotly_chart(fig, use_container_width=True)
            st.caption("Verde = disponível no período | Vermelho = ocupado | Barras vermelhas sobrepostas indicam conflitos.")

        with tab_conf:
            if conflicts_df.empty:
                st.success("✅ Nenhum conflito de agendamento detectado.")
            else:
                st.warning(f"⚠️ {len(conflicts_df)} conflito(s) encontrado(s).")
                st.dataframe(conflicts_df, use_container_width=True, hide_index=True)

        with tab_equipe:
            status_rows = display_members.copy()
            status_rows["Status no período"] = status_rows["Matrícula"].map(
                lambda m: "Disponível" if str(int(m)) in familiar_available else "Ocupado" if pd.notna(m) else "—"
            )
            show_cols = [c for c in [
                "Nome", "Disciplina", "Função", "Projeto", "E-mail", "Matrícula", "Status no período"
            ] if c in status_rows.columns]
            st.dataframe(status_rows[show_cols], use_container_width=True, hide_index=True)

        with tab_export:
            st.subheader("Eventos no período")
            st.dataframe(events_in_window, use_container_width=True, hide_index=True) if not events_in_window.empty else st.info("Sem eventos no período.")
            st.divider()
            c_exp1, c_exp2, c_exp3 = st.columns(3)
            if not events_in_window.empty:
                c_exp1.download_button("📥 Eventos (CSV)", events_in_window.to_csv(index=False).encode("utf-8"), "eventos.csv", "text/csv", use_container_width=True)
            if not conflicts_df.empty:
                c_exp2.download_button("📥 Conflitos (CSV)", conflicts_df.to_csv(index=False).encode("utf-8"), "conflitos.csv", "text/csv", use_container_width=True)
            equipe_csv = equipe_filtered.to_csv(index=False).encode("utf-8")
            c_exp3.download_button("📥 Equipe (CSV)", equipe_csv, "equipe.csv", "text/csv", use_container_width=True)


if __name__ == "__main__":
    App().run()
