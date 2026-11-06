"""
Renderização de gráficos interativos Gantt com Plotly (v0.3).
"""

from datetime import datetime, timedelta
from typing import Dict, List, Optional, Set
import pandas as pd
import plotly.graph_objects as go

from .config import Config


class Visualizer:
    """Geração de gráficos de Gantt interativos e responsivos para o cronograma de equipe."""

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
        """
        Constrói o gráfico de Gantt com suporte a sobreposições, realce de conflitos,
        sombreamento de finais de semana, agrupamento dinâmico por disciplina e linha de hoje robusta.
        """
        names = unique_members["Nome"].astype(str).tolist() if not unique_members.empty else []
        num_colaboradores = len(names)
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
                start_ms = starts.astype("datetime64[ms]").astype("int64")
                end_ms = (ends + pd.Timedelta(days=1)).astype("datetime64[ms]").astype("int64")
                x_ms = (end_ms - start_ms)
                dias = (ends - starts).dt.days + 1

                # Customdata estruturado: Nome, Matrícula, Disciplina, Tipo, Detalhe, Início, Término, Duração
                custom = pd.DataFrame({
                    "Nome": g["Nome"].fillna("—").astype(str),
                    "Matrícula": g["Matrícula"].fillna("—").astype(str),
                    "Disciplina": g["Disciplina"].fillna("—").astype(str),
                    "Tipo": [tipo] * len(g),
                    "Detalhe": det.loc[g.index],
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

        # Realce de eventos em conflito
        if not chart_df.empty and conflict_keys:
            conflict_mask = []
            for _, r in chart_df.iterrows():
                try:
                    mat_val = int(r["Matrícula"]) if pd.notna(r["Matrícula"]) else None
                except (ValueError, TypeError):
                    mat_val = r["Matrícula"]
                key = (mat_val, pd.Timestamp(r["Início"]), r["Tipo"])
                conflict_mask.append(key in conflict_keys)

            chart_df_conf = chart_df[conflict_mask].copy()
            if not chart_df_conf.empty:
                starts_c = chart_df_conf["Início"]
                ends_c = chart_df_conf["Término"]
                start_ms_c = starts_c.astype("datetime64[ms]").astype("int64")
                end_ms_c = (ends_c + pd.Timedelta(days=1)).astype("datetime64[ms]").astype("int64")
                x_ms_c = (end_ms_c - start_ms_c)

                custom_conf = []
                for _, r in chart_df_conf.iterrows():
                    try:
                        mat_val = int(r["Matrícula"]) if pd.notna(r["Matrícula"]) else None
                    except (ValueError, TypeError):
                        mat_val = r["Matrícula"]
                    key = (mat_val, pd.Timestamp(r["Início"]), r["Tipo"])
                    details = conflict_details.get(key, ["Conflito"])
                    custom_conf.append(["<br>".join(details)])

                fig.add_trace(go.Bar(
                    y=chart_df_conf["Nome"].astype(str).tolist(),
                    x=x_ms_c.tolist(),
                    base=start_ms_c.tolist(),
                    orientation="h",
                    name="Conflito",
                    marker=dict(color="rgba(220,38,38,0.25)", line=dict(width=0)),
                    customdata=custom_conf,
                    hovertemplate="<b>⚠️ Conflito(s) Detectado(s)</b><br>%{customdata[0]}<extra></extra>",
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

        # Retângulo de destaque do período ativo
        fig.add_shape(
            type="rect", xref="x", yref="paper",
            x0=start_date, x1=end_date + pd.Timedelta(days=1),
            y0=0, y1=1, fillcolor=Config.COLOR_PERIOD, line_width=0, layer="below",
        )

        # Sombreamento de fins de semana (sábado e domingo)
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

        # --- CORREÇÃO DEFINITIVA DA LINHA "HOJE" ---
        hoje_dt = pd.Timestamp(datetime.now().date())
        fig.add_vline(x=hoje_dt, line_width=2, line_dash="dash", line_color="red")
        fig.add_annotation(
            x=hoje_dt,
            y=1.02,
            yref="paper",
            text="Hoje",
            showarrow=False,
            font=dict(color="red", size=11, family="sans-serif"),
            yanchor="bottom",
        )

        # --- AGRUPAMENTO DINÂMICO POR DISCIPLINA ---
        if not unique_members.empty and "Disciplina" in unique_members.columns:
            discs = unique_members["Disciplina"].fillna("OUTROS").astype(str).tolist()
            prev = None
            first_idx = 0
            for i, disc in enumerate(discs):
                if prev is not None and disc != prev:
                    fig.add_hline(y=i - 0.5, line=dict(color=Config.COLOR_DIVIDER, width=1, dash="dot"))
                    center = (first_idx + i - 1) / 2
                    fig.add_annotation(
                        x=1.0, y=center, xref="paper", yref="y", text=f"<b>{prev}</b>",
                        showarrow=False, xanchor="right", font=dict(size=11, color="#334155"),
                    )
                    first_idx = i
                if prev is None:
                    first_idx = i
                prev = disc
            if prev is not None:
                center = (first_idx + len(discs) - 1) / 2
                fig.add_annotation(
                    x=1.0, y=center, xref="paper", yref="y", text=f"<b>{prev}</b>",
                    showarrow=False, xanchor="right", font=dict(size=11, color="#334155"),
                )

        # Marcadores verticais nas segundas-feiras
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

        # --- ALTURA DINÂMICA DO GRÁFICO (max(450, num_colaboradores * 38 + 150)) ---
        dynamic_height = max(450, num_colaboradores * 38 + 150)

        fig.update_layout(
            height=dynamic_height,
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
                range=[len(names) - 0.5, -0.5] if names else [0, 1],
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
