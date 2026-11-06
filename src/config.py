"""
Configurações, constantes e parâmetros globais da aplicação Gantt & Escalas (v0.3).
"""

from typing import Dict, List, Optional


VERSION = "v0.3"


class Config:
    """Configurações centrais, caminhos de arquivo, paleta de cores e presets."""

    # Caminhos locais padrão dos arquivos Excel
    FILE_PATH_ESTALEIRO: str = "Planejamento Estaleiro.xlsx"
    FILE_PATH_FERIAS: str = "Férias.xlsx"
    FILE_PATH_GERAL: str = "Planejamento Geral.xlsx"

    # IDs para download automático do Google Drive quando necessário
    GOOGLE_DRIVE_IDS: Dict[str, str] = {
        "Planejamento Estaleiro.xlsx": "1VYvNJV9V4vUYeQgCA7DjfZY0ggYOpbP3",
        "Planejamento Geral.xlsx": "1NSP2p69F33kqE_FLvsbfugLjDc2v-fS5",
    }

    # Ordem canônica de exibição das disciplinas no gráfico e listas
    DISCIPLINA_ORDER: List[str] = ["ELET", "INST", "MEC"]

    # Presets de horizonte de visualização temporal (em dias)
    DATE_PRESETS: Dict[str, Optional[int]] = {
        "Próximos 30 dias": 30,
        "Próximos 90 dias": 90,
        "Próximos 6 meses": 180,
        "Próximo ano": 365,
        "Personalizado": None,
    }

    # Mapeamento flexível de aliases de colunas para o padrão interno canônico
    COLUMN_ALIASES: Dict[str, str] = {
        # Matrícula
        "matricula": "Matrícula",
        "matrícula": "Matrícula",
        "mat": "Matrícula",
        "id": "Matrícula",
        "código": "Matrícula",
        "codigo": "Matrícula",
        "registro": "Matrícula",
        # Nome
        "nome": "Nome",
        "colaborador": "Nome",
        "colaboradora": "Nome",
        "funcionario": "Nome",
        "funcionário": "Nome",
        "recurso": "Nome",
        # Disciplina
        "disciplina": "Disciplina",
        "disc": "Disciplina",
        "área": "Disciplina",
        "area": "Disciplina",
        # Início
        "início": "Início",
        "inicio": "Início",
        "data inicio": "Início",
        "data início": "Início",
        "data_inicio": "Início",
        "data_início": "Início",
        "dt_inicio": "Início",
        "dt inicio": "Início",
        "start": "Início",
        "start date": "Início",
        "start_date": "Início",
        # Término
        "término": "Término",
        "termino": "Término",
        "data termino": "Término",
        "data término": "Término",
        "data_termino": "Término",
        "data_término": "Término",
        "dt_termino": "Término",
        "dt termino": "Término",
        "fim": "Término",
        "data fim": "Término",
        "end": "Término",
        "end date": "Término",
        "end_date": "Término",
        # Tipo / Atividade
        "tipo": "Tipo",
        "atividade": "Tipo",
        "tipo atividade": "Tipo",
        "tipo_atividade": "Tipo",
        "evento": "Tipo",
        "status": "Tipo",
        # Detalhamento
        "detalhamento": "Detalhamento",
        "detalhe": "Detalhamento",
        "observação": "Detalhamento",
        "observacao": "Detalhamento",
        "obs": "Detalhamento",
        "descrição": "Detalhamento",
        "descricao": "Detalhamento",
        # Campos de Equipe / Suporte
        "função": "Função",
        "funcao": "Função",
        "cargo": "Função",
        "projeto": "Projeto",
        "experiência": "Experiência",
        "experiencia": "Experiência",
        "senioridade": "Experiência",
        "e-mail": "E-mail",
        "email": "E-mail",
        "plataforma": "Plataforma",
    }

    # Severidade de conflitos
    SEVERITY_CRITICO: str = "Crítico"
    SEVERITY_OPERACIONAL: str = "Operacional"

    # Mapeamento de cores dos tipos de evento para o gráfico de Gantt
    PALETTE: Dict[str, str] = {
        "Estaleiro": "#2563EB",       # Azul vibrante
        "Embarque": "#2563EB",        # Azul vibrante
        "Treinamento": "#2563EB",     # Azul vibrante
        "Folga": "#6B7280",           # Cinza neutro
        "Férias": "#6B7280",          # Cinza neutro
        "Workshop": "#14B8A6",        # Verde água / Teal
        "Visita Técnica": "#EC4899",  # Rosa
        "Licença": "#9F1239",         # Vinho / Rubi
        "Emprestado": "#4338CA",      # Índigo escuro
    }

    # Cores de suporte para overlays e linhas de grade no Plotly
    COLOR_TODAY: str = "#DC2626"
    COLOR_CONFLICT: str = "#DC2626"
    COLOR_WEEKEND: str = "rgba(100,116,139,0.10)"
    COLOR_PERIOD: str = "rgba(59,130,246,0.06)"
    COLOR_GRID: str = "rgba(148,163,184,0.35)"
    COLOR_GRID_MINOR: str = "rgba(226,232,240,0.55)"
    COLOR_DIVIDER: str = "rgba(15,23,42,0.35)"

    # Template HTML do tooltip hover no gráfico Plotly
    HOVER_TEMPLATE: str = (
        "<b>%{customdata[0]}</b> (Matrícula: %{customdata[1]})<br>"
        "<b>Disciplina:</b> %{customdata[2]} | <b>Tipo:</b> %{customdata[3]}<br>"
        "<b>Detalhamento:</b> %{customdata[4]}<br>"
        "<b>Período:</b> %{customdata[5]} → %{customdata[6]} (%{customdata[7]} dias)"
        "<extra></extra>"
    )

    @classmethod
    def palette_for(cls, tipo: str) -> str:
        """Retorna o código hexadecimal de cor para o tipo de evento, com fallback seguro."""
        return cls.PALETTE.get(tipo, "#94A3B8")
