"""
Pacote central de módulos do painel de disponibilidade e escalas (v0.3).
"""

from .config import Config, VERSION
from .schemas import (
    EQUIPE_COLUMNS,
    PLANEJAMENTO_COLUMNS,
    STATUS_TO_TIPO_MAP,
    VALID_ACTIVITY_TYPES,
    MissingColumnError,
    apply_column_aliases,
    normalizar_projeto,
    normalize_dates,
    normalize_text,
    sanitize_dates,
    to_matricula,
    validate_equipe_schema,
    validate_planejamento_schema,
    validate_required_columns,
)
from .data_loader import DataLoader
from .data_processor import DataProcessor, classify_conflict_severity
from .visualizer import Visualizer
from .ui_components import (
    render_conflicts_tab,
    render_export_tab,
    render_kpi_cards,
    render_sidebar_filters,
    render_team_tab,
)

__all__ = [
    "Config",
    "VERSION",
    "EQUIPE_COLUMNS",
    "PLANEJAMENTO_COLUMNS",
    "STATUS_TO_TIPO_MAP",
    "VALID_ACTIVITY_TYPES",
    "MissingColumnError",
    "apply_column_aliases",
    "normalizar_projeto",
    "normalize_dates",
    "normalize_text",
    "sanitize_dates",
    "to_matricula",
    "validate_equipe_schema",
    "validate_planejamento_schema",
    "validate_required_columns",
    "DataLoader",
    "DataProcessor",
    "classify_conflict_severity",
    "Visualizer",
    "render_conflicts_tab",
    "render_export_tab",
    "render_kpi_cards",
    "render_sidebar_filters",
    "render_team_tab",
]
