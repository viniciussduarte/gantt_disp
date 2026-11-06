"""
Testes unitários para o carregador de dados, esquemas, normalizadores e validação fail-fast (v0.3).
"""

import os
import tempfile
import unittest
import numpy as np
import pandas as pd

from src.config import Config
from src.data_loader import DataLoader
from src.schemas import (
    STATUS_TO_TIPO_MAP,
    VALID_ACTIVITY_TYPES,
    MissingColumnError,
    apply_column_aliases,
    normalize_dates,
    normalize_text,
    normalizar_projeto,
    sanitize_dates,
    to_matricula,
    validate_equipe_schema,
    validate_planejamento_schema,
    validate_required_columns,
)


class TestDataLoaderAndSchemas(unittest.TestCase):
    """Bateria de testes cobrindo DataLoader, esquemas e funções de normalização."""

    def test_to_matricula(self):
        """Verifica conversão segura de matrículas numéricas, float e nulas."""
        s = pd.Series([1234, 5678.0, "9999", "abc", np.nan, None])
        result = to_matricula(s)
        self.assertEqual(result.iloc[0], "1234")
        self.assertEqual(result.iloc[1], "5678")
        self.assertEqual(result.iloc[2], "9999")
        self.assertTrue(pd.isna(result.iloc[3]))
        self.assertTrue(pd.isna(result.iloc[4]))
        self.assertTrue(pd.isna(result.iloc[5]))

    def test_apply_column_aliases(self):
        """Testa o mapeamento flexível de aliases de cabeçalho e remoção de espaços em branco."""
        df = pd.DataFrame(columns=["  matricula ", "NOME  ", "  start date ", " fim ", "atividade", "obs"])
        renamed = apply_column_aliases(df)
        expected = ["Matrícula", "Nome", "Início", "Término", "Tipo", "Detalhamento"]
        for exp in expected:
            self.assertIn(exp, renamed.columns)

    def test_validate_required_columns_fail_fast(self):
        """Testa se colunas faltantes disparam MissingColumnError com mensagem amigável."""
        df = pd.DataFrame(columns=["Nome", "Função"])
        with self.assertRaises(MissingColumnError) as ctx:
            validate_required_columns(df, ["Nome", "Matrícula"], filename="teste_equipe.xlsx")
        self.assertIn("Matrícula", ctx.exception.missing_columns)
        self.assertIn("teste_equipe.xlsx", str(ctx.exception))

    def test_sanitize_dates_dayfirst(self):
        """Verifica a conversão com dayfirst=True e separação de inconsistências."""
        df = pd.DataFrame({
            "Nome": ["A", "B", "C"],
            "Início": ["05/06/2026", "20/06/2026", "invalido"],
            "Término": ["15/06/2026", "10/06/2026", "25/06/2026"],
        })
        valid, inc = sanitize_dates(df, "Início", "Término", filename="datas.xlsx")
        self.assertEqual(len(valid), 1)
        self.assertEqual(valid.iloc[0]["Início"], pd.Timestamp("2026-06-05"))
        self.assertEqual(valid.iloc[0]["Término"], pd.Timestamp("2026-06-15"))

        self.assertEqual(len(inc), 2)
        motivos = inc["Motivo"].tolist()
        self.assertIn("Término anterior ao Início", motivos)
        self.assertIn("Data Inválida ou Ausente (NaT)", motivos)

    def test_normalize_dates(self):
        """Verifica a conversão consistente de strings e timestamps."""
        df = pd.DataFrame({
            "Início": ["2026-05-01", "2026-05-15", "data_invalida"],
            "Término": ["2026-05-14", "2026-05-28", "2026-06-01"],
            "Outro": [1, 2, 3],
        })
        norm = normalize_dates(df, ["Início", "Término", "ColunaInexistente"])
        self.assertEqual(norm["Início"].iloc[0], pd.Timestamp("2026-05-01"))
        self.assertEqual(norm["Término"].iloc[1], pd.Timestamp("2026-05-28"))
        self.assertTrue(pd.isna(norm["Início"].iloc[2]))

    def test_normalize_text(self):
        """Verifica limpeza de espaços e conversão para maiúsculas."""
        s = pd.Series(["  joão silva  ", "Maria oliveira", "   "])
        upper = normalize_text(s, uppercase=True)
        self.assertEqual(upper.iloc[0], "JOÃO SILVA")
        self.assertEqual(upper.iloc[1], "MARIA OLIVEIRA")
        self.assertEqual(upper.iloc[2], "")

    def test_normalizar_projeto(self):
        """Testa a padronização de identificadores de plataformas offshore."""
        self.assertEqual(normalizar_projeto("P80"), "P-80")
        self.assertEqual(normalizar_projeto("p-80"), "P-80")
        self.assertEqual(normalizar_projeto("P 82"), "P-82")
        self.assertEqual(normalizar_projeto("P83"), "P-83")
        self.assertEqual(normalizar_projeto(None), "P-80")
        self.assertEqual(normalizar_projeto("OutroProjeto"), "OutroProjeto")

    def test_validate_equipe_schema(self):
        """Testa a verificação de integridade dos campos obrigatórios da equipe."""
        df_valido = pd.DataFrame({"Nome": ["Silva"], "Matrícula": ["123"], "Função": ["MEC"]})
        valido, missing = validate_equipe_schema(df_valido)
        self.assertTrue(valido)
        self.assertEqual(missing, [])

        df_invalido = pd.DataFrame({"Nome": ["Silva"], "Função": ["MEC"]})
        valido, missing = validate_equipe_schema(df_invalido)
        self.assertFalse(valido)
        self.assertIn("Matrícula", missing)

    def test_validate_planejamento_schema(self):
        """Testa validação de campos temporais obrigatórios de eventos."""
        df_valido = pd.DataFrame({"Início": ["2026-05-01"], "Término": ["2026-05-14"]})
        valido, missing = validate_planejamento_schema(df_valido)
        self.assertTrue(valido)
        self.assertEqual(missing, [])

        df_invalido = pd.DataFrame({"Início": ["2026-05-01"]})
        valido, missing = validate_planejamento_schema(df_invalido)
        self.assertFalse(valido)
        self.assertIn("Término", missing)

    def test_is_valid_xlsx(self):
        """Testa validação do cabeçalho binário (magic bytes PK)."""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(b"PK\x03\x04conteudo_zip_simulado")
            valid_path = f.name
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            f.write(b"HTML\x00conteudo_invalido")
            invalid_path = f.name

        try:
            self.assertTrue(DataLoader._is_valid_xlsx(valid_path))
            self.assertFalse(DataLoader._is_valid_xlsx(invalid_path))
        finally:
            if os.path.exists(valid_path):
                os.remove(valid_path)
            if os.path.exists(invalid_path):
                os.remove(invalid_path)

    def test_status_mapping_coverage(self):
        """Garante que todos os status mapeados pertençam a tipos válidos de atividade."""
        for sigla, tipo_canonico in STATUS_TO_TIPO_MAP.items():
            self.assertIn(
                tipo_canonico,
                VALID_ACTIVITY_TYPES,
                f"Status '{sigla}' mapeia para '{tipo_canonico}', que não está em VALID_ACTIVITY_TYPES.",
            )


if __name__ == "__main__":
    unittest.main()
