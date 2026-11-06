"""
Testes unitários para o motor de disponibilidade, detecção de conflitos e severidade (v0.3).
"""

import unittest
import pandas as pd

from src.config import Config
from src.data_processor import DataProcessor, classify_conflict_severity
from src.schemas import sanitize_dates


class TestConflictEngine(unittest.TestCase):
    """Testes exaustivos do motor vetorizado de conflitos, severidade e datas."""

    def setUp(self):
        self.equipe_df = pd.DataFrame([
            {"Matrícula": "1001", "Nome": "ALBERTO SANTOS", "Disciplina": "MEC", "Função": "MECÂNICO", "Projeto": "P-80"},
            {"Matrícula": "1002", "Nome": "BEATRIZ COSTA", "Disciplina": "ELET", "Função": "ELETRICISTA", "Projeto": "P-82"},
            {"Matrícula": "1003", "Nome": "CARLOS DIAS", "Disciplina": "INST", "Função": "INSTRUMENTISTA", "Projeto": "P-83"},
        ])

    def test_no_conflicts_distinct_people_same_dates(self):
        """Dois colaboradores diferentes com eventos na mesma data não devem gerar conflito."""
        events = pd.DataFrame([
            {"Matrícula": "1001", "Nome": "ALBERTO SANTOS", "Início": pd.Timestamp("2026-06-01"), "Término": pd.Timestamp("2026-06-15"), "Tipo": "Embarque", "Projeto": "P-80"},
            {"Matrícula": "1002", "Nome": "BEATRIZ COSTA", "Início": pd.Timestamp("2026-06-01"), "Término": pd.Timestamp("2026-06-15"), "Tipo": "Embarque", "Projeto": "P-82"},
        ])
        conflicts, keys = DataProcessor.detect_conflicts(events)
        self.assertTrue(conflicts.empty)
        self.assertEqual(len(keys), 0)

    def test_ferias_totalmente_dentro_estaleiro_critico(self):
        """Cenário 1: Férias totalmente dentro de um período de Estaleiro (Severidade: Crítico)."""
        events = pd.DataFrame([
            {
                "Matrícula": "1001",
                "Nome": "ALBERTO SANTOS",
                "Início": pd.Timestamp("2026-06-01"),
                "Término": pd.Timestamp("2026-06-30"),
                "Tipo": "Estaleiro",
                "Projeto": "P-80",
            },
            {
                "Matrícula": "1001",
                "Nome": "ALBERTO SANTOS",
                "Início": pd.Timestamp("2026-06-10"),
                "Término": pd.Timestamp("2026-06-20"),
                "Tipo": "Férias",
                "Projeto": "P-80",
            },
        ])
        conflicts, keys = DataProcessor.detect_conflicts(events)
        self.assertEqual(len(conflicts), 1)

        row = conflicts.iloc[0]
        self.assertEqual(row["Matrícula"], "1001")
        self.assertEqual(row["Severidade"], Config.SEVERITY_CRITICO)
        self.assertTrue(row["Tem_Conflito"])
        self.assertIn("Férias", row["Detalhe_Conflito"])
        self.assertIn("Estaleiro", row["Detalhe_Conflito"])
        self.assertEqual(row["Início Conflito"], pd.Timestamp("2026-06-10"))
        self.assertEqual(row["Término Conflito"], pd.Timestamp("2026-06-20"))
        self.assertEqual(row["Sobreposição (dias)"], 11)  # 10 a 20 inclusive = 11 dias

    def test_sobreposicao_parcial_um_dia(self):
        """Cenário 2: Sobreposição parcial de exatamente 1 dia de interseção (dia D termina e inicia outro)."""
        events = pd.DataFrame([
            {
                "Matrícula": "1001",
                "Nome": "ALBERTO SANTOS",
                "Início": pd.Timestamp("2026-06-01"),
                "Término": pd.Timestamp("2026-06-14"),
                "Tipo": "Embarque",
                "Projeto": "P-80",
            },
            {
                "Matrícula": "1001",
                "Nome": "ALBERTO SANTOS",
                "Início": pd.Timestamp("2026-06-14"),
                "Término": pd.Timestamp("2026-06-28"),
                "Tipo": "Treinamento",
                "Projeto": "P-80",
            },
        ])
        conflicts, keys = DataProcessor.detect_conflicts(events)
        self.assertEqual(len(conflicts), 1)

        row = conflicts.iloc[0]
        self.assertEqual(row["Sobreposição (dias)"], 1)
        self.assertEqual(row["Início Conflito"], pd.Timestamp("2026-06-14"))
        self.assertEqual(row["Término Conflito"], pd.Timestamp("2026-06-14"))
        self.assertEqual(row["Severidade"], Config.SEVERITY_OPERACIONAL)

    def test_atividades_adjacentes_sem_conflito(self):
        """Cenário 3: Atividades adjacentes sem conflito (Término no dia D e Início no dia D+1)."""
        events = pd.DataFrame([
            {
                "Matrícula": "1001",
                "Nome": "ALBERTO SANTOS",
                "Início": pd.Timestamp("2026-06-01"),
                "Término": pd.Timestamp("2026-06-14"),
                "Tipo": "Embarque",
                "Projeto": "P-80",
            },
            {
                "Matrícula": "1001",
                "Nome": "ALBERTO SANTOS",
                "Início": pd.Timestamp("2026-06-15"),
                "Término": pd.Timestamp("2026-06-28"),
                "Tipo": "Folga",
                "Projeto": "P-80",
            },
        ])
        conflicts, keys = DataProcessor.detect_conflicts(events)
        self.assertTrue(conflicts.empty)
        self.assertEqual(len(keys), 0)

    def test_datas_invalidas_termino_anterior_ao_inicio(self):
        """Cenário 4: Datas inválidas (Término < Início) devem ser capturadas na auditoria de inconsistências."""
        df_raw = pd.DataFrame([
            {"Nome": "Válido", "Início": "01/06/2026", "Término": "14/06/2026"},
            {"Nome": "Invertido", "Início": "20/06/2026", "Término": "10/06/2026"},  # Término < Início
            {"Nome": "Nulo", "Início": "invalido", "Término": "14/06/2026"},        # NaT
        ])
        valid_df, inconsistencias = sanitize_dates(df_raw, "Início", "Término", filename="teste.xlsx")

        self.assertEqual(len(valid_df), 1)
        self.assertEqual(valid_df.iloc[0]["Nome"], "Válido")

        self.assertEqual(len(inconsistencias), 2)
        motivos = inconsistencias["Motivo"].tolist()
        self.assertIn("Término anterior ao Início", motivos)
        self.assertIn("Data Inválida ou Ausente (NaT)", motivos)

    def test_sobreposicao_operacional_severidade(self):
        """Dois eventos de trabalho simultâneos (Estaleiro + Treinamento) devem ser classificados como Operacional."""
        events = pd.DataFrame([
            {"Matrícula": "1001", "Nome": "ALBERTO", "Início": pd.Timestamp("2026-06-01"), "Término": pd.Timestamp("2026-06-10"), "Tipo": "Estaleiro", "Projeto": "P-80"},
            {"Matrícula": "1001", "Nome": "ALBERTO", "Início": pd.Timestamp("2026-06-05"), "Término": pd.Timestamp("2026-06-15"), "Tipo": "Treinamento", "Projeto": "P-80"},
        ])
        conflicts, _ = DataProcessor.detect_conflicts(events)
        self.assertEqual(len(conflicts), 1)
        self.assertEqual(conflicts.iloc[0]["Severidade"], Config.SEVERITY_OPERACIONAL)

    def test_filter_future_conflicts(self):
        """Testa descarte de conflitos passados, mantendo apenas eventos futuros."""
        hoje = pd.Timestamp("2026-06-15")
        conflicts_df = pd.DataFrame([
            {"Nome": "PASSADO", "Término Conflito": pd.Timestamp("2026-05-10")},
            {"Nome": "HOJE", "Término Conflito": pd.Timestamp("2026-06-15")},
            {"Nome": "FUTURO", "Término Conflito": pd.Timestamp("2026-06-20")},
        ])
        futuros = DataProcessor.filter_future_conflicts(conflicts_df, reference_date=hoje)
        self.assertEqual(len(futuros), 1)
        self.assertEqual(futuros.iloc[0]["Nome"], "FUTURO")

    def test_availability_in_period(self):
        """Verifica identificação precisa de colaboradores disponíveis vs ocupados."""
        combined_df = pd.DataFrame([
            {"Matrícula": "1001", "Início": pd.Timestamp("2026-07-01"), "Término": pd.Timestamp("2026-07-14")},
        ])
        disp = DataProcessor.available_in_period(
            self.equipe_df,
            combined_df,
            start_date=pd.Timestamp("2026-07-05"),
            end_date=pd.Timestamp("2026-07-10"),
        )
        self.assertNotIn("1001", disp)
        self.assertIn("1002", disp)
        self.assertIn("1003", disp)

        ocup = DataProcessor.occupied_in_period(
            combined_df,
            start_date=pd.Timestamp("2026-07-05"),
            end_date=pd.Timestamp("2026-07-10"),
        )
        self.assertEqual(ocup, {"1001"})


if __name__ == "__main__":
    unittest.main()
