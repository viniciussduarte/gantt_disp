"""
Testes unitários para as novas funcionalidades (v0.3):
- Exclusão de fins de semana (sábado e domingo) para atividades em BASE
- Mapeamento e inclusão de 'Embarque' no cronograma (correção Igor Teixeira)
- Grupo FSE e contabilização na Turma a Bordo pela disciplina cadastrada
- Persistência e recuperação de números de RT (Requisição de Transporte)
"""

from datetime import date
import pandas as pd
import pytest

from escala_p80 import (
    EscalaP80DB,
    build_scala_html,
    integrar_escalas_com_cronograma,
)


class TestEscalaFeaturesV03:
    """Validações de regras de negócio introduzidas na v0.3."""

    def test_embarque_tipo_included_in_cronograma(self):
        """Garante que atividades do tipo 'Embarque' são integradas como status 'EMB'."""
        df_crono = pd.DataFrame([{
            "Matrícula": "1385894",
            "Nome": "IGOR TEIXEIRA",
            "Disciplina": "Mecânica",
            "Tipo": "Embarque",
            "Início": pd.Timestamp("2026-10-27"),
            "Término": pd.Timestamp("2026-10-31"),
            "Plataforma": "P80",
            "Detalhamento": "Embarque de Teste"
        }])

        membros = [{
            "nome": "IGOR TEIXEIRA",
            "disciplina": "Mecânica",
            "tipo_colaborador": "Nativo",
            "funcao": "Técnico",
            "regime_folga": "14x21"
        }]

        m_map, m_per, m_membros, conflitos, q_ext = integrar_escalas_com_cronograma(
            escalas_map={},
            mapa_dia_periodo={},
            membros=membros,
            cronograma_df=df_crono,
            proj_norm="P-80",
            ano=2026,
            mes=10
        )

        assert ("IGOR TEIXEIRA", "2026-10-27") in m_map
        assert m_map[("IGOR TEIXEIRA", "2026-10-27")] == "EMB"
        assert m_map[("IGOR TEIXEIRA", "2026-10-31")] == "EMB"

    def test_base_weekend_exclusion(self):
        """Garante que lançamentos de BASE pulam sábados (5) e domingos (6)."""
        # 2026-10-09 é Sexta, 2026-10-10 é Sábado, 2026-10-11 é Domingo, 2026-10-12 é Segunda
        df_crono = pd.DataFrame([{
            "Matrícula": "9999",
            "Nome": "COLABORADOR BASE",
            "Disciplina": "Mecânica",
            "Tipo": "Estaleiro",  # Mapeia para BASE
            "Início": pd.Timestamp("2026-10-09"),
            "Término": pd.Timestamp("2026-10-12"),
            "Plataforma": "P80",
            "Detalhamento": "Atividade em Estaleiro"
        }])

        membros = [{
            "nome": "COLABORADOR BASE",
            "disciplina": "Mecânica",
            "tipo_colaborador": "Nativo",
            "funcao": "Técnico",
            "regime_folga": "14x21"
        }]

        m_map, m_per, m_membros, conflitos, q_ext = integrar_escalas_com_cronograma(
            escalas_map={},
            mapa_dia_periodo={},
            membros=membros,
            cronograma_df=df_crono,
            proj_norm="P-80",
            ano=2026,
            mes=10
        )

        # Sexta-feira: deve constar BASE
        assert m_map.get(("COLABORADOR BASE", "2026-10-09")) == "BASE"
        # Sábado e Domingo: NÃO devem constar BASE
        assert ("COLABORADOR BASE", "2026-10-10") not in m_map
        assert ("COLABORADOR BASE", "2026-10-11") not in m_map
        # Segunda-feira: deve constar BASE
        assert m_map.get(("COLABORADOR BASE", "2026-10-12")) == "BASE"

    def test_fse_grouping_and_turma_counting(self):
        """Garante que FSE fica em grupo visual separado mas pontua para sua disciplina cadastrada."""
        membros = [
            {
                "nome": "FSE MECANICA",
                "disciplina": "Mecânica",
                "tipo_colaborador": "Nativo",
                "funcao": "FSE",
                "regime_folga": "14x21"
            }
        ]
        escalas_map = {("FSE MECANICA", "2026-10-01"): "EMB"}
        mapa_dia_periodo = {
            ("FSE MECANICA", "2026-10-01"): {
                "nome": "FSE MECANICA",
                "disciplina": "Mecânica",
                "projeto": "P-80",
                "inicio": date(2026, 10, 1),
                "fim": date(2026, 10, 14),
                "status": "EMB",
                "dias": 14,
                "observacao": "Embarque FSE"
            }
        }

        html = build_scala_html(
            projeto="P-80",
            ano=2026,
            mes=10,
            membros=membros,
            escalas_map=escalas_map,
            mapa_dia_periodo=mapa_dia_periodo,
            qtd_nativos=1,
            qtd_externos=0,
            is_editor=False
        )

        # Cabeçalho do grupo FSE deve estar presente
        assert "📁 FSE" in html
        # Subtexto indicando FSE · Mecânica deve estar presente
        assert "FSE · Mecânica" in html

    def test_rt_persistence(self):
        """Valida que números de RT são salvos e consultados corretamente no banco SQLite."""
        EscalaP80DB.init_db()
        ok = EscalaP80DB.salvar_embarque_rt(
            projeto="P-80",
            nome="TESTE COLABORADOR",
            data_inicio="2026-10-27",
            data_fim="2026-11-09",
            numero_rt="RT-TESTE-2026"
        )
        assert ok is True

        rts = EscalaP80DB.get_embarque_rts()
        chave = ("P-80", "TESTE COLABORADOR", "2026-10-27", "2026-11-09")
        assert chave in rts
        assert rts[chave] == "RT-TESTE-2026"
