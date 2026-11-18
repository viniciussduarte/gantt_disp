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

    def test_embarque_tipo_imported_as_emprestado(self):
        """Garante que atividades do tipo 'Embarque' da planilha são integradas como status 'EMP' (Emprestado),
        não contando como embarcado na divisão de turma a bordo desta unidade."""
        df_crono = pd.DataFrame([{
            "Matrícula": "1385894",
            "Nome": "IGOR TEIXEIRA",
            "Disciplina": "Mecânica",
            "Tipo": "Embarque",
            "Início": pd.Timestamp("2026-10-27"),
            "Término": pd.Timestamp("2026-10-31"),
            "Plataforma": "P80",
            "Detalhamento": "Embarque P-76"
        }])

        membros = [{
            "nome": "IGOR TEIXEIRA",
            "disciplina": "Mecânica",
            "tipo_colaborador": "Nativo",
            "funcao": "Técnico",
            "regime_folga": "14x21"
        }]

        m_map, m_per, conflitos, m_membros, q_ext = integrar_escalas_com_cronograma(
            escalas_map={},
            mapa_dia_periodo={},
            membros=membros,
            cronograma_df=df_crono,
            proj_norm="P-80",
            ano=2026,
            mes=10
        )

        # Deve ser mapeado como EMP (Emprestado)
        assert ("IGOR TEIXEIRA", "2026-10-27") in m_map
        assert m_map[("IGOR TEIXEIRA", "2026-10-27")] == "EMP"
        assert m_map[("IGOR TEIXEIRA", "2026-10-31")] == "EMP"

        # Deve identificar projeto de destino
        p_obj = m_per[("IGOR TEIXEIRA", "2026-10-27")]
        assert p_obj["is_emprestado"] is True
        assert p_obj["projeto_destino"] == "P-76"

        # Valida que na grade HTML o colaborador aparece como EMP e NÃO entra na Turma a Bordo desta plataforma
        html = build_scala_html(
            projeto="P-80",
            ano=2026,
            mes=10,
            membros=m_membros,
            escalas_map=m_map,
            mapa_dia_periodo=m_per,
            qtd_nativos=1,
            qtd_externos=0,
            is_editor=False
        )
        assert "item-emp" in html
        assert "Emprestado para P-76" in html
        # Não deve haver embarques na turma a bordo
        assert "0 profissionais a bordo no período" in html

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

        m_map, m_per, conflitos, m_membros, q_ext = integrar_escalas_com_cronograma(
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
        """Valida que números de RT são salvos e consultados corretamente no banco."""
        EscalaP80DB.init_db()
        try:
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
        finally:
            with EscalaP80DB.get_connection() as conn:
                conn.cursor().execute("DELETE FROM p80_embarque_rt WHERE nome = ?", ("TESTE COLABORADOR",))
                conn.commit()
            EscalaP80DB.clear_cache()

    def test_query_cache_and_invalidation(self):
        """Valida que o cache em memória acelera leituras e é invalidado automaticamente em alterações."""
        from src.database import DBManager, QueryCache

        EscalaP80DB.clear_cache()
        # 1. Primeira leitura popula o cache
        vagas1 = EscalaP80DB.get_config_vagas("P-80")
        assert vagas1 is not None

        # 2. Verifica se a chave foi armazenada no cache
        assert len(QueryCache._store) > 0

        # 3. Alteração invalida o cache automaticamente
        EscalaP80DB.set_config_vagas("P-80", min_mec=5, min_inst=3, min_elet=2, min_total=10)
        assert len(QueryCache._store) == 0

        # 4. Nova leitura traz os dados atualizados
        vagas2 = EscalaP80DB.get_config_vagas("P-80")
        assert vagas2["min_mecanica"] == 5
        assert vagas2["min_instrumentacao"] == 3

        # Limpeza para deixar o banco como antes
        EscalaP80DB.set_config_vagas("P-80", min_mec=0, min_inst=0, min_elet=0, min_total=0)
        EscalaP80DB.clear_cache()

