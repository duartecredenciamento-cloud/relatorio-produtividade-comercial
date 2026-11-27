import unittest
from datetime import date, timedelta

import pandas as pd

from src.rules import (add_responsible, calculate_pending, classify_indicator_records,
                       is_active, is_inactive, is_card_or_partnership, make_quality, norm_key)


class RuleTests(unittest.TestCase):
    def test_pending_date_boundary_and_missing(self):
        today = date(2026, 1, 8)
        self.assertEqual(calculate_pending("Sim", today - timedelta(days=7), today, 7, 10), "Dentro do prazo")
        self.assertEqual(calculate_pending("Revisar", today - timedelta(days=8), today, 7, 10), "META PENDENTE — LARANJA")
        self.assertEqual(calculate_pending("Não", today - timedelta(days=10), today, 7, 10), "META PENDENTE — LARANJA")
        self.assertEqual(calculate_pending("Sim", today - timedelta(days=11), today, 7, 10), "META PENDENTE — VERMELHA")
        self.assertEqual(calculate_pending("Sim", None, today, 7, 10), "META PENDENTE — DATA NÃO INFORMADA")

    def test_normalization_and_inactive_statuses(self):
        self.assertEqual(norm_key("Clínica  São Camilo!"), "CLINICASAOCAMILO")
        self.assertFalse(is_active("NÃO SE APLICA"))
        self.assertFalse(is_active("Cliente sem interesse"))
        self.assertFalse(is_active("Aguardando editais"))
        self.assertTrue(is_inactive("Cliente sem interesse"))
        self.assertTrue(is_active("DESCREDENCIADO"))
        self.assertFalse(is_inactive("DESCREDENCIADO"))
        self.assertTrue(is_active("EM NEGOCIAÇÃO"))
        self.assertTrue(is_card_or_partnership("Cartão de desconto"))
        self.assertTrue(is_card_or_partnership("PARCERIA"))
        self.assertFalse(is_card_or_partnership("OPERADORA"))

    def test_official_priority_overrides(self):
        records = pd.DataFrame([
            {"client": "Cliente Exemplo", "priority": "CREDENCIAMENTO"},
            {"client": "Cliente Exemplo", "priority": "META CREDENCIAMENTO"},
            {"client": "Cliente Exemplo", "priority": "REAJUSTE"},
        ])
        roster = pd.DataFrame([
            {"client": "Cliente Exemplo", "responsible": "Responsável Alfa", "priority": "Todas as PRIORIDADES, com exceção CREDENCIAMENTO e META CREDENCIAMENTO"},
            {"client": "TODOS OS CLIENTES", "responsible": "Responsável Beta", "priority": "CREDENCIAMENTO"},
            {"client": "TODOS OS CLIENTES", "responsible": "Responsável Gama", "priority": "META CREDENCIAMENTO"},
        ])
        result = add_responsible(records, roster, pd.DataFrame(columns=["apelido_cliente", "cliente_cadastro"]))
        self.assertEqual(result.responsible.tolist(), ["Responsável Beta", "Responsável Gama", "Responsável Alfa"])

    def test_matrix_quality_errors_keep_available_operator_even_without_records(self):
        files = [{"name": "matriz-teste.xlsm", "operator": "Operadora Teste", "error": "Falha de leitura", "warnings": ["Estrutura incompleta"]}]
        quality = make_quality(pd.DataFrame(), files, date(2026, 1, 8), 7)
        self.assertEqual(len(quality), 2)
        self.assertEqual(quality.operadora.tolist(), ["Operadora Teste", "Operadora Teste"])
        self.assertEqual(quality.tipo.tolist(), ["Falha de leitura; dados anteriores preservados", "Estrutura incompleta"])

    def test_indicators_separate_card_and_operator_categories(self):
        rows = pd.DataFrame([
            {"service": "CREDENCIAMENTO", "status": "EM NEGOCIAÇÃO", "category": "CARTÃO DE DESCONTO"},
            {"service": "CREDENCIAMENTO", "status": "AGUARDANDO CONTRATO", "category": "OPERADORA"},
            {"service": "CREDENCIAMENTO", "status": "IMPLANTAÇÃO", "category": "PARCERIA"},
        ])
        indicators = classify_indicator_records(rows)
        self.assertEqual(len(indicators["EM NEG. CARTÃO E PARCERIA (EM NEGOCIAÇÃO)"]), 1)
        self.assertEqual(len(indicators["PROCESSOS FINALIZANDO OPERADORA (AG. CONTRATO/IMPLANTAÇÃO)"]), 1)
        self.assertEqual(len(indicators["FINALIZANDO AGUARDANDO CONTRATO"]), 1)


if __name__ == "__main__":
    unittest.main()
