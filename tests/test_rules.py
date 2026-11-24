import unittest
from datetime import date, timedelta

import pandas as pd

from src.rules import (add_responsible, calculate_pending, classify_indicator_records,
                       is_active, is_inactive, is_card_or_partnership, norm_key)


class RuleTests(unittest.TestCase):
    def test_pending_date_boundary_and_missing(self):
        today = date(2026, 1, 8)
        self.assertEqual(calculate_pending("Sim", today - timedelta(days=7), today, 7, 10), "Dentro do prazo")
        self.assertEqual(calculate_pending("Sim", today - timedelta(days=8), today, 7, 10), "META PENDENTE — LARANJA")
        self.assertEqual(calculate_pending("Sim", today - timedelta(days=10), today, 7, 10), "META PENDENTE — LARANJA")
        self.assertEqual(calculate_pending("Sim", today - timedelta(days=11), today, 7, 10), "META PENDENTE — VERMELHA")
        self.assertEqual(calculate_pending("Sim", None, today, 7, 10), "META PENDENTE — DATA NÃO INFORMADA")
        self.assertTrue(calculate_pending("Revisar", today - timedelta(days=30), today, 7, 10).startswith("Não classificado"))

    def test_normalization_and_inactive_statuses(self):
        self.assertEqual(norm_key("Clínica  São Camilo!"), "CLINICASAOCAMILO")
        self.assertFalse(is_active("NÃO SE APLICA"))
        self.assertFalse(is_active("Cliente sem interesse"))
        self.assertFalse(is_active("Aguardando editais"))
        self.assertTrue(is_inactive("Cliente sem interesse"))
        self.assertFalse(is_active("DESCREDENCIADO"))
        self.assertTrue(is_active("EM NEGOCIAÇÃO"))
        self.assertTrue(is_card_or_partnership("Cartão de desconto"))
        self.assertTrue(is_card_or_partnership("PARCERIA"))
        self.assertFalse(is_card_or_partnership("OPERADORA"))

    def test_official_priority_overrides(self):
        records = pd.DataFrame([
            {"client": "FISIO LIFE", "priority": "CREDENCIAMENTO"},
            {"client": "FISIO LIFE", "priority": "META CREDENCIAMENTO"},
            {"client": "FISIO LIFE", "priority": "REAJUSTE"},
        ])
        roster = pd.DataFrame([
            {"client": "FISIO LIFE", "responsible": "Dulcineia", "priority": "Todas as PRIORIDADES, com exceção CREDENCIAMENTO e META CREDENCIAMENTO"},
            {"client": "TODOS OS CLIENTES", "responsible": "Felipe", "priority": "CREDENCIAMENTO"},
            {"client": "TODOS OS CLIENTES", "responsible": "Niziane", "priority": "META CREDENCIAMENTO"},
        ])
        result = add_responsible(records, roster, pd.DataFrame(columns=["apelido_cliente", "cliente_cadastro"]))
        self.assertEqual(result.responsible.tolist(), ["Felipe", "Niziane", "Dulcineia"])

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
