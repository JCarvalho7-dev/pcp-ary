"""Testes de negócio, persistência simulada e fluxos reais de Streamlit.

Execute na raiz do projeto: python -m unittest discover -s tests -v
"""

import io
import math
from datetime import date
from pathlib import Path
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app
from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest


class BusinessTests(unittest.TestCase):
    def setUp(self):
        self.tables = app.seed_tables()
        self.today = date(2026, 11, 18)

    def test_exact_seed_shape(self):
        self.assertEqual([len(v) for v in self.tables.values()], [13, 18, 10, 4])
        self.assertEqual(list(self.tables), list(app.SCHEMAS))
        self.assertEqual(self.tables['Ordens_Producao']['Qtd_Ton'].sum(), 211.0)

    def test_dashboard_reference_values(self):
        orders, stock = self.tables['Ordens_Producao'], self.tables['Estoque_Atual']
        self.assertEqual(len(orders[orders.Status.ne('Expedida')]), 10)
        self.assertEqual(orders.loc[orders.Status.isin(app.QUEUE_STATUS), 'Qtd_Ton'].sum(), 159.0)
        self.assertEqual(orders.loc[orders.Status.eq('Aguardando_Laudo'), 'Qtd_Ton'].sum(), 32.0)
        self.assertEqual(stock.Saldo_Disponivel.lt(stock.Estoque_Minimo).sum(), 5)

    def test_mrp_exact_requirements_with_losses(self):
        result, detail = app.calculate_mrp(self.tables)
        actual = result.set_index('SKU')
        expected = {'MP-ALUM-H02': 75.5812, 'MP-SIL-03': 32.136,
                    'MP-BAUX-01': 40.341, 'MP-ADIT-04': 116.15,
                    'MP-ALV-05': 243.6, 'EMB-SV-25': 1827.0,
                    'EMB-BB-1000': 98.49, 'EMB-PAL-PBR': 103.0}
        for sku, quantity in expected.items():
            self.assertAlmostEqual(actual.at[sku, 'Necessidade_Bruta'], quantity, places=6)
        self.assertEqual(detail.Num_OP.nunique(), 6)
        self.assertAlmostEqual(actual.at['MP-ALUM-H02', 'Necessidade_Compra_Urgente'], 13.0812)
        self.assertAlmostEqual(actual.at['EMB-SV-25', 'Necessidade_Compra_Urgente'], 427)
        self.assertEqual(actual.at['EMB-SV-25', 'Compra_Sugerida'], 2427)

    def test_mrp_in_process_and_quarantine_exclusion(self):
        result, detail = app.calculate_mrp(self.tables, True)
        actual = result.set_index('SKU')
        self.assertEqual(detail.Num_OP.nunique(), 7)
        self.assertAlmostEqual(actual.at['MP-ALUM-H02', 'Necessidade_Bruta'], 91.5748)
        self.assertAlmostEqual(actual.at['MP-ALUM-H02', 'Necessidade_Compra_Urgente'], 29.0748)

    def test_empty_queue(self):
        self.tables['Ordens_Producao'].loc[:, 'Status'] = 'Liberada'
        result, detail = app.calculate_mrp(self.tables)
        self.assertTrue(detail.empty)
        self.assertEqual(result.Necessidade_Bruta.sum(), 0)
        original, optimized = app.sequence_scenarios(self.tables)
        self.assertEqual(app.count_setups(original, self.tables['Capacidade_Linhas']), (0, 0))
        self.assertTrue(optimized.empty)

    def test_setup_comparison(self):
        original, optimized = app.sequence_scenarios(self.tables)
        self.assertEqual(app.count_setups(original, self.tables['Capacidade_Linhas']), (5, 23.5))
        self.assertEqual(app.count_setups(optimized, self.tables['Capacidade_Linhas']), (3, 14.5))
        self.assertEqual(optimized.Num_OP.tolist(), ['OP-2026-1044', 'OP-2026-1046', 'OP-2026-1045', 'OP-2026-1048', 'OP-2026-1047', 'OP-2026-1049'])

    def test_atp_immediate(self):
        result = app.calculate_atp(self.tables, 'PA-ALC-325', 20, self.today)
        self.assertTrue(result['immediate'])
        self.assertEqual(result['release'], self.today)

    def test_atp_partial_stock(self):
        result = app.calculate_atp(self.tables, 'PA-ALC-325', 50, self.today)
        self.assertEqual(result['net'], 12)
        self.assertEqual(result['queue_tons'], 58)
        self.assertEqual(result['hours'], 24)
        self.assertEqual(result['days'], 2)
        self.assertEqual(result['release'], date(2026, 11, 20))
        self.assertFalse(result['material_risks'].empty)

    def test_atp_validation(self):
        for quantity in [0, -1, math.nan, math.inf]:
            with self.assertRaises(app.DataError):
                app.calculate_atp(self.tables, 'PA-ALC-325', quantity, self.today)

    def test_business_weekend(self):
        self.assertEqual(app.add_business_days(date(2026, 11, 20), 2), date(2026, 11, 24))
        self.assertEqual(app.add_business_days(date(2026, 11, 21), 1), date(2026, 11, 24))

    def test_approval_transfer_and_idempotency(self):
        updated, info = app.approval_result(self.tables, 'OP-2026-1041')
        stock = updated['Estoque_Atual'].set_index('SKU')
        self.assertEqual(stock.at['PA-ALT-014', 'Saldo_Disponivel'], 32.5)
        self.assertEqual(stock.at['PA-ALT-014', 'Saldo_Quarentena'], 0)
        self.assertEqual(info['shortfall'], 0)
        self.assertEqual(self.tables['Estoque_Atual'].set_index('SKU').at['PA-ALT-014', 'Saldo_Disponivel'], 14.5)
        with self.assertRaises(app.DataError):
            app.approval_result(updated, 'OP-2026-1041')

    def test_approval_clamps_and_reports_divergence(self):
        stock = self.tables['Estoque_Atual']
        stock.loc[stock.SKU.eq('PA-ALT-014'), 'Saldo_Quarentena'] = 10
        updated, info = app.approval_result(self.tables, 'OP-2026-1041')
        self.assertEqual(updated['Estoque_Atual'].set_index('SKU').at['PA-ALT-014', 'Saldo_Quarentena'], 0)
        self.assertEqual(info['shortfall'], 8)

    def test_replenishment_does_not_duplicate_queued_production(self):
        self.assertEqual(app.replenishment_plan(self.tables).Producao_Sugerida_Ton.sum(), 0)
        self.tables['Ordens_Producao'].loc[:, 'Status'] = 'Liberada'
        plan = app.replenishment_plan(self.tables).set_index('SKU')
        self.assertEqual(plan.at['PA-ALT-014', 'Producao_Sugerida_Ton'], 10.5)
        self.assertEqual(plan.at['PA-FRS-75', 'Producao_Sugerida_Ton'], 6)

    def test_exceptions_reconcile_initial_data(self):
        alerts = app.exceptions(self.tables, self.today)
        self.assertEqual(len(alerts[alerts.Tipo.eq('MATERIAL')]), 4)
        self.assertEqual(len(alerts[alerts.Tipo.eq('CONCILIAÇÃO')]), 2)
        self.assertIn('OP-2026-1041', alerts.Referencia.tolist())

    def test_schema_rejects_duplicates_and_invalid_values(self):
        self.tables['Estoque_Atual'].iloc[0, 4] = -10
        with self.assertRaises(app.DataError):
            app.validate_tables(self.tables)
        tables = app.seed_tables()
        tables['Ordens_Producao'].iloc[1, 0] = tables['Ordens_Producao'].iloc[0, 0]
        with self.assertRaises(app.DataError):
            app.validate_tables(tables)

    def test_mrp_rejects_missing_bom(self):
        self.tables['Fichas_Tecnicas_BOM'] = self.tables['Fichas_Tecnicas_BOM'].loc[lambda x: x.SKU_Pai.ne('PA-FRS-75')]
        with self.assertRaises(app.DataError):
            app.calculate_mrp(self.tables)

    def test_exports_roundtrip_and_formula_protection(self):
        data = app.workbook_bytes(self.tables)
        workbook = load_workbook(io.BytesIO(data), data_only=False)
        self.assertEqual(workbook.sheetnames, list(app.SCHEMAS))
        self.assertEqual(workbook['Ordens_Producao'].max_row, 11)
        self.assertEqual(workbook['Ordens_Producao']['A2'].value, 'OP-2026-1041')
        self.assertTrue(app.csv_bytes(self.tables['Estoque_Atual']).startswith(b'\xef\xbb\xbf'))
        malicious = self.tables['Ordens_Producao'].copy()
        malicious.iloc[0, 1] = '=HYPERLINK("https://example.test")'
        exported = app.workbook_bytes({'Orders': malicious})
        self.assertEqual(load_workbook(io.BytesIO(exported))['Orders']['B2'].data_type, 's')


class PersistenceTests(unittest.TestCase):
    def setUp(self):
        self.tables = app.seed_tables()

    def test_sandbox_create_approve_and_block(self):
        state = {'sandbox_tables': self.tables}
        repo = app.Repository('sandbox')
        with patch.object(app.st, 'session_state', state):
            repo.approve('OP-2026-1041')
            with self.assertRaises(app.DataError):
                repo.approve('OP-2026-1041')
            repo.change_status('OP-2026-1042', 'Bloqueada')
            stock = repo.load()['Estoque_Atual'].set_index('SKU')
            self.assertEqual(stock.at['PA-FRS-75', 'Saldo_Quarentena'], 14)
            payload = {'Num_OP': 'OP-2026-1051', 'Cliente': 'Teste', 'SKU_Produto': 'PA-ALC-325', 'Qtd_Ton': 2.0,
                       'Familia_Setup': 'Alumina_Padrao', 'Data_Entrega': '2026-11-30', 'Status': 'Programada', 'Lote': 'LT-TESTE-001'}
            repo.add_order(payload)
            self.assertEqual(len(repo.load()['Ordens_Producao']), 11)
            with self.assertRaises(app.DataError):
                repo.add_order(payload)

    def test_cloud_approval_is_one_atomic_batch_and_clears_cache(self):
        repo = app.Repository('cloud', 'https://example.test', '{}')
        sheet = MagicMock()
        sheet.worksheets.return_value = [MagicMock(title=name, id=i + 10) for i, name in enumerate(app.SCHEMAS)]
        with patch.object(app.Repository, 'load', return_value=self.tables), patch.object(app, 'cloud_connection', return_value=sheet), patch.object(app.st, 'session_state', {}), patch.object(app.st.cache_data, 'clear') as clear:
            repo.approve('OP-2026-1041')
            requests = sheet.batch_update.call_args.args[0]['requests']
            self.assertEqual(len(requests), 3)
            self.assertEqual(sheet.batch_update.call_count, 1)
            self.assertEqual(clear.call_count, 1)
            self.assertEqual(requests[0]['updateCells']['rows'][0]['values'][0]['userEnteredValue']['stringValue'], 'Liberada')

    def test_failed_cloud_write_invalidates_stale_cache(self):
        repo = app.Repository('cloud', 'https://example.test', '{}')
        sheet = MagicMock()
        sheet.worksheets.return_value = [MagicMock(title=name, id=i) for i, name in enumerate(app.SCHEMAS)]
        sheet.batch_update.side_effect = RuntimeError('Connection lost')
        with patch.object(app, 'cloud_connection', return_value=sheet), patch.object(app.st.cache_data, 'clear') as clear:
            with self.assertRaises(RuntimeError):
                repo.commit(self.tables, [('Ordens_Producao', 2, 'Status', 'Liberada')])
            clear.assert_called_once()

    def test_cloud_read_keeps_physical_rows_when_blank_lines_exist(self):
        values = []
        for name, frame in self.tables.items():
            rows = [app.SCHEMAS[name], [], *frame.values.tolist()]
            values.append({'values': rows})
        sheet = MagicMock()
        sheet.values_batch_get.return_value = {'valueRanges': values}
        read = app.cloud_tables(sheet)
        self.assertEqual(read['Ordens_Producao'].index[0], 3)
        self.assertEqual(read['Estoque_Atual'].index[0], 3)


class InterfaceTests(unittest.TestCase):
    def app_test(self):
        at = AppTest.from_file(str(Path(app.__file__)), default_timeout=60)
        at.secrets['authentication_required'] = False
        return at.run()

    def test_all_five_pages_render_without_exceptions(self):
        at = self.app_test()
        self.assertFalse(at.exception, str(at.exception))
        for page in app.NAVIGATION[1:]:
            at.radio(key='navigation').set_value(page).run()
            self.assertFalse(at.exception, f'{page}: {at.exception}')

    def test_atp_submit(self):
        at = self.app_test()
        at.radio(key='navigation').set_value(app.NAVIGATION[3]).run()
        at.button(key='FormSubmitter:atp_form-Calcular disponibilidade e prazo').click().run()
        self.assertFalse(at.exception, str(at.exception))
        self.assertIn('atp_request', at.session_state)

    def test_new_order_form_persists(self):
        at = self.app_test()
        next(widget for widget in at.text_input if widget.label == 'Cliente').set_value('Cliente de homologacao')
        at.button(key='FormSubmitter:new_order-Cadastrar OP').click().run()
        self.assertFalse(at.exception, str(at.exception))
        orders = at.session_state['sandbox_tables']['Ordens_Producao']
        self.assertEqual(len(orders), 11)
        self.assertEqual(orders.iloc[-1]['Cliente'], 'Cliente de homologacao')

    def test_mrp_toggle_changes_scope(self):
        at = self.app_test()
        at.radio(key='navigation').set_value(app.NAVIGATION[1]).run()
        at.toggle(key='mrp_in_process').set_value(True).run()
        self.assertFalse(at.exception, str(at.exception))
        self.assertTrue(at.toggle(key='mrp_in_process').value)

    def test_quality_approve_removes_action_and_transfers_stock(self):
        at = self.app_test()
        at.radio(key='navigation').set_value(app.NAVIGATION[4]).run()
        at.button(key='approve_OP-2026-1041').click().run()
        self.assertFalse(at.exception, str(at.exception))
        stock = at.session_state['sandbox_tables']['Estoque_Atual'].set_index('SKU')
        self.assertEqual(stock.at['PA-ALT-014', 'Saldo_Disponivel'], 32.5)
        self.assertNotIn('approve_OP-2026-1041', [b.key for b in at.button])

    def test_login_rejects_wrong_password_and_accepts_correct_password(self):
        at = AppTest.from_file(str(Path(app.__file__)), default_timeout=60)
        at.secrets['app_username'] = 'Aryane'
        at.secrets['app_password'] = 'test-secret-123'
        at.secrets['auth_cookie_key'] = 'test-only-cookie-signing-key-minimum-32-characters'
        at.run()
        next(widget for widget in at.text_input if widget.label == 'Login').set_value('Aryane')
        next(widget for widget in at.text_input if widget.label == 'Senha').set_value('wrong')
        at.button[0].click().run()
        self.assertTrue(at.error)
        next(widget for widget in at.text_input if widget.label == 'Senha').set_value('test-secret-123')
        at.button[0].click().run()
        self.assertFalse(at.exception, str(at.exception))
        self.assertIn('navigation', at.session_state)
        at.radio(key='navigation').set_value(app.NAVIGATION[1]).run()
        self.assertTrue(at.session_state['authentication_status'])

    def test_production_refuses_missing_persistent_storage(self):
        with patch.object(app, 'secret', side_effect=lambda key, default=None: True if key == 'production_mode' else default):
            with self.assertRaises(app.DataError):
                app.Repository.configured()

    def test_default_personal_password_is_hashed(self):
        import bcrypt
        self.assertTrue(bcrypt.checkpw(b'Ary123', app.DEFAULT_PASSWORD_HASH.encode()))


if __name__ == '__main__':
    unittest.main()
