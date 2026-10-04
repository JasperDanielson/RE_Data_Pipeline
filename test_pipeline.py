"""Focused checks; no network, source cache or personal files required."""
import ast
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
import numpy as np
import pandas as pd
import RE_Pipeline as p

class PipelineTests(unittest.TestCase):
    def test_syntax(self):
        ast.parse(Path(p.__file__).read_text(encoding='utf-8'))

    def test_source_flags_are_not_zero(self):
        for flag in ['++', '**', '--', '']:
            value, status = p.cmhc_number(flag)
            self.assertTrue(np.isnan(value))
            self.assertNotEqual(status, 'observed')
        self.assertEqual(p.cmhc_number(0), (0., 'observed'))
        with self.assertRaises(ValueError):
            p.cmhc_number('unrecognized source token')

    def test_duplicate_keys_rejected(self):
        with self.assertRaises(ValueError):
            p.unique(pd.DataFrame({'year': [2025, 2025]}), ['year'], 'fixture')

    def test_valuation_independent_arithmetic(self):
        values, required = p.valuation(.05)
        row = values[(values.noi_growth_pct == 5) & (values.cap_rate_change_bps == 50)].iloc[0]
        self.assertAlmostEqual(row.implied_value_change_pct, (105/.055/(100/.05)-1)*100)
        self.assertAlmostEqual(required.loc[required.cap_rate_change_bps == 50, 'required_noi_growth_pct'].iloc[0], 10.)
        with self.assertRaises(ValueError):
            p.valuation(.005)

    def test_strict_json_missing_values(self):
        with TemporaryDirectory() as td:
            path = Path(td)/'result.json'
            p.write_json(path, {'missing': np.nan, 'infinite': np.inf})
            self.assertEqual(json.loads(path.read_text()), {'missing': None, 'infinite': None})

    def test_corrupt_offline_cache_rejected(self):
        with TemporaryDirectory() as td:
            source = p.Sources(Path(td), p.arguments(['--offline']))
            path = Path(td)/'input.json'
            path.write_text('{}')
            p.write_json(path.with_suffix('.json.metadata.json'), {'url': 'https://example.invalid', 'sha256': 'wrong'})
            try:
                with patch.object(source.session, 'get', side_effect=AssertionError('Network forbidden')):
                    with self.assertRaises(RuntimeError):
                        source.fetch('https://example.invalid', 'input.json', 'json')
            finally:
                source.session.close()

if __name__ == '__main__':
    unittest.main()
