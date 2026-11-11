import datetime as dt
import json
from pathlib import Path
import tempfile
import unittest

import build_report as report


class ReportTests(unittest.TestCase):
    def test_business_days(self):
        self.assertEqual(report.networkdays_minus_one(dt.date(2026, 9, 18), dt.date(2026, 9, 21)), 1)
        self.assertEqual(report.networkdays_minus_one(dt.date(2026, 9, 21), dt.date(2026, 9, 21)), 0)

    def test_sample_checks_and_draft(self):
        source = Path('sample_data/Sample_Campari_September_2026_KPI.xlsx')
        with tempfile.TemporaryDirectory(dir='output') as folder:
            path, checks = report.build(source, Path(folder))
            self.assertTrue(checks['issues'])
            self.assertIsNone(path)
            path, _ = report.build(source, Path(folder), draft=True)
            html = path.read_text(encoding='utf-8')
            payload = html.split('const R = ', 1)[1]
            data, _ = json.JSONDecoder().raw_decode(payload)
            self.assertEqual(sum(p['end'] for p in data['products']), 85631)
            self.assertTrue(data['meta']['draft'])
            self.assertEqual(len(data['meta']['sourceSha256']), 64)
            self.assertEqual(sum(not o['validStages'] for o in data['orders']), 2)
            self.assertIn("connect-src 'none'", html)
            self.assertNotIn('/*__REPORT_DATA__*/', html)
            for api in ('fetch(', 'XMLHttpRequest', 'sendBeacon', 'localStorage', 'type="file"'):
                self.assertNotIn(api, html)


if __name__ == '__main__':
    unittest.main()
