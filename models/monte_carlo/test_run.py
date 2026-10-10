"""Numerical and integration regression tests; python -m unittest discover -s models/monte_carlo."""
import copy
import json
from pathlib import Path
import tempfile
import unittest
import run as mc
from mc_prepare import triangular_quantile


class MonteCarloTests(unittest.TestCase):
    def test_discrete_quantiles(self):
        self.assertEqual(mc.quantile([30, 0, 10, 20], .5), 10)
        self.assertEqual(mc.weighted_quantile([(0, .1), (20, .9)], .1), 0)
        self.assertEqual(mc.weighted_quantile([(0, .1), (20, .9)], .10001), 20)
        self.assertEqual(mc.weighted_quantile([(0, 0), (20, 1)], 0), 20)

    def test_rank_ties(self):
        self.assertEqual(mc.ranks([3, 1, 1, 4]), [3, 1.5, 1.5, 4])
        self.assertEqual(mc.correlation([1, 2, 3], [3, 2, 1]), -1)
        self.assertIsNone(mc.correlation([1, 1], [1, 2]))

    def test_sampler_endpoints(self):
        for p in mc.SPEC['priors'].values():
            a, m, b = p['min'], p['mode'], p['max']
            self.assertAlmostEqual(triangular_quantile(0, a, m, b), a)
            self.assertAlmostEqual(triangular_quantile(1, a, m, b), b)
            self.assertAlmostEqual(triangular_quantile((m-a)/(b-a), a, m, b), m)
        self.assertEqual(triangular_quantile(.3, 4, 4, 4), 4)

    def test_seed_prefix(self):
        self.assertEqual(list(mc.draws(10, 42)), list(mc.draws(20, 42))[:10])
        self.assertNotEqual(list(mc.draws(10, 42)), list(mc.draws(10, 43)))

    def test_zero_volatility_matches_working(self):
        _, q = next(mc.draws(1))
        o = mc.overrides_at(q, mc.SPEC, zero_volatility=True)
        self.assertEqual(mc.engine.resolve(mc.CFG, overrides=o), mc.engine.resolve(mc.CFG))
        r = mc.engine.evaluate(mc.CFG, overrides=o, retain=True)
        self.assertAlmostEqual(r['expected_12month_value'], 28.311616928946084, places=9)
        mc.assert_result(r, mc.engine.resolve(mc.CFG, overrides=o))
        bad = copy.deepcopy(r)
        bad['expected_12month_value'] *= .9
        with self.assertRaisesRegex(ValueError, 'Double or missing'):
            mc.assert_result(bad, mc.engine.resolve(mc.CFG, overrides=o))

    def test_scenarios_preserve_fixed_policies(self):
        _, q = next(mc.draws(1))
        for access in mc.SPEC['access_strata']:
            for structure in mc.SPEC['structural_strata']:
                o = mc.overrides_at(q, mc.SPEC, access=access, structural=structure)
                for k, v in mc.SPEC['hard_fixed'].items():
                    self.assertEqual(o[k], v)
        linked = mc.overrides_at(q, mc.SPEC, dependence='commercial_rank_linked')
        for k in mc.SPEC['commercial_rank_linked']:
            p = mc.SPEC['priors'][k]
            self.assertEqual(linked[k], triangular_quantile(q['category_switch'], p['min'], p['mode'], p['max']))

    def test_exact_clinical_enumeration(self):
        rows = [mc.evaluate_task((i,q,'working','primary','primitive_independent')) for i,q in mc.draws(4)]
        for row in rows:
            self.assertNotIn('error', row)
        s = mc.summarize(rows)
        self.assertAlmostEqual(s['expected_target']['mean'], s['clinical_outcomes_target']['mean'], places=12)
        self.assertEqual(s['invalid_draws'], 0)

    def test_checkpoint_resume_and_worker_reproducibility(self):
        with tempfile.TemporaryDirectory() as t:
            a, b = Path(t)/'a', Path(t)/'b'
            mc.batch(2, 71, 'working', 'primary', 'primitive_independent', 1, a)
            mc.batch(4, 71, 'working', 'primary', 'primitive_independent', 2, a)
            mc.batch(4, 71, 'working', 'primary', 'primitive_independent', 1, b)
            rel = Path('seed_71/working__primary__primitive_independent/draws.jsonl')
            self.assertEqual((a/rel).read_bytes(), (b/rel).read_bytes())
            identity = a/rel.parent/'identity.json'
            x = json.loads(identity.read_text()); x['hashes']['injected'] = 'changed'
            identity.write_text(json.dumps(x))
            with self.assertRaisesRegex(ValueError, 'hashes'):
                mc.batch(4, 71, 'working', 'primary', 'primitive_independent', 1, a)


if __name__ == '__main__':
    unittest.main()
