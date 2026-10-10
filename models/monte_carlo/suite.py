"""Run prespecified paired sensitivity strata; no scenario probability weights."""
import argparse
import json
from pathlib import Path
import run as mc

def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--draws', type=int, default=1024)
    p.add_argument('--workers', type=int, default=4)
    p.add_argument('--output', type=Path, default=mc.HERE/'results')
    a=p.parse_args()
    cases=[(s,'primary','primitive_independent') for s in mc.SPEC['access_strata'] if s!='working']
    cases += [('working',s,'primitive_independent') for s in mc.SPEC['structural_strata'] if s!='primary']
    cases += [('working','primary','commercial_rank_linked')]
    for access,structure,dependence in cases:
        saved=a.output/f'seed_{mc.SPEC["seed"]}'/mc.scenario_name(access,structure,dependence)/'summary.json'
        n=max(a.draws,json.loads(saved.read_text())['draws']) if saved.exists() else a.draws
        mc.batch(n,mc.SPEC['seed'],access,structure,dependence,a.workers,a.output)

if __name__=='__main__':
    main()
