"""Net population scaling only; not PHVS uptake or separate demographic flows."""
import json
from pathlib import Path
BASE=Path(__file__).resolve().parent
def project_population(start, annual_rate, years):
    if start < 0 or annual_rate <= -1 or years < 0:
        raise ValueError("Invalid population, rate, or horizon")
    return start*(1+annual_rate)**years
if __name__ == '__main__':
    cfg=json.loads((BASE/'working_assumptions.json').read_text())['population_growth']
    result={str(g):{str(t):project_population(100,g,t) for t in [0,5,10]} for g in cfg['scenario_annual_rates']}
    assert project_population(100,0,10)==100
    assert abs(project_population(100,.015,5)*(1.015)**5-project_population(100,.015,10))<1e-9
    (BASE/'population_growth_results.json').write_text(json.dumps({'unit':'index;initial population100','status':'net population scaling, not PHVS forecast','scenarios':result},indent=2)+'\n')
    print(result)
