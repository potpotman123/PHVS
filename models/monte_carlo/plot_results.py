"""Optional publication charts from compact artifacts; requires matplotlib."""
import gzip
import hashlib
import json
import math
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter

HERE=Path(__file__).resolve().parent

def main():
    folder=HERE/'results'
    a=json.loads((folder/'analysis.json').read_text());s=a['primary']
    d=json.loads(gzip.decompress((folder/f'seed_{s["seed"]}'/'working__primary__primitive_independent/draws_compact.json.gz').read_bytes()))
    ix={c:i for i,c in enumerate(d['columns'])};rows=d['rows'];n=len(rows)
    expected=[r[ix['expected_target']] for r in rows]
    outcome=[];weights=[]
    for r in rows:
        for state,w in s['clinical_outcomes_target']['clinical_weights'].items():
            outcome.append(r[ix[f'{state}.target.price']]);weights.append(w/n)
    plt.rcParams.update({'font.family':'DejaVu Sans','font.size':10,'axes.spines.top':False,'axes.spines.right':False})
    fig,axes=plt.subplots(1,2,figsize=(13.2,5.8),gridspec_kw={'width_ratios':[1.15,1]})
    fig.patch.set_facecolor('#fafbfc')
    ax=axes[0];bins=list(range(0,2*math.ceil(max(expected+outcome)/2)+3,2))
    ax.hist(outcome,bins=bins,weights=weights,color='#dc9656',alpha=.65,label='Clinical outcomes, weighted once')
    ax.hist(expected,bins=bins,weights=[1/n]*n,histtype='step',linewidth=2.1,color='#285f8f',label='Risk-adjusted expected values')
    ax.set_xlabel('Modeled target value, USD per share');ax.set_ylabel('Probability mass per $2 bin')
    ax.yaxis.set_major_formatter(PercentFormatter(1));ax.set_xlim(0,bins[-1])
    ax.set_title('Two different uncertainty distributions',loc='left',fontweight='bold',pad=12)
    ax.legend(frameon=False,fontsize=8.8,loc='upper right');ax.grid(axis='y',alpha=.15)
    top=s['rank_associations_prior_dependent_not_causal'][:8][::-1]
    labels={"ir_doses_ltp":"Rescue doses on prevention",'switch_mature_multiplier':'Mature switching opportunity','population':'Diagnosed population','category_switch':'Category switching rate','ir_retention':'Rescue continuation','cost_scale':'Operating cost scale','oral_launch_share':'PHVS oral choice at launch','growth':'Net population growth','xr_mature_retention':'Mature preventive retention'}
    ax=axes[1];x=[r['spearman'] for r in top]
    ax.barh([labels.get(r['parameter'],r['parameter']) for r in top],x,color=['#b76449' if v<0 else '#285f8f' for v in x],height=.65)
    ax.axvline(0,color='#888',linewidth=.7);ax.set_xlim(-.4,.55)
    ax.set_xlabel('Spearman rank correlation with expected value')
    ax.set_title('Leading drivers under these priors',loc='left',fontweight='bold',pad=12)
    ax.grid(axis='x',alpha=.15)
    for i,v in enumerate(x): ax.text(v+(.012 if v>=0 else -.012),i,f'{v:+.2f}',va='center',ha='left' if v>=0 else 'right',fontsize=9)
    fig.suptitle('PHVS | Standalone valuation uncertainty',x=.04,ha='left',fontsize=19,fontweight='bold',y=.97)
    fig.text(.04,.89,f'{n:,} commercial draws • target 9 Oct 2027 • source e7594c4 • seed {s["seed"]}',fontsize=10,color='#555')
    fig.text(.04,.035,'Analyst-prior sensitivity, not an empirical confidence interval or stock-price forecast.\nRank correlations are not causal; fixed assumptions are absent from the ranking.',fontsize=9,color='#555')
    fig.subplots_adjust(left=.065,right=.97,bottom=.18,top=.79,wspace=.76)
    fig.savefig(folder/'distribution.png',dpi=165,facecolor=fig.get_facecolor())
    plt.close(fig)
    path=folder/'distribution.png'
    (folder/'chart_manifest.json').write_text(json.dumps({'matplotlib':matplotlib.__version__,'distribution.png':{'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'bytes':path.stat().st_size}},indent=2)+'\n')

if __name__=='__main__':
    main()
