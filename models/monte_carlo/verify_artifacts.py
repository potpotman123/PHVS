"""Validate committed output tables against summaries, samples and live engine replay."""
import gzip
import hashlib
import json
from pathlib import Path
import statistics
import run as mc
from package_results import compact

def read_gz(path):
    return json.loads(gzip.decompress(path.read_bytes()))

def main():
    root=mc.HERE/'results'
    manifest=json.loads((root/'artifact_manifest.json').read_text())
    for name,meta in manifest.items():
        data=(root/name).read_bytes()
        mc.require(len(data)==meta['bytes'] and hashlib.sha256(data).hexdigest()==meta['sha256'],'Artifact integrity: '+name)
    results=[]
    for file in sorted(root.glob('seed_*/*/draws_compact.json.gz')):
        c=read_gz(file);ix={k:i for i,k in enumerate(c['columns'])};rows=c['rows']
        s=json.loads((file.parent/'summary.json').read_text())
        mc.require(s['hashes']==mc.hashes(),'Current model differs from report hashes')
        mc.require(len(rows)==s['draws'],'Output row count')
        mc.require([r[ix['draw_id']] for r in rows]==list(range(len(rows))),'Output draw identity')
        pairs=[]
        for row in rows:
            expected=sum(row[ix[f'{state}.target.price']]*w for state,w in c['clinical_weights'].items())
            mc.require(abs(expected-row[ix['expected_target']])<1e-9,'Published clinical weighting')
            pairs.extend((row[ix[f'{state}.target.price']],w/len(rows)) for state,w in c['clinical_weights'].items())
        values=[r[ix['expected_target']] for r in rows]
        mc.require(mc.describe(values)==s['expected_target'],'Published expected summary')
        for metric,q in [('p10',.1),('median',.5),('p90',.9)]:
            mc.require(mc.weighted_quantile(pairs,q)==s['clinical_outcomes_target'][metric],'Published outcome quantile')
        # Recompute draw zero, the median and both extremes for every scenario.
        picks={0,values.index(mc.quantile(values,.5)),values.index(min(values)),values.index(max(values))}
        for i,q in mc.draws(len(rows),s['seed']):
            if i not in picks: continue
            r=mc.evaluate_task((i,q,s['access'],s['structural'],s['dependence']))
            mc.require('error' not in r,'Replay engine error')
            mc.require(compact([r])['rows'][0]==rows[i],'Engine replay differs from published output')
        results.append({'scenario':str(file.parent.relative_to(root)),'rows':len(rows),'replayed_draw_ids':sorted(picks),'passed':True})
    for file in root.glob('seed_*/primitive_inputs.json.gz'):
        c=read_gz(file);seed=int(file.parent.name.split('_')[1]);keys=c['columns'][1:]
        for (i,q),row in zip(mc.draws(len(c['rows']),seed),c['rows']):
            p=mc.overrides_at(q,mc.SPEC)
            mc.require(row==[i]+[p[k] for k in keys],'Published primitive input differs')
    mc.dump(root/'artifact_verification.json',{'passed':True,'artifact_hashes_verified':len(manifest),'scenario_checks':results,
       'total_output_rows_checked':sum(r['rows'] for r in results),'exact_engine_replays':sum(len(r['replayed_draw_ids']) for r in results),
       'verification_file_note':'This file is generated after artifact_manifest.json and is not self-hashed.'})
    print(json.dumps({'passed':True,'scenarios':len(results),'rows':sum(r['rows'] for r in results),'replays':sum(len(r['replayed_draw_ids']) for r in results)}))

if __name__=='__main__':
    main()
