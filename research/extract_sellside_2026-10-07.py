from pathlib import Path
import zipfile, hashlib, json, re
from pypdf import PdfReader
import pdfplumber

root = Path('/Users/kwu/PHVS')
dest = root / 'references/user_supplied/sellside_2026-10-07'
raw = root / 'research/raw/sellside_2026-10-07'
dest.mkdir(parents=True, exist_ok=True)
raw.mkdir(parents=True, exist_ok=True)
manifest = []
for archive in ['drive-download-20261006T203953Z-1-001.zip', 'drive-download-20261006T204003Z-1-001.zip']:
    with zipfile.ZipFile(Path('/Users/kwu/Downloads') / archive) as z:
        for info in z.infolist():
            if not info.filename.lower().endswith('.pdf'):
                continue
            ident = re.search(r'-(\d+)\.pdf$', info.filename).group(1)
            content = z.read(info)
            pdf = dest / (ident + '.pdf')
            pdf.write_bytes(content)
            textpath = raw / (ident + '.txt')
            with pdfplumber.open(pdf) as doc:
                pages=[p.extract_text(layout=True) or '' for p in doc.pages]
            textpath.write_text('\n\n'.join(f'[PDF PAGE {i+1}]\n{t}' for i,t in enumerate(pages)))
            manifest.append(dict(id=ident, archive=archive, original_name=info.filename,
                                 pdf=str(pdf.relative_to(root)), text=str(textpath.relative_to(root)),
                                 pages=len(pages), sha256=hashlib.sha256(content).hexdigest()))
            print(f'{ident} | {len(pages)} pages | {info.filename}')
            hits=[]
            for i,t in enumerate(pages[:7]):
                for line in t.splitlines():
                    if re.search(r'probability|POS|WACC|discount rate|DCF|price target|peak sales|peak revenue|risk.adjust|valuation',line,re.I):
                        hits.append(f'  p{i+1}: {line.strip()[:220]}')
            print('\n'.join(hits[:12]))
(raw/'manifest.json').write_text(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n')
