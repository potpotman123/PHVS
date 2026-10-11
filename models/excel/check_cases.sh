#!/bin/zsh
# Build, then reconcile each case against the engine in a disposable copy.
set -e
cd "$(dirname "$0")/../.."
RECALC="${RECALC:-/private/tmp/claude-501/-Users-kwu-Library-Application-Support-Claude-scratch-workspaces-665f01e2-9652-4512-9fda-0497b55d354c-06f88a99-9105-4941-a8cf-87e6b02aff65-scratch-2026-10-10-7ee708/b9401f11-15ac-4c44-8c9e-8eb278b73abd/scratchpad/xlsx_scripts/recalc.py}"
python3 models/excel/build_model.py | tail -1
python3 "$RECALC" models/excel/PHVS_Crimson_Capital_Model.xlsx 900 | python3 -c "import sys,json; d=json.load(sys.stdin); print('recalc', d.get('status'), d.get('total_formulas'), 'errors', d.get('total_errors'))"
for CASE in working cautious upside_challenge; do
  LABEL=$(python3 -c "print({'working':'Working','cautious':'Cautious','upside_challenge':'Upside'}['$CASE'])")
  python3 - "$CASE" "$LABEL" <<'PY'
import sys, openpyxl, warnings; warnings.filterwarnings('ignore')
case, label = sys.argv[1], sys.argv[2]
import json
wb = openpyxl.load_workbook('models/excel/PHVS_Crimson_Capital_Model.xlsx'); maps = json.load(open('models/excel/rowmap.json')); cfg = json.load(open('models/finalized/inputs.json'))
phys, off = maps['OFFS'].get('Inputs', ['Inputs', 0]); ws = wb[phys]
col = {'cautious': 6, 'upside_challenge': 7, 'working': 5}[case]
for k, rr in maps['INPUT_ROW'].items():
    v = ws.cell(rr + off, col).value
    if v is not None and not (isinstance(v, str) and v.startswith('=')): ws.cell(rr + off, 9).value = v
wb.save(f'/tmp/phvs_{case}.xlsx')
PY
  python3 "$RECALC" /tmp/phvs_$CASE.xlsx 900 > /dev/null
  echo "=== $CASE ==="; python3 models/excel/reconcile.py $CASE /tmp/phvs_$CASE.xlsx | grep -v '^OK ' 
done
