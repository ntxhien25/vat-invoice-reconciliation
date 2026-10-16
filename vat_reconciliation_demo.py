"""Run the working reconciliation against the five masked Excel files."""
from pathlib import Path
import json
import matplotlib.pyplot as plt
import vat_reconciliation as vr
ROOT = Path(__file__).resolve().parent
MONTH = '2026-08'

def run_demo(root=ROOT):
    root = Path(root).resolve()
    tax_code = json.loads((root/'demo_config.json').read_text())['tax_code']
    gdt, diagnostics = vr.load_gdt_folder(root, MONTH, tax_code)
    if not diagnostics.status.eq('loaded').all():raise ValueError('All four GDT source files are required.')
    misa = vr.read_misa(root/'MISA_VAT_Return_Schedules.xlsx', MONTH)
    results = {d: vr.reconcile(gdt,misa,MONTH,d) for d in ('purchases','sales')}
    report_dir = root/'reports'/MONTH;report_dir.mkdir(parents=True,exist_ok=True)
    (root/'assets').mkdir(exist_ok=True)
    for direction,title in (('purchases','Input'),('sales','Output')):
        vr.export_reconciliation_workbook(results[direction],report_dir/f'{title}_VAT_Reconciliation_{MONTH}.xlsx')
    vr.save_chart(vr.plot_comparison(vr.source_status_overview(gdt,misa),MONTH),root/'assets/invoice-status-comparison.png')
    summary = vr.matching_overview(results,gdt,misa)
    vr.save_chart(vr.plot_comparison(summary,MONTH,matching=True),root/'assets/reconciliation-overview.png')
    plt.close('all')
    return {'gdt':gdt,'misa':misa},results,summary

if __name__ == '__main__':
    _,_,summary = run_demo()
    print(summary.to_string(index=False))
