"""Generate and export the updated high-resolution VAT reconciliation charts.

Palette & Design System:
- Figure 1 (Source Volume & Status):
    * New Invoices: #326B77 (Deep Maritime Teal anchor)
    * Adjustment Invoices: #D9822B (Rich Golden Amber, with #A05206 label text)
    * Replacement Invoices: #6D3507 (Deep Espresso Amber - darker same hue)
    * Cancelled Invoices: #1E293B (Deep Slate / Navy)
    * Other: #CBD5E1 (Neutral Classic Grey)

- Figure 2 (Reconciliation Results - Separate Audit Family):
    * Matching Amounts: #A0AEC0 (Silent Cool Slate - quiet baseline)
    * Amount Diff / Missing Values: #EA580C (Vibrant Terracotta / Orange)
    * Unmatched / Duplicate-Key Records: #991B1B (Deep Crimson Wine)

- Layout Polish:
    * GDT callout positioned at x = 0.28, MISA callout at x = 1.52 (clear horizontal gap)
    * Compact Amount Diff labels keep callouts clear of totals and bar counts.
    * Rounded badge borders and bold inner numbers for clean executive presentation
"""
import sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from pathlib import Path
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# Local module import
import vat_reconciliation as vr

FOLDER = Path(__file__).resolve().parent
ROOT = FOLDER.parent
PUBLIC = FOLDER if (FOLDER / 'demo_config.json').is_file() else ROOT / 'portfolio-artifacts/04-python-vat-reconciliation'
MONTH = '2026-08'


def export_charts(output_dir=None):
    """Generate both Figure 1 and Figure 2 using demo data and save to output_dir."""
    if output_dir is None:
        output_dir = FOLDER / 'reports' / MONTH
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    # Load public demo dataset
    tax_code = json.loads((PUBLIC / 'demo_config.json').read_text())['tax_code']
    gdt, diagnostics = vr.load_gdt_folder(PUBLIC, MONTH, tax_code)
    if not diagnostics.status.eq('loaded').all():
        raise ValueError('All four GDT source files are required.')
    
    misa = vr.read_misa(PUBLIC / 'MISA_VAT_Return_Schedules.xlsx', MONTH)
    results = {d: vr.reconcile(gdt, misa, MONTH, d) for d in ('purchases', 'sales')}

    # 1. Figure 1: Invoice Volume by Source and Status
    status_summary = vr.source_status_overview(gdt, misa)
    fig1 = vr.plot_comparison(status_summary, MONTH, matching=False, palette='modern')
    chart1_path = output_dir / 'invoice-status-comparison.png'
    vr.save_chart(fig1, chart1_path)
    print(f'Exported Figure 1: {chart1_path.name} -> {chart1_path.parent.name}')

    # 2. Figure 2: Reconciliation Results (Audit Overview)
    match_summary = vr.matching_overview(results, gdt, misa)
    fig2 = vr.plot_comparison(match_summary, MONTH, matching=True, palette='modern')
    chart2_path = output_dir / 'reconciliation-overview.png'
    vr.save_chart(fig2, chart2_path)
    print(f'Exported Figure 2: {chart2_path.name} -> {chart2_path.parent.name}')

    plt.close('all')
    return chart1_path, chart2_path


if __name__ == '__main__':
    # 1. Save in VAT reconciliation/reports/2026-08/
    export_charts(FOLDER / 'reports' / MONTH)

    # 2. Also keep public portfolio assets in sync
    assets_dir = PUBLIC / 'assets'
    assets_dir.mkdir(parents=True, exist_ok=True)
    export_charts(assets_dir)

    print('\nAll updated charts generated and saved successfully.')
