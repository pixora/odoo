#!/usr/bin/env python3
"""Genera i template di importazione PRECOMPILATI con i dati demo del repository
pixora/demo-alina (moduli pix_demo_family_office e pix_demo_family_office_hr_doc).

Uso:  python3 genera_demo_alina.py /percorso/demo-alina
Output: cartella demo_alina/ con i file DEMO_ALINA_*.xlsx

Gli ID esterni sono quelli dei moduli demo (es. pix_demo_family_office.partner_energia_nord):
se il modulo è installato l'import AGGIORNA i record esistenti invece di duplicarli.
Tutti i dati sono fittizi (provengono dal modulo demo).
"""
import ast
import copy
import os
import sys
import xml.etree.ElementTree as ET

from openpyxl.styles import PatternFill

from genera_template import HERE, TEMPLATES, build

MOD = 'pix_demo_family_office'
MOD_HR = 'pix_demo_family_office_hr_doc'
DEMO_FILL = PatternFill('solid', fgColor='E2F0D9')
NOTE = ('Rosso = obbligatorio · Blu = consigliato · Grigio = facoltativo. Le righe verdi sono i DATI DEMO del '
        'family office (repo pixora/demo-alina): sono pronte per l\'import. Gli ID esterni coincidono con quelli dei '
        'moduli demo: se il modulo è già installato l\'import aggiorna i record senza duplicarli.')


def constants(path, names):
    """Legge le costanti (liste/dict letterali) da un file Python senza eseguirlo."""
    tree = ast.parse(open(path, encoding='utf-8').read())
    out = {}
    for node in tree.body:
        if isinstance(node, ast.Assign) and isinstance(node.targets[0], ast.Name) and node.targets[0].id in names:
            out[node.targets[0].id] = ast.literal_eval(node.value)
    return out


def xml_records(path, model):
    root = ET.parse(path).getroot()
    for rec in root.iter('record'):
        if rec.get('model') != model:
            continue
        vals = {}
        for f in rec.findall('field'):
            vals[f.get('name')] = f.get('ref') or f.get('eval') or (f.text or '')
        yield rec.get('id'), vals


def tpl(file_name):
    t = copy.deepcopy(next(t for t in TEMPLATES if t['file'] == file_name))
    t['file'] = 'DEMO_ALINA_' + file_name
    t['rows_note'] = NOTE
    return t


def main(repo):
    base = os.path.join(repo, MOD)
    partners_xml = os.path.join(base, 'data', 'res_partner_data.xml')
    analytic_xml = os.path.join(base, 'data', 'account_analytic_data.xml')
    c = constants(os.path.join(base, 'hooks.py'), {'VENDORS', 'BUDGET_LINES', 'VENDOR_ANALYTICS'})
    c_hr = constants(os.path.join(repo, MOD_HR, 'hooks.py'), {'EMPLOYEES'})
    out_dir = os.path.join(HERE, 'demo_alina')
    os.makedirs(out_dir, exist_ok=True)
    built = []

    # 01 Fornitori ---------------------------------------------------------------
    t = tpl('01_Fornitori.xlsx')
    fields = [col[0] for col in t['cols']]
    vendor_cat = {v: cat for v, _prop, cat, _nat in c['VENDOR_ANALYTICS']}
    rows = []
    for xmlid, v in xml_records(partners_xml, 'res.partner'):
        row = dict.fromkeys(fields, '')
        key = xmlid.removeprefix('partner_')
        row.update({
            'id': f'{MOD}.{xmlid}',
            'is_company': 'TRUE' if v.get('is_company') == 'True' else 'FALSE',
            'name': v.get('name', ''),
            'vat': v.get('vat', ''),
            'street': v.get('street', ''),
            'zip': v.get('zip', ''),
            'city': v.get('city', ''),
            'state_id/id': 'base.state_it_to' if v.get('city') == 'Torino' else (
                'base.state_it_mi' if v.get('city') == 'Milano' else ''),
            'country_id/id': v.get('country_id', ''),
            'phone': v.get('phone', ''),
            'email': v.get('email', ''),
            'supplier_rank': int(v.get('supplier_rank') or 1),
            'property_supplier_payment_term_id/id': 'account.account_payment_term_30days',
            'category_id': vendor_cat.get(key, '').capitalize(),
            'lang': 'it_IT',
        })
        rows.append([row[f] for f in fields])
    t['examples'] = rows
    built.append(build(t, out_dir, DEMO_FILL))

    # 02 Conti bancari -----------------------------------------------------------
    t = tpl('02_Conti_Bancari.xlsx')
    fields = [col[0] for col in t['cols']]
    rows = []
    for xmlid, v in xml_records(partners_xml, 'res.partner.bank'):
        row = dict.fromkeys(fields, '')
        row.update({
            'id': f'{MOD}.{xmlid}',
            'partner_id/id': f"{MOD}.{v['partner_id']}",
            'account_number': v.get('account_number', ''),
            'bank_name': v.get('bank_name', ''),
            'bank_bic': v.get('bank_bic', ''),
            'allow_out_payment': 'TRUE' if v.get('allow_out_payment') == 'True' else 'FALSE',
        })
        rows.append([row[f] for f in fields])
    t['examples'] = rows
    built.append(build(t, out_dir, DEMO_FILL))

    # 03 Dipendenti --------------------------------------------------------------
    t = tpl('03_Dipendenti.xlsx')
    fields = [col[0] for col in t['cols']]
    people = [('emp_resp_amministrativo', 'Responsabile Amministrativo (demo)', 'Responsabile amministrativo', '')]
    people += [(x, n, j, f'{MOD_HR}.emp_resp_amministrativo') for x, n, j in c_hr['EMPLOYEES']]
    rows = []
    for xmlid, name, job, parent in people:
        row = dict.fromkeys(fields, '')
        row.update({
            'id': f'{MOD_HR}.{xmlid}',
            'name': name,
            'job_title': job,
            'parent_id/id': parent,
            'employee_type_id/id': 'hr.contract_type_employee',
            'work_email': f"{xmlid.removeprefix('emp_').replace('_', '.')}.demo@example.com" if parent else '',
            'country_id/id': 'base.it',
            'private_country_id/id': 'base.it',
        })
        rows.append([row[f] for f in fields])
    t['examples'] = rows
    built.append(build(t, out_dir, DEMO_FILL))

    # 04 Piano dei conti ---------------------------------------------------------
    # Il modulo demo usa solo conti standard l10n_it: proponiamo sottoconti dedicati al family office
    t = tpl('04_Piano_dei_Conti.xlsx')
    t['examples'] = [
        ['DEMO_CONTO_4214', '4214', 'Costi di giardinaggio (demo)', 'expense', 'FALSE', 'FALSE', 'l10n_it.account_tag_B_PL', '', 'Giardini e aree verdi delle residenze', ''],
        ['DEMO_CONTO_4215', '4215', 'Costi di guardaroba e lavanderia (demo)', 'expense', 'FALSE', 'FALSE', 'l10n_it.account_tag_B_PL', '', '', ''],
        ['DEMO_CONTO_4216', '4216', 'Manutenzioni straordinarie immobili (demo)', 'expense', 'FALSE', 'FALSE', 'l10n_it.account_tag_B_PL', '', 'Lavori straordinari soggetti ad approvazione del Principal', ''],
        ['DEMO_CONTO_4217', '4217', 'Costi di cucina e dispensa (demo)', 'expense', 'FALSE', 'FALSE', 'l10n_it.account_tag_B_PL', '', '', ''],
        ['DEMO_CONTO_1530', '1530', 'Anticipi al personale di casa (demo)', 'asset_current', 'TRUE', 'TRUE', 'l10n_it.account_tag_C_ATT', '', 'Fondi spese consegnati a chef/autisti', ''],
    ]
    built.append(build(t, out_dir, DEMO_FILL))

    # 05 Budget analitico --------------------------------------------------------
    names = {xmlid: v['name'] for xmlid, v in xml_records(analytic_xml, 'account.analytic.account')}
    plans = {xmlid: v['name'] for xmlid, v in xml_records(analytic_xml, 'account.analytic.plan')}
    t = tpl('05_Budget_Analitico.xlsx')
    head = [col for col in t['cols'] if not col[0].startswith('budget_line_ids/')]
    amount = next(col for col in t['cols'] if col[0] == 'budget_line_ids/budget_amount')
    plan_cols = [
        (f"budget_line_ids/{plans['plan_proprieta']}", f"Riga: piano {plans['plan_proprieta']}", 'C', 'Nome conto analitico', 'Residenza Torino',
         'Colonna indicata con il NOME del piano analitico: Odoo la abbina al campo x_planN_id corrispondente. Verificare l\'abbinamento nella schermata di import.', 20, None),
        (f"budget_line_ids/{plans['plan_categoria']}", f"Riga: piano {plans['plan_categoria']}", 'F', 'Nome conto analitico', 'Utenze', 'Come sopra.', 18, None),
        (f"budget_line_ids/{plans['plan_natura']}", f"Riga: piano {plans['plan_natura']}", 'C', 'Nome conto analitico', 'Ordinario', 'Come sopra.', 14, None),
    ]
    t['cols'] = head + plan_cols + [amount]
    rows = []
    for i, (prop, cat, nat, amt) in enumerate(c['BUDGET_LINES']):
        first = ['DEMO_BDG_2026_TORINO', 'Budget 2026 Ufficio Torino (demo)', '2026-01-01', '2026-12-31', 'expense', '', ''] if i == 0 else [''] * 7
        rows.append(first + [names[f'aa_prop_{prop}'], names[f'aa_cat_{cat}'] if cat else '', names[f'aa_nat_{nat}'], amt])
    t['examples'] = rows
    built.append(build(t, out_dir, DEMO_FILL))

    # 06 Budget finanziario ------------------------------------------------------
    # Dai costi mensili tipici dei fornitori demo (VENDORS) -> budget per conto per mese 2026
    monthly = {}
    for _vendor, code, amt in c['VENDORS']:
        monthly[code] = monthly.get(code, 0.0) + amt
    coa_names = {}
    import csv  # noqa: PLC0415
    from genera_template import L10N_IT_COA  # noqa: PLC0415
    with open(L10N_IT_COA, encoding='utf-8') as f:
        coa_names = {r['code']: r['name@it'] for r in csv.DictReader(f)}
    t = tpl('06_Budget_Finanziario.xlsx')
    rows = []
    first = True
    for code in sorted(monthly):
        for m in range(1, 13):
            head = ['DEMO_BUDFIN_2026', 'Budget 2026 Ufficio Torino (demo)', ''] if first else ['', '', '']
            rows.append(head + [code, f'2026-{m:02d}-01', monthly[code]])
            first = False
    t['examples'] = rows
    t['matrix'] = [[code, coa_names.get(code, '')] + [monthly[code]] * 12 for code in sorted(monthly)]
    built.append(build(t, out_dir, DEMO_FILL))

    for b in built:
        print('creato', b)


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, '..', '..', 'demo-alina'))
