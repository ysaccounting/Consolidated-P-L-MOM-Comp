from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
import io

SOLO_TAB_COMPANIES = {'YS Affiliates'}  # companies that get their own tab


def read_pl_sheet(ws):
    headers = None
    rows = []
    for row in ws.iter_rows(values_only=True):
        if headers is None and isinstance(row[1], str) and 'Tickets' in str(row[1]):
            headers = list(row)
        else:
            rows.append(list(row))
    return headers, rows


def extract_ordered_rows(rows):
    result = []
    seen = set()
    for row in rows:
        label = row[0]
        if label is None:
            continue
        vals = row[1:]
        if label not in seen:
            result.append((label, vals))
            seen.add(label)
    return result


def merge_label_orders(primary_labels, secondary_labels):
    known = set(primary_labels)
    result = list(primary_labels)
    for i, label in enumerate(secondary_labels):
        if label in known:
            continue
        insert_after = None
        for j in range(i - 1, -1, -1):
            if secondary_labels[j] in known:
                insert_after = secondary_labels[j]
                break
        if insert_after and insert_after in result:
            pos = result.index(insert_after) + 1
            result.insert(pos, label)
        else:
            result.append(label)
        known.add(label)
    return result


def write_pl_sheet(ws, companies, all_labels, mar_map, apr_map,
                   mar_label, apr_label, year,
                   SKIP_ROWS, TOTAL_ROWS, SECTION_DIVIDERS, SECTION_LABEL_ONLY):
    """Write a P&L sheet for the given list of (orig_idx, name) companies."""

    NAVY = "1F2D5A"
    TEAL = "006D6F"
    LTGRAY = "F2F3F5"
    WHITE = "FFFFFF"
    BLUE_HDR = "E3EAF5"
    TOTAL_BG = "C8D6EF"

    n_companies = len(companies)

    def company_start_col(ci):
        return 2 + ci * 5

    total_cols = 1 + n_companies * 5 - 1

    def row_style(label):
        if label == 'Gross Profit': return 'gross'
        if label == 'Net Operating Income': return 'noi'
        if label == 'Net Income': return 'ni'
        if label == 'Net Other Income': return 'netother'
        if label in TOTAL_ROWS: return 'total'
        if label in SECTION_DIVIDERS or label in SECTION_LABEL_ONLY: return 'section'
        return 'data'

    def has_any_data(label):
        mv = mar_map.get(label, [])
        av = apr_map.get(label, [])
        return (any(isinstance(v, (int, float)) and v != 0 for v in mv) or
                any(isinstance(v, (int, float)) and v != 0 for v in av))

    # Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    tc = ws.cell(row=1, column=1, value=f"Consolidated P&L — {mar_label} vs {apr_label} {year}")
    tc.font = Font(bold=True, size=14, color=WHITE, name="Arial")
    tc.fill = PatternFill("solid", fgColor=NAVY)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    # Company headers row
    ws.cell(row=2, column=1).fill = PatternFill("solid", fgColor=NAVY)
    for ci, (orig_idx, company) in enumerate(companies):
        cs = company_start_col(ci)
        hdr_color = "1F4E79" if company == 'Total' else TEAL
        ws.merge_cells(start_row=2, start_column=cs, end_row=2, end_column=cs + 3)
        c = ws.cell(row=2, column=cs, value=company)
        c.font = Font(bold=True, size=8, color=WHITE, name="Arial")
        c.fill = PatternFill("solid", fgColor=hdr_color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if ci < n_companies - 1:
            ws.cell(row=2, column=cs + 4).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[2].height = 30

    # Sub-headers
    ws.cell(row=3, column=1).fill = PatternFill("solid", fgColor=NAVY)
    for ci in range(n_companies):
        cs = company_start_col(ci)
        for j, hdr in enumerate([mar_label[:3], apr_label[:3], "$ Chg", "% Chg"]):
            c = ws.cell(row=3, column=cs + j, value=hdr)
            c.font = Font(bold=True, size=8, color=WHITE, name="Arial")
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.alignment = Alignment(horizontal="center")
        if ci < n_companies - 1:
            ws.cell(row=3, column=cs + 4).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[3].height = 15

    spacer_cols = {company_start_col(ci) + 4 for ci in range(n_companies - 1)}
    pct_cols = {company_start_col(ci) + 3 for ci in range(n_companies)}

    # Build set of labels that have at least one non-zero value for these companies
    def company_has_data(label):
        mv = mar_map.get(label, [])
        av = apr_map.get(label, [])
        for orig_idx, _ in companies:
            m = mv[orig_idx] if orig_idx < len(mv) else None
            a = av[orig_idx] if orig_idx < len(av) else None
            if (isinstance(m, (int, float)) and m != 0) or (isinstance(a, (int, float)) and a != 0):
                return True
        return False

    # Only always show the key P&L summary lines; everything else filters by data
    ALWAYS_SHOW = {'Gross Profit', 'Net Operating Income', 'Net Income', 'Net Other Income'}

    def should_show(label):
        if label in ALWAYS_SHOW:
            return True
        return company_has_data(label)

    data_start_row = 4
    used_rows = []
    gp_excel_row = None
    income_excel_row = None

    for label in all_labels:
        if label in SKIP_ROWS or label == 'Consolidated P&L':
            continue
        if not should_show(label):
            continue

        excel_row = data_start_row + len(used_rows)
        used_rows.append(label)
        style = row_style(label)

        if label == 'Gross Profit':
            gp_excel_row = excel_row
        if label == 'Total for Income':
            income_excel_row = excel_row

        if style == 'gross':
            bg, fg = "C8E6C9", "1B5E20"
        elif style == 'noi':
            bg, fg = "A5D6A7", "1B5E20"
        elif style == 'ni':
            bg, fg = TOTAL_BG, NAVY
        elif style == 'netother':
            bg, fg = "B3C6E7", NAVY
        elif style == 'total':
            bg, fg = "D6E0F5", NAVY
        elif style == 'section':
            bg, fg = BLUE_HDR, "1F2D5A"
        else:
            bg = WHITE if len(used_rows) % 2 == 0 else LTGRAY
            fg = "000000"

        lc = ws.cell(row=excel_row, column=1, value=label)
        lc.font = Font(name="Arial", size=9, bold=(style != 'data'), color=fg)
        lc.fill = PatternFill("solid", fgColor=bg)
        lc.alignment = Alignment(horizontal="left", vertical="center",
                                  indent=(1 if style == 'data' else 0))

        mar_vals = mar_map.get(label, [])
        apr_vals = apr_map.get(label, [])
        write_data = style != 'section' or has_any_data(label)

        for ci, (orig_idx, company) in enumerate(companies):
            cs = company_start_col(ci)
            mv = mar_vals[orig_idx] if orig_idx < len(mar_vals) else None
            av = apr_vals[orig_idx] if orig_idx < len(apr_vals) else None
            mv = mv if isinstance(mv, (int, float)) else None
            av = av if isinstance(av, (int, float)) else None

            delta = pct = None
            if write_data:
                if mv is not None and av is not None:
                    delta = av - mv
                elif av is not None:
                    delta = av
                elif mv is not None:
                    delta = -mv
                pct = (delta / abs(mv)) if (mv is not None and mv != 0 and delta is not None) else None

            cell_bg = bg
            if company == 'Total' and style == 'data':
                cell_bg = "EEF2FA" if len(used_rows) % 2 == 0 else "E4EAF7"

            for j, (val, fmt) in enumerate([
                (mv if write_data else None, '#,##0;(#,##0);"-"'),
                (av if write_data else None, '#,##0;(#,##0);"-"'),
                (delta, '#,##0;(#,##0);"-"'),
                (pct, '0.00%;(0.00%);"-"'),
            ]):
                cell = ws.cell(row=excel_row, column=cs + j)
                cell.fill = PatternFill("solid", fgColor=cell_bg)
                cell.font = Font(name="Arial", size=9, bold=(style != 'data'), color=fg)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if val is not None:
                    cell.value = val
                    cell.number_format = fmt

            if ci < n_companies - 1:
                ws.cell(row=excel_row, column=cs + 4).fill = PatternFill("solid", fgColor="CCCCCC")

        ws.row_dimensions[excel_row].height = 14

    # GPM row
    if gp_excel_row and income_excel_row:
        gpm_excel_row = gp_excel_row + 1
        ws.insert_rows(gpm_excel_row)
        bg_gpm, fg_gpm = "C8E6C9", "1B5E20"
        lc = ws.cell(row=gpm_excel_row, column=1, value="Gross Profit Margin")
        lc.font = Font(name="Arial", size=9, bold=True, color=fg_gpm)
        lc.fill = PatternFill("solid", fgColor=bg_gpm)
        lc.alignment = Alignment(horizontal="left", vertical="center")

        for ci in range(n_companies):
            cs = company_start_col(ci)
            for j, src_col in enumerate([cs, cs + 1]):
                cell = ws.cell(row=gpm_excel_row, column=cs + j)
                cell.value = f"={get_column_letter(src_col)}{gp_excel_row}/{get_column_letter(src_col)}{income_excel_row}"
                cell.number_format = '0.00%;(0.00%);"-"'
                cell.font = Font(name="Arial", size=9, bold=True, color=fg_gpm)
                cell.fill = PatternFill("solid", fgColor=bg_gpm)
                cell.alignment = Alignment(horizontal="right", vertical="center")

            dc = ws.cell(row=gpm_excel_row, column=cs + 2)
            dc.fill = PatternFill("solid", fgColor=bg_gpm)
            dc.font = Font(name="Arial", size=9, bold=True, color=fg_gpm)

            pc = ws.cell(row=gpm_excel_row, column=cs + 3)
            pc.value = f"={get_column_letter(cs + 1)}{gpm_excel_row}-{get_column_letter(cs)}{gpm_excel_row}"
            pc.number_format = '0.00%;(0.00%);"-"'
            pc.font = Font(name="Arial", size=9, bold=True, color=fg_gpm)
            pc.fill = PatternFill("solid", fgColor=bg_gpm)
            pc.alignment = Alignment(horizontal="right", vertical="center")

            if ci < n_companies - 1:
                ws.cell(row=gpm_excel_row, column=cs + 4).fill = PatternFill("solid", fgColor="CCCCCC")

        ws.row_dimensions[gpm_excel_row].height = 14

    # Column widths
    col_widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            col = cell.column
            if col in spacer_cols:
                continue
            val = cell.value
            if val is None or (isinstance(val, str) and val.startswith('=')):
                continue
            if col in pct_cols and isinstance(val, float):
                rendered = f"({abs(val * 100):,.2f}%)" if val < 0 else f"{val * 100:,.2f}%"
            elif isinstance(val, (int, float)):
                rendered = f"({abs(val):,.0f})" if val < 0 else f"{val:,.0f}"
            else:
                rendered = str(val)
            col_widths[col] = max(col_widths.get(col, 0), len(rendered))

    for col, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(max(width + 3, 8), 40)
    for col in spacer_cols:
        ws.column_dimensions[get_column_letter(col)].width = 3

    ws.freeze_panes = "B4"


def build_combined_pl(mar_file, apr_file, mar_label="March", apr_label="April", year="2026"):
    mar_wb = load_workbook(mar_file, data_only=True)
    apr_wb = load_workbook(apr_file, data_only=True)
    mar_ws = mar_wb.active
    apr_ws = apr_wb.active

    mar_headers, mar_rows = read_pl_sheet(mar_ws)
    apr_headers, apr_rows = read_pl_sheet(apr_ws)

    all_companies = mar_headers[1:]

    # Split into main companies and solo-tab companies
    main_companies = [(i, name) for i, name in enumerate(all_companies)
                      if name != 'Eliminations' and name not in SOLO_TAB_COMPANIES]
    solo_companies = [(i, name) for i, name in enumerate(all_companies)
                      if name in SOLO_TAB_COMPANIES]
    # Total always goes last on main sheet
    total_companies = [(i, name) for i, name in main_companies if name == 'Total']
    main_companies = [(i, name) for i, name in main_companies if name != 'Total'] + total_companies

    mar_ordered = extract_ordered_rows(mar_rows)
    apr_ordered = extract_ordered_rows(apr_rows)
    mar_map = {r[0]: r[1] for r in mar_ordered}
    apr_map = {r[0]: r[1] for r in apr_ordered}

    mar_labels = [r[0] for r in mar_ordered]
    apr_labels = [r[0] for r in apr_ordered]
    all_labels = merge_label_orders(mar_labels, apr_labels)

    SECTION_LABEL_ONLY = {'Miscellaneous Fees', 'Professional Fees', 'Salaries', 'K-1 Income'}
    SKIP_ROWS = {'Consolidated P&L', 'March 2026', 'April 2026',
                 f'{mar_label} {year}', f'{apr_label} {year}'}
    TOTAL_ROWS = {l for l in all_labels if l and str(l).startswith('Total for')}
    SECTION_DIVIDERS = {'Income', 'Cost of Goods Sold', 'Expenses', 'Other Income', 'Other Expenses'}

    shared_args = (all_labels, mar_map, apr_map, mar_label, apr_label, year,
                   SKIP_ROWS, TOTAL_ROWS, SECTION_DIVIDERS, SECTION_LABEL_ONLY)

    wb = Workbook()

    # Main sheet
    ws_main = wb.active
    ws_main.title = "Combined P&L"
    write_pl_sheet(ws_main, main_companies, *shared_args)

    # Solo tabs
    for orig_idx, name in solo_companies:
        tab_title = name.replace(' Tickets', '').replace(' LLC', '').strip()[:31]
        ws_solo = wb.create_sheet(title=tab_title)
        write_pl_sheet(ws_solo, [(orig_idx, name)], *shared_args)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
