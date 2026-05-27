import os
import sys
import json
import traceback
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "Define Studio"
APP_SUBTITLE = "Spec Parser and Review Utility"
APP_VERSION = "v2.0"

# --------------------------
# Optional dependency loader
# --------------------------

def ensure_package(pkg_name, import_name=None):
    import_name = import_name or pkg_name
    try:
        return __import__(import_name)
    except ImportError:
        import subprocess
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg_name])
            return __import__(import_name)
        except Exception as exc:
            raise RuntimeError(f"Unable to install required package '{pkg_name}'. Error: {exc}")

pd = ensure_package('pandas')
try:
    openpyxl = __import__('openpyxl')
except Exception:
    openpyxl = ensure_package('openpyxl')

# --------------------------
# Theme - toned down, more professional
# --------------------------
THEME = {
    'bg': '#f3f6fb',
    'panel': '#ffffff',
    'panel_alt': '#eef3f9',
    'card': '#ffffff',
    'card_soft': '#f7f9fc',
    'header': '#1f3a5f',
    'header_accent': '#4f6f92',
    'line': '#d7dee8',
    'text': '#1f2937',
    'muted': '#60758a',
    'accent': '#385b84',
    'accent_soft': '#e8eff7',
    'accent2': '#6f87a3',
    'accent3': '#7f8ea3',
    'success': '#557a68',
    'warning': '#8c6b49',
    'selection': '#dbe7f4',
    'selection_text': '#102a43',
    'entry': '#fbfcfe',
    'entry_text': '#1f2937',
    'shadow': '#e5ebf3',
}

# --------------------------
# Utility helpers
# --------------------------

def clean_col(c):
    return str(c).strip().lower().replace('\n', ' ').replace('\r', ' ')


def standardize_columns(df):
    df = df.copy()
    df.columns = [clean_col(c) for c in df.columns]
    return df


def pick_first_existing(df, options):
    cols = {clean_col(c): c for c in df.columns}
    for opt in options:
        if clean_col(opt) in cols:
            return cols[clean_col(opt)]
    return None


def coerce_text(x):
    if pd.isna(x):
        return ''
    return str(x).strip()


def rounded_polygon_points(x1, y1, x2, y2, radius=18):
    r = max(0, min(radius, abs(x2 - x1) / 2, abs(y2 - y1) / 2))
    return [
        x1 + r, y1,
        x2 - r, y1,
        x2, y1,
        x2, y1 + r,
        x2, y2 - r,
        x2, y2,
        x2 - r, y2,
        x1 + r, y2,
        x1, y2,
        x1, y2 - r,
        x1, y1 + r,
        x1, y1,
    ]


class RoundedCard(tk.Canvas):
    def __init__(self, master, bg_color, border_color=None, radius=18, padding=1, **kwargs):
        super().__init__(master, highlightthickness=0, bd=0, bg=master.cget('bg'), **kwargs)
        self.bg_color = bg_color
        self.border_color = border_color or bg_color
        self.radius = radius
        self.padding = padding
        self.inner = tk.Frame(self, bg=bg_color)
        self.window_id = self.create_window((0, 0), window=self.inner, anchor='nw')
        self.bind('<Configure>', self._redraw)

    def _redraw(self, event=None):
        self.delete('shape')
        w = max(2, self.winfo_width())
        h = max(2, self.winfo_height())
        p = self.padding
        pts = rounded_polygon_points(p, p, w - p, h - p, self.radius)
        self.create_polygon(pts, smooth=True, fill=self.border_color, outline=self.border_color, tags='shape')
        if self.padding > 1:
            p2 = self.padding + 1
            pts2 = rounded_polygon_points(p2, p2, w - p2, h - p2, max(0, self.radius - 1))
            self.create_polygon(pts2, smooth=True, fill=self.bg_color, outline=self.bg_color, tags='shape')
        self.coords(self.window_id, self.padding + 10, self.padding + 10)
        self.itemconfigure(self.window_id, width=max(10, w - (self.padding + 10) * 2), height=max(10, h - (self.padding + 10) * 2))
        self.tag_lower('shape')


class SpecParser:
    def __init__(self):
        self.file_path = None
        self.workbook = {}
        self.summary_df = pd.DataFrame()
        self.domain_map = {}
        self.domain_rows = {}

    def reset(self):
        self.file_path = None
        self.workbook = {}
        self.summary_df = pd.DataFrame()
        self.domain_map = {}
        self.domain_rows = {}

    def load_excel(self, file_path):
        self.reset()
        self.file_path = file_path
        xls = pd.ExcelFile(file_path)
        self.workbook = {sheet: standardize_columns(pd.read_excel(file_path, sheet_name=sheet)) for sheet in xls.sheet_names}
        self._build_model()

    def _build_model(self):
        domains_sheet_name = None
        for s in self.workbook:
            if clean_col(s) in ('domains', 'domain', 'datasets', 'dataset'):
                domains_sheet_name = s
                break

        if domains_sheet_name:
            ddf = self.workbook[domains_sheet_name].copy()
            dataset_col = pick_first_existing(ddf, ['dataset', 'domain', 'name'])
            label_col = pick_first_existing(ddf, ['description', 'dataset label', 'label'])
            class_col = pick_first_existing(ddf, ['class', 'domain class'])
            structure_col = pick_first_existing(ddf, ['structure'])
            purpose_col = pick_first_existing(ddf, ['purpose'])
            rows = []
            if dataset_col:
                for _, row in ddf.iterrows():
                    ds = coerce_text(row.get(dataset_col, ''))
                    if not ds:
                        continue
                    entry = {
                        'dataset': ds.upper(),
                        'label': coerce_text(row.get(label_col, '')) if label_col else '',
                        'class': coerce_text(row.get(class_col, '')) if class_col else '',
                        'structure': coerce_text(row.get(structure_col, '')) if structure_col else '',
                        'purpose': coerce_text(row.get(purpose_col, '')) if purpose_col else '',
                        'sheet': ds,
                    }
                    rows.append(entry)
                self.summary_df = pd.DataFrame(rows)

        for sname, sdf in self.workbook.items():
            if self.summary_df.shape[0] > 0 and sname.upper() in set(self.summary_df['dataset'].str.upper()):
                domain = sname.upper()
            elif len(sname) <= 8 and sname.upper().isalnum() and sname.upper() not in ('DOMAINS', 'DATASETS'):
                domain = sname.upper()
            else:
                continue

            var_col = pick_first_existing(sdf, ['variable', 'name', 'variable name'])
            label_col = pick_first_existing(sdf, ['label', 'variable label', 'description'])
            type_col = pick_first_existing(sdf, ['type', 'datatype', 'data type'])
            length_col = pick_first_existing(sdf, ['length', 'display length'])
            format_col = pick_first_existing(sdf, ['format'])
            codelist_col = pick_first_existing(sdf, ['codelist', 'controlled terms', 'terms'])
            origin_col = pick_first_existing(sdf, ['origin'])
            role_col = pick_first_existing(sdf, ['role'])
            core_col = pick_first_existing(sdf, ['core'])
            comment_col = pick_first_existing(sdf, ['comment', 'comments', 'notes'])

            rows = []
            for _, row in sdf.iterrows():
                variable = coerce_text(row.get(var_col, '')) if var_col else ''
                if not variable:
                    continue
                rows.append({
                    'dataset': domain,
                    'variable': variable,
                    'label': coerce_text(row.get(label_col, '')) if label_col else '',
                    'type': coerce_text(row.get(type_col, '')) if type_col else '',
                    'length': coerce_text(row.get(length_col, '')) if length_col else '',
                    'format': coerce_text(row.get(format_col, '')) if format_col else '',
                    'codelist': coerce_text(row.get(codelist_col, '')) if codelist_col else '',
                    'origin': coerce_text(row.get(origin_col, '')) if origin_col else '',
                    'role': coerce_text(row.get(role_col, '')) if role_col else '',
                    'core': coerce_text(row.get(core_col, '')) if core_col else '',
                    'comment': coerce_text(row.get(comment_col, '')) if comment_col else '',
                })
            self.domain_rows[domain] = pd.DataFrame(rows)
            self.domain_map[domain] = sname

        if self.summary_df.empty and self.domain_rows:
            rows = []
            for ds in sorted(self.domain_rows):
                rows.append({'dataset': ds, 'label': '', 'class': '', 'structure': '', 'purpose': '', 'sheet': ds})
            self.summary_df = pd.DataFrame(rows)

    def export_all_to_folder(self, folder):
        folder = Path(folder)
        folder.mkdir(parents=True, exist_ok=True)
        self.summary_df.to_csv(folder / 'dataset_summary.csv', index=False, encoding='utf-8-sig')
        domain_folder = folder / 'domains'
        domain_folder.mkdir(exist_ok=True)
        for ds, df in self.domain_rows.items():
            df.to_csv(domain_folder / f'{ds}.csv', index=False, encoding='utf-8-sig')

        combined_rows = []
        for _, r in self.summary_df.iterrows():
            ds = r.get('dataset', '')
            label = r.get('label', '')
            ddf = self.domain_rows.get(ds, pd.DataFrame())
            for _, vr in ddf.iterrows():
                combined_rows.append({'dataset': ds, 'dataset_label': label, **vr.to_dict()})
        pd.DataFrame(combined_rows).to_csv(folder / 'all_variables_combined.csv', index=False, encoding='utf-8-sig')

        meta = {
            'source_file': str(self.file_path),
            'dataset_count': int(self.summary_df.shape[0]),
            'variable_count': int(sum(df.shape[0] for df in self.domain_rows.values())),
            'datasets': sorted(self.domain_rows.keys())
        }
        with open(folder / 'parse_summary.json', 'w', encoding='utf-8') as f:
            json.dump(meta, f, indent=2)
        return folder


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry('1420x880')
        self.minsize(1220, 780)
        self.configure(bg=THEME['bg'])
        self.parser = SpecParser()
        self.filtered_datasets = []
        self.current_dataset = None
        self._style()
        self._build_ui()

    def _style(self):
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except Exception:
            pass

        style.configure('TFrame', background=THEME['bg'])
        style.configure('TLabel', background=THEME['bg'], foreground=THEME['text'], font=('Segoe UI', 10))
        style.configure('Treeview',
                        background=THEME['entry'],
                        fieldbackground=THEME['entry'],
                        foreground=THEME['entry_text'],
                        borderwidth=0,
                        rowheight=30,
                        font=('Segoe UI', 10))
        style.configure('Treeview.Heading',
                        background='#eef3f9',
                        foreground=THEME['text'],
                        relief='flat',
                        font=('Segoe UI', 10, 'bold'))
        style.map('Treeview', background=[('selected', THEME['selection'])], foreground=[('selected', THEME['selection_text'])])
        style.map('Treeview.Heading', background=[('active', '#e6edf6')])

    def make_action_button(self, parent, text, command, bg, fg='white'):
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=bg,
            activeforeground=fg,
            relief='flat',
            bd=0,
            font=('Segoe UI', 10, 'bold'),
            padx=14,
            pady=9,
            cursor='hand2'
        )

    def info_card(self, parent, title, var, bg):
        card = RoundedCard(parent, bg_color=bg, border_color=bg, radius=22, padding=1, height=94)
        card.pack(side='left', fill='both', expand=True, padx=6, pady=0)
        inner = card.inner
        tk.Label(inner, text=title, bg=bg, fg='#ffffff', font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(4, 2))
        tk.Label(inner, textvariable=var, bg=bg, fg='#ffffff', font=('Segoe UI', 10), wraplength=215, justify='left').pack(anchor='w')
        return card

    def _build_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Top banner with gentle professional colors and curved corners
        top_wrap = tk.Frame(self, bg=THEME['bg'])
        top_wrap.grid(row=0, column=0, sticky='ew', padx=14, pady=(14, 8))
        top_wrap.grid_columnconfigure(0, weight=1)

        banner = RoundedCard(top_wrap, bg_color=THEME['header'], border_color=THEME['header'], radius=28, padding=1, height=108)
        banner.grid(row=0, column=0, sticky='ew')
        b = banner.inner
        b.grid_columnconfigure(1, weight=1)

        logo = tk.Canvas(b, width=68, height=68, bg=THEME['header'], highlightthickness=0, bd=0)
        logo.grid(row=0, column=0, rowspan=2, sticky='w', padx=(6, 14), pady=4)
        pts = rounded_polygon_points(5, 5, 63, 63, 18)
        logo.create_polygon(pts, smooth=True, fill=THEME['header_accent'], outline=THEME['header_accent'])
        logo.create_text(34, 34, text='DS', fill='white', font=('Segoe UI', 18, 'bold'))

        tk.Label(b, text=APP_TITLE, bg=THEME['header'], fg='white', font=('Segoe UI', 22, 'bold')).grid(row=0, column=1, sticky='sw')
        tk.Label(b, text='Professional spec browser with rounded panels, dataset review, variable lookup, and parsed output export', bg=THEME['header'], fg='#d7e4f1', font=('Segoe UI', 10)).grid(row=1, column=1, sticky='nw')

        btnbar = tk.Frame(b, bg=THEME['header'])
        btnbar.grid(row=0, column=2, rowspan=2, sticky='e')
        self.make_action_button(btnbar, 'Open Excel Spec', self.open_spec, THEME['success']).pack(side='left', padx=5)
        self.make_action_button(btnbar, 'Export Parsed Files', self.export_outputs, THEME['warning']).pack(side='left', padx=5)
        self.make_action_button(btnbar, 'About', self.show_about, THEME['header_accent']).pack(side='left', padx=5)

        main = tk.Frame(self, bg=THEME['bg'])
        main.grid(row=1, column=0, sticky='nsew', padx=14, pady=(0, 8))
        main.grid_columnconfigure(1, weight=1)
        main.grid_rowconfigure(1, weight=1)

        left_card = RoundedCard(main, bg_color=THEME['card'], border_color=THEME['line'], radius=26, padding=2, width=330)
        left_card.grid(row=0, column=0, rowspan=2, sticky='nsew', padx=(0, 12))
        left_card.grid_propagate(False)
        left = left_card.inner
        left.grid_rowconfigure(2, weight=1)
        left.grid_columnconfigure(0, weight=1)

        tk.Label(left, text='Datasets', bg=THEME['card'], fg=THEME['text'], font=('Segoe UI', 13, 'bold')).grid(row=0, column=0, sticky='w', padx=6, pady=(4, 8))
        self.dataset_search = tk.StringVar()
        search_entry = tk.Entry(left, textvariable=self.dataset_search, bg=THEME['entry'], fg=THEME['entry_text'], relief='flat', highlightthickness=1, highlightbackground=THEME['line'], highlightcolor=THEME['accent'], font=('Segoe UI', 10))
        search_entry.grid(row=1, column=0, sticky='ew', padx=6, pady=(0, 10), ipady=9)
        search_entry.bind('<KeyRelease>', lambda e: self.refresh_dataset_list())

        self.dataset_list = tk.Listbox(left, bg=THEME['entry'], fg=THEME['entry_text'], selectbackground=THEME['selection'], selectforeground=THEME['selection_text'], relief='flat', highlightthickness=1, highlightbackground=THEME['line'], font=('Consolas', 11))
        self.dataset_list.grid(row=2, column=0, sticky='nsew', padx=6, pady=(0, 6))
        self.dataset_list.bind('<<ListboxSelect>>', self.on_dataset_select)

        info_card = RoundedCard(main, bg_color=THEME['card'], border_color=THEME['line'], radius=26, padding=2, height=126)
        info_card.grid(row=0, column=1, sticky='nsew')
        info = info_card.inner
        info.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        self.info_vars = {
            'file': tk.StringVar(value='No file loaded'),
            'datasets': tk.StringVar(value='0'),
            'variables': tk.StringVar(value='0'),
            'selected': tk.StringVar(value='-'),
            'label': tk.StringVar(value='-'),
        }

        colors = [THEME['accent'], THEME['accent2'], THEME['accent3'], '#5f7590', '#6f7e93']
        titles = ['Source File', 'Datasets', 'Variables', 'Selected', 'Dataset Label']
        keys = ['file', 'datasets', 'variables', 'selected', 'label']
        for title, key, color in zip(titles, keys, colors):
            self.info_card(info, title, self.info_vars[key], color)

        table_card = RoundedCard(main, bg_color=THEME['card'], border_color=THEME['line'], radius=26, padding=2)
        table_card.grid(row=1, column=1, sticky='nsew', pady=(12, 0))
        table = table_card.inner
        table.grid_columnconfigure(0, weight=1)
        table.grid_rowconfigure(1, weight=1)

        top_bar = tk.Frame(table, bg=THEME['card'])
        top_bar.grid(row=0, column=0, sticky='ew', padx=6, pady=(4, 10))
        top_bar.grid_columnconfigure(1, weight=1)
        tk.Label(top_bar, text='Variables', bg=THEME['card'], fg=THEME['text'], font=('Segoe UI', 13, 'bold')).grid(row=0, column=0, sticky='w')
        self.var_search = tk.StringVar()
        var_entry = tk.Entry(top_bar, textvariable=self.var_search, bg=THEME['entry'], fg=THEME['entry_text'], relief='flat', highlightthickness=1, highlightbackground=THEME['line'], highlightcolor=THEME['accent'], font=('Segoe UI', 10))
        var_entry.grid(row=0, column=1, sticky='ew', padx=(12, 12), ipady=9)
        var_entry.bind('<KeyRelease>', lambda e: self.refresh_variable_table())
        self.make_action_button(top_bar, 'Copy Selected Row', self.copy_selected_row, THEME['header_accent']).grid(row=0, column=2, sticky='e')

        columns = ['variable', 'label', 'type', 'length', 'format', 'codelist', 'origin', 'role', 'core', 'comment']
        self.tree = ttk.Treeview(table, columns=columns, show='headings')
        headings = {
            'variable': 'Variable', 'label': 'Label', 'type': 'Type', 'length': 'Length', 'format': 'Format',
            'codelist': 'Codelist', 'origin': 'Origin', 'role': 'Role', 'core': 'Core', 'comment': 'Comment'
        }
        widths = {'variable': 140, 'label': 280, 'type': 80, 'length': 80, 'format': 100, 'codelist': 120, 'origin': 120, 'role': 120, 'core': 80, 'comment': 220}
        for c in columns:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], anchor='w', stretch=True)

        self.tree.grid(row=1, column=0, sticky='nsew', padx=6, pady=(0, 6))
        ysb = ttk.Scrollbar(table, orient='vertical', command=self.tree.yview)
        xsb = ttk.Scrollbar(table, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        ysb.grid(row=1, column=1, sticky='ns', pady=(0, 6))
        xsb.grid(row=2, column=0, sticky='ew', padx=6, pady=(0, 6))

        self.status = tk.StringVar(value='Ready')
        status_wrap = tk.Frame(self, bg=THEME['bg'])
        status_wrap.grid(row=2, column=0, sticky='ew', padx=14, pady=(0, 12))
        status_card = RoundedCard(status_wrap, bg_color=THEME['panel_alt'], border_color=THEME['line'], radius=20, padding=1, height=42)
        status_card.pack(fill='x')
        tk.Label(status_card.inner, textvariable=self.status, bg=THEME['panel_alt'], fg=THEME['muted'], anchor='w', padx=4, font=('Segoe UI', 9)).pack(fill='both', expand=True)

    def set_status(self, msg):
        self.status.set(msg)
        self.update_idletasks()

    def show_about(self):
        messagebox.showinfo(
            APP_TITLE,
            f"{APP_TITLE} {APP_VERSION}\n\nUpdated with a more professional muted color palette, rounded cards, and a top banner styled in the spirit of AnnotateCRF Studio.\n\nFunctions included:\n- Load Excel spec\n- Read domains summary sheet\n- Read dataset sheets like AE, DM, LB\n- Search datasets and variables\n- Export parsed files to CSV/JSON"
        )

    def open_spec(self):
        file_path = filedialog.askopenfilename(
            title='Select Excel Spec',
            filetypes=[('Excel files', '*.xlsx *.xlsm *.xls'), ('All files', '*.*')]
        )
        if not file_path:
            return
        self.set_status('Loading spec...')
        threading.Thread(target=self._load_spec_thread, args=(file_path,), daemon=True).start()

    def _load_spec_thread(self, file_path):
        try:
            self.parser.load_excel(file_path)
            self.after(0, self._load_spec_success)
        except Exception as exc:
            tb = traceback.format_exc()
            self.after(0, lambda: self._load_spec_fail(exc, tb))

    def _load_spec_success(self):
        self.info_vars['file'].set(os.path.basename(self.parser.file_path))
        self.info_vars['datasets'].set(str(self.parser.summary_df.shape[0]))
        total_vars = sum(df.shape[0] for df in self.parser.domain_rows.values())
        self.info_vars['variables'].set(str(total_vars))
        self.refresh_dataset_list()
        self.set_status('Spec loaded successfully.')

    def _load_spec_fail(self, exc, tb):
        self.set_status('Failed to load spec.')
        messagebox.showerror('Load Error', f'Unable to load the selected spec.\n\n{exc}\n\nTechnical details:\n{tb}')

    def refresh_dataset_list(self):
        self.dataset_list.delete(0, tk.END)
        q = self.dataset_search.get().strip().lower()
        self.filtered_datasets = []
        if self.parser.summary_df.empty:
            return
        for _, row in self.parser.summary_df.sort_values('dataset').iterrows():
            ds = coerce_text(row.get('dataset', ''))
            label = coerce_text(row.get('label', ''))
            text = f"{ds:<8}  {label}"
            if q and q not in ds.lower() and q not in label.lower():
                continue
            self.filtered_datasets.append(ds)
            self.dataset_list.insert(tk.END, text)

    def on_dataset_select(self, event=None):
        sel = self.dataset_list.curselection()
        if not sel:
            return
        ds = self.filtered_datasets[sel[0]]
        self.current_dataset = ds
        self.info_vars['selected'].set(ds)
        label = ''
        if not self.parser.summary_df.empty:
            sdf = self.parser.summary_df[self.parser.summary_df['dataset'].str.upper() == ds.upper()]
            if not sdf.empty:
                label = coerce_text(sdf.iloc[0].get('label', ''))
        self.info_vars['label'].set(label or '-')
        self.refresh_variable_table()
        self.set_status(f'{ds} selected.')

    def refresh_variable_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if not self.current_dataset:
            return
        df = self.parser.domain_rows.get(self.current_dataset, pd.DataFrame()).copy()
        q = self.var_search.get().strip().lower()
        if q and not df.empty:
            mask = pd.Series(False, index=df.index)
            for c in df.columns:
                mask = mask | df[c].astype(str).str.lower().str.contains(q, na=False)
            df = df[mask]
        for _, row in df.iterrows():
            vals = [coerce_text(row.get(c, '')) for c in self.tree['columns']]
            self.tree.insert('', tk.END, values=vals)
        self.set_status(f'Showing {df.shape[0]} variable rows.')

    def copy_selected_row(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo('Copy Row', 'Please select a variable row first.')
            return
        vals = self.tree.item(sel[0], 'values')
        header = list(self.tree['columns'])
        txt = '\n'.join(f'{h}: {v}' for h, v in zip(header, vals))
        self.clipboard_clear()
        self.clipboard_append(txt)
        self.set_status('Selected row copied to clipboard.')

    def export_outputs(self):
        if self.parser.summary_df.empty and not self.parser.domain_rows:
            messagebox.showinfo('Export', 'Load a spec first.')
            return
        folder = filedialog.askdirectory(title='Select Output Folder')
        if not folder:
            return
        try:
            out = self.parser.export_all_to_folder(folder)
            self.set_status(f'Export completed: {out}')
            messagebox.showinfo('Export Complete', f'Parsed files exported to:\n\n{out}')
        except Exception as exc:
            messagebox.showerror('Export Error', str(exc))


def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
