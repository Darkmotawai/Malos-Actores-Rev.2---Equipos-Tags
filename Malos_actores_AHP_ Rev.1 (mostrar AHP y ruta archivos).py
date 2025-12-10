# -*- coding: utf-8 -*-
"""
Created on Thu Jul 17 01:09:44 2025

@author: 3jbrunet
"""

import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime, timedelta
import json
import os

CONFIG_PATH = "ahp_config.json"


weights = {}
subcriteria_config = {}


default_config = {
    "weights": {
        "FF": 0.10,
        "COS": 0.15,
        "HH": 0.05,
        "CRI": 0.10,
        "CD": 0.35,
        "IND": 0.25
    },
    "subcriteria_config": {
        "FF": [(0, 5, 0.1), (5, 10, 0.2), (10, 20, 0.3)],
        "COS": [(0, 1000, 0.1), (1000, 5000, 0.2), (5000, 10000, 0.3)],
        "HH": [(0, 10, 0.05), (10, 50, 0.15), (50, 200, 0.25)],
        "CRI": [["A", 0.05], ["B", 0.1], ["C", 0.2], ["D", 0.3], ["E", 0.4]],
        "CD": [(0, 5000, 0.1), (5000, 10000, 0.2), (10000, 20000, 0.3)],
        "IND": [(0.0, 0.01, 0.05), (0.01, 0.05, 0.10), (0.05, 0.10, 0.15), (0.10, 0.15, 0.30), (0.15, 1.0, 0.40)]
    }
}

def load_config():
    global weights, subcriteria_config
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r") as f:
            config = json.load(f)
            weights = config["weights"]
            subcriteria_config = {}

            for k, v in config["subcriteria_config"].items():

                if k == "CRI":
                    subcriteria_config[k] = [tuple(x[:2]) for x in v]
                else:
                    subcriteria_config[k] = [tuple(x) for x in v]
    else:
        weights = default_config["weights"].copy()
        subcriteria_config = default_config["subcriteria_config"].copy()

def save_config():
    config = {
        "weights": weights,
        "subcriteria_config": {k: list(v) for k, v in subcriteria_config.items()}
    }
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=4)




def read_excel_with_logging(filename, **kwargs):
    file_path = os.path.abspath(filename)
    print(f"Extrayendo datos desde: {file_path}")
    return pd.read_excel(file_path, **kwargs)


sap_df = read_excel_with_logging("Datos.xlsx", sheet_name="SAP")
solomon_df = read_excel_with_logging("Datos.xlsx", sheet_name="SOLOMON")


sap_df['Fecha de inicio extrema'] = pd.to_datetime(sap_df['Fecha de inicio extrema'], errors='coerce')
solomon_df['Fecha caida'] = pd.to_datetime(solomon_df['Fecha caida'], errors='coerce')


sap_df.rename(columns={
    'TAG': 'TAG_SAP',
    'Equipo': 'EQUIPO',
    'Costes tot.reales': 'Coste',
    'Trabajo real': 'Horas',
    'Indicador ABC': 'ABC',
    'Fecha de inicio extrema': 'Fecha'
}, inplace=True)

solomon_df.rename(columns={
    'TAG': 'TAG_SOLOMON',
    'Equipo': 'EQUIPO',
    'Costos': 'Costos',
    'Dias det': 'Dias',
    'Fecha caida': 'Fecha'
}, inplace=True)


def calculate_criteria(selected_year, selected_month):
  
    end_date = datetime(selected_year, selected_month, 1)
    start_date = end_date - timedelta(days=365)
 
    sap_filtered = sap_df[(sap_df['Fecha'] >= start_date) & (sap_df['Fecha'] < end_date)]
    solomon_filtered = solomon_df[(solomon_df['Fecha'] >= start_date) & (solomon_df['Fecha'] < end_date)]

    all_equipment = pd.Series(pd.concat([sap_filtered['EQUIPO'], solomon_filtered['EQUIPO']]).dropna().unique(), name='EQUIPO')
    equipment_df = pd.DataFrame(all_equipment)

    ff = sap_filtered.groupby('EQUIPO').size().rename("FF")
    cos = sap_filtered.groupby('EQUIPO')['Coste'].sum().rename("COS")
    hh = sap_filtered.groupby('EQUIPO')['Horas'].mean().rename("HH")
    cri = sap_filtered.sort_values('Fecha').groupby('EQUIPO')['ABC'].last().rename("CRI")
    cd = solomon_filtered.groupby('EQUIPO')['Costos'].sum().rename("CD")
    dd = solomon_filtered.groupby('EQUIPO')['Dias'].sum().rename("DD")

    merged = equipment_df.set_index('EQUIPO').join([ff, cos, hh, cri, cd, dd]).fillna(0).reset_index()
    
    kedc_df = read_excel_with_logging("kEDC.xlsx")
    
    solomon_filtered = solomon_filtered.copy()
    solomon_filtered['Año'] = solomon_filtered['Fecha'].dt.year
    solomon_filtered['Mes'] = solomon_filtered['Fecha'].dt.month
    solomon_filtered['Diasmes'] = solomon_filtered['Fecha'].dt.days_in_month
    
    solomon_enriched = pd.merge(
        solomon_filtered,
        kedc_df,
        left_on=['Planta', 'Año'],
        right_on=['Plantas', 'Año'],
        how='left'
    )
    
    solomon_enriched['Indisp_ref'] = solomon_enriched['kEDC'] * (solomon_enriched['Dias'] / solomon_enriched['Diasmes'])
    
    total_kedc_per_year = kedc_df.groupby('Año')['kEDC'].sum().rename("Total_kEDC")
    solomon_enriched = solomon_enriched.merge(total_kedc_per_year, on='Año', how='left')
    solomon_enriched['Indisp_pct'] = solomon_enriched['Indisp_ref'] / solomon_enriched['Total_kEDC']
    indisp_pct_sum = solomon_enriched.groupby('EQUIPO')['Indisp_pct'].sum().rename('IND')
    event_count = solomon_filtered.groupby('EQUIPO').size().rename('Eventos_SOL')

    merged = merged.set_index('EQUIPO').join([indisp_pct_sum, event_count]).fillna(0).reset_index()

    return merged

def apply_ahp_scoring(df):
    def get_score(criterion, value):
        if criterion == "CRI":
            for cat, weight in subcriteria_config["CRI"]:
                if str(value).upper() == cat:
                    return weight
            return 0.0
        else:
            for lower, upper, weight in subcriteria_config.get(criterion, []):
                if lower == 0 and value == 0:
                    return weight
                if lower <= value < upper:
                    return weight
            return 0.0

    total_scores = []
    contribution_tracker = {f"{crit}_TOTAL": [] for crit in weights}
    for _, row in df.iterrows():
        score = 0.0
        for crit in weights:
            raw_value = row.get(crit, 0)
            sub_score = get_score(crit, raw_value)
            contribution = weights[crit] * sub_score
            contribution_tracker[f"{crit}_TOTAL"].append(contribution)
            score += contribution
        total_scores.append(score)

    for column_name, values in contribution_tracker.items():
        df[column_name] = values
    df['AHP_SCORE'] = total_scores
    return df.sort_values(by='AHP_SCORE', ascending=False).reset_index(drop=True)

def edit_subcriteria_popup(criterion):
    popup = tk.Toplevel()
    popup.title(f"Subcriterio {criterion}")
    popup.geometry("300x400")

    container = tk.Frame(popup)
    container.pack(padx=10, pady=10, fill='both', expand=True)

    canvas = tk.Canvas(container)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    entry_widgets = []

    is_cri = criterion == "CRI"

    if is_cri:
        tk.Label(scrollable_frame, text="Categoría").grid(row=0, column=0, padx=5)
        tk.Label(scrollable_frame, text="Peso").grid(row=0, column=1, padx=5)
    else:
        tk.Label(scrollable_frame, text="Desde").grid(row=0, column=0, padx=5)
        tk.Label(scrollable_frame, text="Hasta").grid(row=0, column=1, padx=5)
        tk.Label(scrollable_frame, text="Peso").grid(row=0, column=2, padx=5)

    def add_row(*args):
        row = len(entry_widgets) + 1

        if is_cri:
            cat_val, weight = args if args else ("", 0.0)
            cat_entry = tk.Entry(scrollable_frame, width=8)
            cat_entry.insert(0, str(cat_val))
            cat_entry.grid(row=row, column=0, padx=5, pady=2)

            weight_entry = tk.Entry(scrollable_frame, width=8)
            weight_entry.insert(0, str(weight))
            weight_entry.grid(row=row, column=1, padx=5, pady=2)

            def remove_row():
                cat_entry.destroy()
                weight_entry.destroy()
                delete_btn.destroy()
                entry_widgets.remove((cat_entry, weight_entry))

            delete_btn = tk.Button(scrollable_frame, text="❌", command=remove_row)
            delete_btn.grid(row=row, column=2, padx=5)
            entry_widgets.append((cat_entry, weight_entry))
        else:
            min_val, max_val, weight = args if args else (0.0, 0.0, 0.0)
            min_entry = tk.Entry(scrollable_frame, width=8)
            min_entry.insert(0, str(min_val))
            min_entry.grid(row=row, column=0, padx=5, pady=2)

            max_entry = tk.Entry(scrollable_frame, width=8)
            max_entry.insert(0, str(max_val))
            max_entry.grid(row=row, column=1, padx=5, pady=2)

            weight_entry = tk.Entry(scrollable_frame, width=8)
            weight_entry.insert(0, str(weight))
            weight_entry.grid(row=row, column=2, padx=5, pady=2)

            def remove_row():
                min_entry.destroy()
                max_entry.destroy()
                weight_entry.destroy()
                delete_btn.destroy()
                entry_widgets.remove((min_entry, max_entry, weight_entry))

            delete_btn = tk.Button(scrollable_frame, text="❌", command=remove_row)
            delete_btn.grid(row=row, column=3, padx=5)
            entry_widgets.append((min_entry, max_entry, weight_entry))


    existing = subcriteria_config.get(criterion, [])
    for entry in existing:
        add_row(*entry)

    tk.Button(popup, text="Agregar", command=add_row).pack(pady=5)

    def save_ranges():
        new_config = []
        try:
            if is_cri:
                for cat_entry, weight_entry in entry_widgets:
                    category = cat_entry.get().strip().upper()
                    weight = float(weight_entry.get())
                    if category:
                        new_config.append([category, weight])
            else:
                for min_entry, max_entry, weight_entry in entry_widgets:
                    min_val = float(min_entry.get())
                    max_val = float(max_entry.get())
                    weight = float(weight_entry.get())
                    new_config.append((min_val, max_val, weight))
        except ValueError:
            messagebox.showerror("Input Error", "Check your inputs.")
            return

        subcriteria_config[criterion] = sorted(new_config, key=lambda x: x[0] if not is_cri else x[0])
        save_config()
        popup.destroy()

    tk.Button(popup, text="Guardar", command=save_ranges).pack(pady=10)

COLUMN_LABELS = {
    "EQUIPO": "ACTIVO",
    "FF": "FREC. DE FALLA",
    "COS": "COS. DE MTTO",
    "HH": "MTTR",
    "CRI": "CRIT. ABC",
    "IND": "IND.",
    "Eventos_SOL": "N° EVENTOS",
    "Puntaje AHP": "PUNTAJE AHP",
    "Ranking": "Ranking"
}

def show_table(dataframe, selected_year, selected_month, weight_snapshot):
    dataframe = dataframe.copy()

    dataframe.rename(columns={'AHP_SCORE': 'Puntaje AHP'}, inplace=True)

    dataframe.drop(columns=['DD', 'INDISP_REF', 'INDISP_PORCENTAJE'], inplace=True, errors='ignore')


    cols = list(dataframe.columns)
    contribution_cols = [f"{crit}_TOTAL" for crit in weights if f"{crit}_TOTAL" in cols]
    for contrib_col in contribution_cols:
        cols.remove(contrib_col)
    if 'Puntaje AHP' in cols:
        insert_index = cols.index('Puntaje AHP')
        for offset, contrib_col in enumerate(contribution_cols):
            cols.insert(insert_index + offset, contrib_col)

    if 'IND' in cols and 'Puntaje AHP' in cols:
        cols.remove('IND')
        cols.insert(cols.index('Puntaje AHP'), 'IND')

    if 'Eventos_SOL' in cols:
        cols.remove('Eventos_SOL')
        cols.insert(cols.index('IND') + 1, 'Eventos_SOL')
    dataframe = dataframe[cols]

 
    criteria_order = [k for k, _ in sorted(weights.items(), key=lambda x: -x[1])]
    sort_columns = ['Puntaje AHP'] + criteria_order
    dataframe['Ranking'] = dataframe.sort_values(
        by=sort_columns,
        ascending=[False] * len(sort_columns)
    ).reset_index().index + 1


    table_window = tk.Toplevel()
    table_window.title("Listado AHP")

    style = ttk.Style(table_window)
    style.theme_use("default")
    style.configure("Treeview.Heading", background="black", foreground="white", font=("Arial", 10, "bold"))
    style.configure("Treeview", font=("Arial", 10), rowheight=25)
    style.map("Treeview", background=[("selected", "#ececec")])
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

    frame = ttk.Frame(table_window, borderwidth=1, relief="solid")
    frame.pack(fill='both', expand=True, padx=10, pady=10)

    tree = ttk.Treeview(frame, columns=cols, show='headings')
    tree.pack(side='left', fill='both', expand=True)

    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side='right', fill='y')

    for col in cols:
        tree.heading(col, text=col, command=lambda _col=col: sort_column(tree, _col, False))
        tree.column(col, anchor='center', width=120)

    for _, row in dataframe.iterrows():
        formatted = []
        for col in cols:
            val = row[col]
            if col in ['FF', 'COS', 'HH']:
                formatted.append(f"{int(val)}")
            elif col == 'CD':
                formatted.append(f"{val:,.1f}")
            elif col == 'IND':
                formatted.append(f"{val * 100:.1f}%")
            elif col == 'Eventos_SOL':
                formatted.append(f"{int(val)}")
            elif col.endswith('_TOTAL'):
                formatted.append(f"{val:.3f}")
            elif col == 'Puntaje AHP':
                formatted.append(f"{val:.3f}")
            elif col == 'Ranking':
                formatted.append(f"{val}°")
            else:
                formatted.append(str(val))
        tree.insert("", "end", values=formatted)

 
    def back_to_main():
        table_window.destroy()
        open_selection_window(
            prev_year=selected_year,
            prev_month=selected_month,
            prev_weights=weight_snapshot
        )

    tk.Button(table_window, text="← Back to Main", bg="black", fg="white",
              font=("Arial", 10, "bold"), command=back_to_main).pack(pady=10)



def sort_column(treeview, col, reverse):
    data = [(treeview.set(k, col), k) for k in treeview.get_children('')]
    try:
        data.sort(key=lambda t: float(t[0]), reverse=reverse)
    except ValueError:
        data.sort(reverse=reverse)
    for index, (_, k) in enumerate(data):
        treeview.move(k, '', index)
    treeview.heading(col, command=lambda: sort_column(treeview, col, not reverse))


def open_selection_window(prev_year=None, prev_month=None, prev_weights=None):
    root = tk.Toplevel()
    root.title("MALOS ACTORES AHP")

    default_weights = weights.copy()
    if prev_weights:
        default_weights.update(prev_weights)
    weight_vars = {crit: tk.DoubleVar(value=default_weights.get(crit, 0.1)) for crit in ['FF', 'COS', 'HH', 'CRI', 'CD', 'IND']}

    tk.Label(root, text="Año").grid(row=0, column=1)
    year_var = tk.StringVar(value=str(prev_year or "2025"))
    tk.Entry(root, textvariable=year_var, width=6).grid(row=1, column=1)

    tk.Label(root, text="Mes").grid(row=0, column=2)
    month_var = tk.StringVar(value=str(prev_month or "6"))
    tk.Entry(root, textvariable=month_var, width=6).grid(row=1, column=2)

 
    criteria = ['FF', 'COS', 'HH', 'CRI', 'CD', 'IND']
    for i, crit in enumerate(criteria):
        tk.Label(root, text=crit).grid(row=2, column=i)
        tk.Entry(root, textvariable=weight_vars[crit], width=6).grid(row=3, column=i)
        tk.Button(root, text="Editar\nsubcriterio", width=12,
          command=lambda c=crit: edit_subcriteria_popup(c)).grid(row=4, column=i)
 
    def on_submit():
        try:
            selected_month = int(month_var.get())
            selected_year = int(year_var.get())
            if selected_month < 1 or selected_month > 12:
                raise ValueError
        except ValueError:
            messagebox.showerror("Input Error", "Enter valid month (1-12) and year (e.g., 2025).")
            return

        for crit in criteria:
            weights[crit] = float(weight_vars[crit].get())
        save_config()

        root.destroy()
        result_df = calculate_criteria(selected_year, selected_month)
        result_df = apply_ahp_scoring(result_df)
        show_table(result_df, selected_year, selected_month, {crit: float(weight_vars[crit].get()) for crit in criteria})
        
    def save_weights_only():
        for crit in criteria:
            weights[crit] = float(weight_vars[crit].get())
        save_config()
        messagebox.showinfo("Guardado", "Pesos guardados correctamente.")
    
    tk.Button(root, text="GUARDAR PESOS", width=20, height=2, command=save_weights_only).grid(row=5, column=0, columnspan=6, pady=5)
    tk.Button(root, text="VER LISTADO", width=20, height=2, command=on_submit).grid(row=6, column=0, columnspan=6, pady=10)
    root.mainloop()

load_config()
open_selection_window()



