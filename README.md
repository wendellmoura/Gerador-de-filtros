import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
import os
import json
import numpy as np

class ExcelToHTMLConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de filtros HTML")
        self.root.geometry("1025x700")

        self.file_path = None
        self.df = None
        self.pivot_rows = []
        self.pivot_columns = []
        self.pivot_values = None
        self.pivot_agg = 'sum'

        self.create_widgets()

    def create_widgets(self):
        # Notebook para abas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Frame principal para a aba de exportação
        main_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(main_frame, text="Exportação Principal")

        # Frame para a aba dinâmica
        pivot_frame = ttk.Frame(self.notebook)
        self.notebook.add(pivot_frame, text="Tabela Dinâmica")

        # Conteúdo da aba principal
        title_label = ttk.Label(
            main_frame,
            text="Exportador de filtros",
            font=("Helvetica", 14, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        file_frame = ttk.LabelFrame(main_frame, text="Selecionar Arquivo", padding=10)
        file_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 20))

        self.file_entry = ttk.Entry(file_frame, width=50)
        self.file_entry.grid(row=0, column=0, padx=(0, 10))

        browse_btn = ttk.Button(
            file_frame,
            text="Procurar",
            command=self.browse_file
        )
        browse_btn.grid(row=0, column=1)

        options_frame = ttk.LabelFrame(main_frame, text="Opções de Exportação", padding=10)
        options_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 20))

        ttk.Label(options_frame, text="Nome do arquivo HTML:").grid(row=0, column=0, sticky="w")
        self.output_name = ttk.Entry(options_frame, width=30)
        self.output_name.grid(row=0, column=1, sticky="w", padx=(0, 10))
        self.output_name.insert(0, "tabela_exportada.html")

        ttk.Label(options_frame, text="Título da página:").grid(row=1, column=0, sticky="w")
        self.page_title = ttk.Entry(options_frame, width=30)
        self.page_title.grid(row=1, column=1, sticky="w", padx=(0, 10))
        self.page_title.insert(0, "Dados Exportados")

        self.alternate_rows = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Linhas com cores alternadas",
            variable=self.alternate_rows
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(5, 0))

        self.horizontal_scroll = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Adicionar rolagem horizontal quando necessário",
            variable=self.horizontal_scroll
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(5, 0))

        self.highlight_search = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Realçar termos pesquisados",
            variable=self.highlight_search
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(5, 0))

        export_btn = ttk.Button(
            main_frame,
            text="Exportar HTML",
            command=self.export_to_html
        )
        export_btn.grid(row=3, column=0, columnspan=2, pady=(10, 0))

        preview_frame = ttk.LabelFrame(main_frame, text="Pré-visualização", padding=10)
        preview_frame.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(20, 0))

        main_frame.rowconfigure(4, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        self.tree = ttk.Treeview(preview_frame)
        self.tree.pack(fill=tk.BOTH, expand=True)

        scroll_y = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scroll_y.set)

        scroll_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=scroll_x.set)

        # Conteúdo da aba dinâmica
        pivot_config_frame = ttk.LabelFrame(pivot_frame, text="Configuração da Tabela Dinâmica", padding=10)
        pivot_config_frame.pack(fill=tk.X, padx=10, pady=5)

        # Configuração de linhas
        ttk.Label(pivot_config_frame, text="Linhas:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.row_combo = ttk.Combobox(pivot_config_frame, state="readonly")
        self.row_combo.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        
        # Configuração de colunas (agora opcional)
        ttk.Label(pivot_config_frame, text="Colunas:").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.col_combo = ttk.Combobox(pivot_config_frame, state="readonly")
        self.col_combo.grid(row=0, column=3, sticky="ew", padx=(0, 10))
        
        # Configuração de valores
        ttk.Label(pivot_config_frame, text="Valores:").grid(row=0, column=4, sticky="w", padx=(0, 5))
        self.val_combo = ttk.Combobox(pivot_config_frame, state="readonly")
        self.val_combo.grid(row=0, column=5, sticky="ew", padx=(0, 10))
        
        # Configuração de agregação
        ttk.Label(pivot_config_frame, text="Agregação:").grid(row=0, column=6, sticky="w", padx=(0, 5))
        self.agg_combo = ttk.Combobox(pivot_config_frame, 
                                    values=['soma', 'média', 'contagem', 'máximo', 'mínimo'], 
                                    state="readonly")
        self.agg_combo.grid(row=0, column=7, sticky="ew")
        self.agg_combo.current(0)

        # Botão para gerar pivot
        gen_btn = ttk.Button(
            pivot_config_frame,
            text="Gerar Tabela Dinâmica",
            command=self.generate_pivot
        )
        gen_btn.grid(row=0, column=8, padx=(10, 0))

        # Área de visualização da pivot
        pivot_preview_frame = ttk.LabelFrame(pivot_frame, text="Pré-visualização da Tabela Dinâmica", padding=10)
        pivot_preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.pivot_tree = ttk.Treeview(pivot_preview_frame)
        self.pivot_tree.pack(fill=tk.BOTH, expand=True)

        pivot_scroll_y = ttk.Scrollbar(pivot_preview_frame, orient="vertical", command=self.pivot_tree.yview)
        pivot_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.pivot_tree.configure(yscrollcommand=pivot_scroll_y.set)

        pivot_scroll_x = ttk.Scrollbar(pivot_preview_frame, orient="horizontal", command=self.pivot_tree.xview)
        pivot_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.pivot_tree.configure(xscrollcommand=pivot_scroll_x.set)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

            try:
                if file_path.endswith('.csv'):
                    self.df = pd.read_csv(file_path)
                else:
                    self.df = pd.read_excel(file_path)
                if self.df.empty or self.df.columns.empty:
                    raise Exception("Arquivo está vazio ou não possui colunas.")
                self.df = self.df.dropna(axis=1, how='all')
                self.show_preview()
                
                # Atualizar comboboxes da pivot com opção vazia para colunas
                cols = list(self.df.columns)
                self.row_combo['values'] = cols
                self.col_combo['values'] = [''] + cols  # Adiciona opção vazia
                self.val_combo['values'] = cols
                if cols:
                    self.row_combo.current(0)
                    self.col_combo.set('')  # Começa vazio
                    self.val_combo.current(0)
                    
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{str(e)}")

    def show_preview(self):
        if self.df is not None:
            self.df = self.df.dropna(axis=1, how='all')
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = list(self.df.columns)
        self.tree["show"] = "headings"
        for col in self.df.columns:
            self.tree.heading(col, text=col)
        for _, row in self.df.head(50).iterrows():
            self.tree.insert("", tk.END, values=list(row))

    def generate_pivot(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Por favor, selecione um arquivo primeiro.")
            return
            
        row = self.row_combo.get()
        col = self.col_combo.get()
        val = self.val_combo.get()
        agg = self.agg_combo.get()
        
        # Validação ajustada (colunas agora opcionais)
        if not row or not val:
            messagebox.showwarning("Aviso", "Selecione pelo menos uma coluna para Linhas e Valores.")
            return
            
        try:
            # Mapear função de agregação
            agg_map = {
                'soma': 'sum',
                'média': 'mean',
                'contagem': 'count',
                'máximo': 'max',
                'mínimo': 'min'
            }
            agg_func = agg_map.get(agg, 'sum')
            
            # Gerar tabela dinâmica (colunas vazias permitidas)
            pivot = pd.pivot_table(self.df, 
                                  index=[row], 
                                  columns=[col] if col else None, 
                                  values=[val], 
                                  aggfunc=agg_func, 
                                  fill_value=0)
            
            # Resetar índice para exibição
            pivot_df = pivot.reset_index()
            
            # Salvar configuração para exportação
            self.pivot_rows = [row]
            self.pivot_columns = [col] if col else []
            self.pivot_values = val
            self.pivot_agg = agg
            
            # Atualizar visualização
            for item in self.pivot_tree.get_children():
                self.pivot_tree.delete(item)
                
            # Configurar colunas
            cols = list(pivot_df.columns)
            self.pivot_tree["columns"] = cols
            self.pivot_tree["show"] = "headings"
            for col_name in cols:
                self.pivot_tree.heading(col_name, text=str(col_name))
                
            # Adicionar dados
            for _, row in pivot_df.iterrows():
                self.pivot_tree.insert("", tk.END, values=list(row))
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar tabela dinâmica:\n{str(e)}")

    def export_to_html(self):
        if self.df is None:
            messagebox.showwarning("Aviso", "Por favor, selecione um arquivo primeiro.")
            return

        output_file = self.output_name.get()
        if not output_file:
            messagebox.showwarning("Aviso", "Por favor, especifique um nome para o arquivo de saída.")
            return

        try:
            df_export = self.df.dropna(axis=1, how='all')
            html_content = self.generate_html(df_export)
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(html_content)
            messagebox.showinfo(
                "Sucesso",
                f"Arquivo exportado com sucesso para:\n{os.path.abspath(output_file)}"
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao exportar:\n{str(e)}")

    def generate_html(self, df=None):
        if df is None:
            df = self.df
        page_title = self.page_title.get()
        alternate_rows = self.alternate_rows.get()
        horizontal_scroll = self.horizontal_scroll.get()
        columns = list(df.columns)

        def to_display(val):
            if isinstance(val, float) and val.is_integer():
                return str(int(val))
            if pd.isnull(val):
                return ""
            return str(val)

        filters = {
            col: sorted(
                [to_display(x) for x in df[col].dropna().unique()],
                key=lambda v: v
            ) for col in columns
        }

        all_duplicates = {}
        for col in columns:
            counts = df[col].value_counts()
            dups = counts[counts > 1]
            if not dups.empty:
                list_dups = []
                for val, count in dups.items():
                    val_disp = to_display(val)
                    list_dups.append((val_disp, int(count)))
                all_duplicates[str(col)] = list_dups
        all_duplicates_json = json.dumps(all_duplicates, ensure_ascii=False)

        filters_html = '''
        <div class="filters-row align-items-end" style="margin-bottom:10px;">
        {}
        </div>
        '''.format(''.join([
            '''
            <div class="dropdown-checkbox" data-col="{i}">
                <span class="dropdown-toggle btn btn-light btn-sm border" onclick="toggleDropdown({i})">{col}</span>
                <div class="dropdown-menu" id="menu_{i}">
                    <input type="text" placeholder="Buscar..." class="form-control form-control-sm mb-2" onkeyup="filterCheckboxes({i}, this.value)" onmousedown="event.stopPropagation()">
                    <div class="dropdown-scroll" style="max-height:220px;overflow-y:auto;padding-right:3px;">
                        <label>
                            <input type="checkbox" class="filter-checkbox" data-col="{i}" value="__ALL__" checked onchange="allCheck({i})"> <b>Todos</b>
                        </label>
                        {checks}
                    </div>
                </div>
            </div>
            '''.format(
                i=i,
                col=col,
                checks=''.join([
                    '<label><input type="checkbox" class="filter-checkbox" data-col="{i}" value="{val}"> {val}</label>'.format(i=i, val=val)
                    for val in filters[col]
                ])
            )
            for i, col in enumerate(columns)
        ]))

        thead = ''.join('<th class="sortable" data-col="{}">{}</th>'.format(i, col) for i, col in enumerate(columns))
        tbody = ""

        data_json = df.applymap(to_display).fillna("").to_dict(orient='records')

        # Gerar HTML para duplicados
        duplicates_html = f"""
            <div class="mb-2">
                <label>Coluna:&nbsp;
                    <select id='dupColSelect' class='form-select form-select-sm' style='display:inline-block; width:auto;'>
                        <option value='_ALL_'>Todas</option>
                        {''.join(f"<option value='{col}'>{col}</option>" for col in columns)}
                    </select>
                </label>
            </div>
            <div id="duplicatesTableDiv"></div>
        """

        # Gerar HTML para tabela dinâmica se configurada
        pivot_html = ""
        pivot_js = ""
        pivot_config_js = ""
        if self.pivot_rows and self.pivot_values:
            try:
                # Mapear função de agregação
                agg_map = {
                    'soma': 'sum',
                    'média': 'mean',
                    'contagem': 'count',
                    'máximo': 'max',
                    'mínimo': 'min'
                }
                agg_func = agg_map.get(self.pivot_agg, 'sum')
                
                # Gerar pivot table
                pivot = pd.pivot_table(df, 
                                      index=self.pivot_rows, 
                                      columns=self.pivot_columns if self.pivot_columns else None, 
                                      values=self.pivot_values, 
                                      aggfunc=agg_func, 
                                      fill_value=0)
                
                # Converter para formato JSON
                pivot_data = pivot.reset_index()
                pivot_json = pivot_data.to_dict(orient='records')
                pivot_cols = list(pivot_data.columns)
                
                # Gerar HTML da tabela com atributos de dados para clique
                pivot_html = "<table class='table table-bordered table-sm' id='pivotTable'><thead><tr>"
                for col in pivot_cols:
                    pivot_html += f"<th>{col}</th>"
                pivot_html += "</tr></thead><tbody>"
                
                for record in pivot_json:
                    pivot_html += "<tr>"
                    for col_idx, col in enumerate(pivot_cols):
                        value = record[col]
                        # Adicionar atributos de dados para linha e coluna
                        row_value = str(record[self.pivot_rows[0]]) if self.pivot_rows else ''
                        col_value = str(pivot_cols[col_idx]) if col_idx > 0 and self.pivot_columns else ''
                        
                        pivot_html += f"<td data-row='{row_value}' data-col='{col_value}'>{value}</td>"
                    pivot_html += "</tr>"
                pivot_html += "</tbody></table>"
                
                # Configuração para JavaScript
                pivot_config_js = f"""
                    const pivotConfig = {{
                        rows: {json.dumps(self.pivot_rows)},
                        columns: {json.dumps(self.pivot_columns)},
                        values: {json.dumps(self.pivot_values)},
                        aggFunc: '{self.pivot_agg}'
                    }};
                    const pivotData = {json.dumps(pivot_json)};
                    const pivotColumns = {json.dumps(pivot_cols)};
                """
                
                # Função para criar nova aba com detalhes (CORRIGIDA)
                pivot_js = """
                function openDetailTab(rowValue, colValue) {
                    // Criar título único para a aba
                    const tabId = `detail_${Date.now()}`;
                    const tabTitle = `Detalhes: ${rowValue || ''}${colValue ? ' - ' + colValue : ''}`;
                    
                    // Adicionar aba
                    const navTabs = document.querySelector('.nav-tabs');
                    const tabContent = document.querySelector('.tab-content');
                    
                    const tabLink = document.createElement('li');
                    tabLink.className = 'nav-item';
                    tabLink.innerHTML = `
                        <a class="nav-link" id="${tabId}-tab" data-bs-toggle="tab" href="#${tabId}" role="tab">
                            ${tabTitle}
                            <button type="button" class="btn-close btn-sm" aria-label="Fechar" onclick="closeTab('${tabId}')"></button>
                        </a>
                    `;
                    navTabs.appendChild(tabLink);
                    
                    const tabPane = document.createElement('div');
                    tabPane.className = 'tab-pane fade';
                    tabPane.id = tabId;
                    tabPane.role = 'tabpanel';
                    
                    // Filtrar dados originais CORRETAMENTE
                    const filteredData = allData.filter(item => {
                        const rowMatch = pivotConfig.rows.length > 0 ? 
                            pivotConfig.rows.every(rowField => 
                                item[rowField] == rowValue) : true;
                            
                        const colMatch = pivotConfig.columns.length > 0 ? 
                            pivotConfig.columns.every(colField => 
                                item[colField] == colValue) : true;
                                
                        return rowMatch && colMatch;
                    });
                    
                    // Gerar tabela de detalhes
                    let detailTable = `<table class="table table-bordered table-striped"><thead><tr>`;
                    columns.forEach(col => detailTable += `<th>${col}</th>`);
                    detailTable += `</tr></thead><tbody>`;
                    
                    filteredData.forEach(row => {
                        detailTable += `<tr>`;
                        columns.forEach(col => {
                            detailTable += `<td>${row[col] || ''}</td>`;
                        });
                        detailTable += `</tr>`;
                    });
                    
                    detailTable += `</tbody></table>`;
                    
                    tabPane.innerHTML = `
                        <div class="p-3">
                            <h5>${tabTitle}</h5>
                            <div>Total de registros: ${filteredData.length}</div>
                            ${detailTable}
                        </div>
                    `;
                    
                    tabContent.appendChild(tabPane);
                    
                    // Ativar a nova aba
                    new bootstrap.Tab(tabLink.querySelector('.nav-link')).show();
                }
                
                function closeTab(tabId) {
                    const tabLink = document.querySelector(`#${tabId}-tab`).parentNode;
                    const tabPane = document.querySelector(`#${tabId}`);
                    
                    tabLink.remove();
                    tabPane.remove();
                    
                    // Ativar primeira aba se necessário
                    const firstTab = document.querySelector('.nav-tabs .nav-link');
                    if (firstTab) {
                        new bootstrap.Tab(firstTab).show();
                    }
                }
                
                function setupPivotTable() {
                    const table = document.getElementById('pivotTable');
                    if (!table) return;
                    
                    table.addEventListener('click', function(e) {
                        const cell = e.target.closest('td');
                        if (!cell) return;
                        
                        // Obter valores dos atributos de dados
                        const rowValue = cell.getAttribute('data-row');
                        const colValue = cell.getAttribute('data-col');
                        
                        if (rowValue || colValue) {
                            openDetailTab(rowValue, colValue);
                        }
                    });
                }
                """

            except Exception as e:
                pivot_html = f"<div class='alert alert-danger'>Erro ao gerar tabela dinâmica: {str(e)}</div>"

        html = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{page_title}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {{ padding: 20px; background-color: #f8f9fa; }}
        .filter-section {{ margin-bottom: 15px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; }}
        .filters-row {{
            display: flex;
            flex-wrap: nowrap;
            overflow-x: auto;
            gap: 8px;
            padding-bottom: 6px;
        }}
        .dropdown-checkbox {{
            display: inline-block;
            position: relative;
            margin-right: 0;
            margin-bottom: 6px;
            min-width: 120px;
            max-width: 220px;
            vertical-align: top;
        }}
        .dropdown-checkbox .dropdown-toggle {{
            cursor: pointer;
            width: 100%;
            text-overflow: ellipsis;
            overflow: hidden;
            white-space: nowrap;
        }}
        .dropdown-checkbox .dropdown-menu {{
            display: none;
            position: fixed;
            min-width: 220px;
            background: #fff;
            border: 1px solid #ccc;
            box-shadow: 0 4px 8px 0 rgba(0,0,0,.15);
            z-index: 9999;
            padding: 8px 8px 8px 8px;
        }}
        .dropdown-checkbox.show .dropdown-menu {{
            display: block;
        }}
        .dropdown-checkbox label {{ display: block; font-weight: normal; margin-bottom: 0; }}
        .dropdown-checkbox .dropdown-toggle::after {{
            content: "▼";
            margin-left: 5px;
            font-size: 5px;
        }}
        .dropdown-scroll {{
            max-height: 220px;
            overflow-y: auto;
            scrollbar-width: thin;
        }}
        .dropdown-scroll::-webkit-scrollbar {{
            width: 6px;
            background: #f1f1f1;
        }}
        .dropdown-scroll::-webkit-scrollbar-thumb {{
            background: #cccccc;
            border-radius: 3px;
        }}
        .highlight {{ background-color: #fff3cd !important; font-weight: bold; }}
        .table-striped tbody tr:nth-of-type(odd) {{ background-color: rgba(0, 0, 0, 0.02); }}
        .badge-filter {{ margin-right: 5px; margin-bottom: 5px; }}
        .pagination-controls {{
            margin-top: 8px;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .tab-pane-content {{ padding: 12px 0; }}
        .sortable {{ cursor: pointer; user-select: none; }}
        .sortable:after {{ content: " ⇅"; font-size: 0.8em; color: #888; }}
        .sortable.sorted-asc:after {{ content: " ↑"; }}
        .sortable.sorted-desc:after {{ content: " ↓"; }}
        .footer-credit {{
            position:fixed;
            left:10px;
            bottom:5px;
            color:#888;
            font-size:11px;
            opacity:0.15;
            z-index:9999;
            pointer-events:none;
            user-select:none;
            letter-spacing:1px;
        }}
        .clear-filters-btn {{
            background: none;
            border: none;
            color: #888;
            font-size: .92em;
            margin-left: 8px;
            padding: 3px 10px;
            border-radius: 4px;
            transition: background 0.2s, color 0.2s;
        }}
        .clear-filters-btn:hover {{
            background: #e6e6e6;
            color: #444;
        }}
        .nav-tabs .btn-close {{
            font-size: 0.6rem;
            padding: 0.15rem 0.25rem;
            margin-left: 8px;
            line-height: 0.8;
        }}
        @media (max-width: 600px) {{
            .dropdown-checkbox .dropdown-menu {{
                min-width: 170px;
                max-width: 99vw;
            }}
        }}
    </style>
</head>
<body>
    <div class="container-fluid">
        <h2>{page_title}</h2>
        <ul class="nav nav-tabs" id="mainTab" role="tablist">
            <li class="nav-item" role="presentation">
                <button class="nav-link active" id="table-tab" data-bs-toggle="tab" data-bs-target="#table-pane" type="button" role="tab" aria-controls="table-pane" aria-selected="true">Tabela</button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="duplicates-tab" data-bs-toggle="tab" data-bs-target="#duplicates-pane" type="button" role="tab" aria-controls="duplicates-pane" aria-selected="false">Duplicados</button>
            </li>
            {pivot_tab}
        </ul>
        <div class="tab-content">
            <div class="tab-pane fade show active tab-pane-content" id="table-pane" role="tabpanel" aria-labelledby="table-tab">
                <div class="filter-section">
                    <div class="d-flex align-items-center flex-wrap mb-2">
                        <input class="form-control me-2 mb-2" id="globalSearch" style="max-width:300px;" placeholder="Pesquisar em todas as colunas..." type="text"/>
                        <button class="clear-filters-btn mb-2" id="clearFiltersBtn" type="button" title="Limpar filtros" aria-label="Limpar filtros">
                            <svg xmlns="http://www.w3.org/2000/svg" height="18" viewBox="0 0 24 24" width="18" style="vertical-align:middle;margin-right:2px;opacity:.7;"><path d="M19 13H5v-2h14v2z" fill="#888"/></svg> Limpar filtros
                        </button>
                    </div>
                    {filters_html}
                    <div id="activeFilters"></div>
                </div>
                <div class="table-responsive{fixed_header}">
                    <table class="table table-bordered{table_striped}" id="dataTable">
                        <thead><tr>
                            {thead}
                        </tr></thead>
                        <tbody id="tableBody">
                            {tbody}
                        </tbody>
                    </table>
                    <div id="paginationControls" class="pagination-controls"></div>
                </div>
                <div class="mt-2"><small>Exportado em {export_time} | Total de registros: <span id="totalRecords">{df_len}</span></small></div>
            </div>
            <div class="tab-pane fade tab-pane-content" id="duplicates-pane" role="tabpanel" aria-labelledby="duplicates-tab">
                <h5>Valores Duplicados por Coluna</h5>
                {duplicates_html}
            </div>
            {pivot_pane}
        </div>
    </div>
    <div class="footer-credit">
        Desenvolvido por Wendell Moura
    </div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
    const allData = {all_data};
    const columns = {columns};
    let filteredData = [...allData];
    let rowsPerPage = 100;
    let currentPage = 1;
    let sortCol = null;
    let sortDir = null;
    
    {pivot_config_js}
    
    function getActiveFilterValues() {{
        let filters = {{}};
        for(let i=0;i<columns.length;i++) {{
            let menu = document.getElementById('menu_'+i);
            let checks = menu.querySelectorAll('.filter-checkbox:checked');
            let vals = Array.from(checks).map(cb => cb.value).filter(v => v !== "__ALL__");
            if(vals.length>0) filters[i] = vals;
        }}
        return filters;
    }}
    
    // CROSS-FILTER: Atualiza opções de filtro com base nos outros filtros
    function updateFilterOptions() {{
        let filters = getActiveFilterValues();
        for(let i=0;i<columns.length;i++) {{
            let filtersWithoutCurrent = Object.assign({{}}, filters);
            delete filtersWithoutCurrent[i];
            let possibleData = allData.filter(row => {{
                for (let idx in filtersWithoutCurrent) {{
                    let col = columns[idx];
                    let v = String(row[col]);
                    if (!filtersWithoutCurrent[idx].includes(v)) return false;
                }}
                return true;
            }});
            let col = columns[i];
            let possibleVals = new Set();
            possibleData.forEach(row => {{
                let v = row[col];
                if(v !== undefined) possibleVals.add(String(v));
            }});
            let menu = document.getElementById('menu_'+i);
            let checkboxes = menu.querySelectorAll('.filter-checkbox:not([value="__ALL__"])');
            checkboxes.forEach(cb => {{
                // Se já está marcado, deixa visível e habilitado SEMPRE
                if(cb.checked || possibleVals.has(cb.value)) {{
                    cb.parentElement.style.display = '';
                    cb.disabled = false;
                }} else {{
                    cb.parentElement.style.display = 'none';
                    cb.disabled = true;
                }}
            }});
            let anyChecked = Array.from(checkboxes).some(cb => cb.checked);
            let allCb = menu.querySelector('.filter-checkbox[value="__ALL__"]');
            if (!anyChecked) allCb.checked = true;
        }}
    }}
    
    function filterData() {{
        let filters = getActiveFilterValues();
        let search = document.getElementById("globalSearch").value.toLowerCase();
        if (Object.keys(filters).length === 0 && !search) return allData;
        return allData.filter(row => {{
            for (let idx in filters) {{
                let col = columns[idx];
                let v = String(row[col]);
                if (!filters[idx].includes(v)) return false;
            }}
            if (search) {{
                let found = false;
                for (let col of columns) {{
                    let val = row[col];
                    if (val !== undefined && String(val).toLowerCase().includes(search)) {{
                        found = true;
                        break;
                    }}
                }}
                if (!found) return false;
            }}
            return true;
        }});
    }}
    
    function renderTable(data) {{
        let tbody = '';
        let start = (currentPage - 1) * rowsPerPage;
        let end = Math.min(start + rowsPerPage, data.length);
        let dataToRender = data;
        if(sortCol !== null) {{
            dataToRender = [...data].sort((a, b) => {{
                let va = a[columns[sortCol]];
                let vb = b[columns[sortCol]];
                if(va === undefined || va === null) return 1;
                if(vb === undefined || vb === null) return -1;
                va = String(va).toLowerCase();
                vb = String(vb).toLowerCase();
                if(va < vb) return sortDir === 'asc' ? -1 : 1;
                if(va > vb) return sortDir === 'asc' ? 1 : -1;
                return 0;
            }});
        }}
        for(let i = start; i < end; i++) {{
            let row = dataToRender[i];
            tbody += '<tr>' + columns.map(col => {{
                let v = row[col];
                if (v === undefined || v === null) v = "";
                return `<td>${{v}}</td>`;
            }}).join('') + '</tr>';
        }}
        document.getElementById('tableBody').innerHTML = tbody;
        document.getElementById('totalRecords').innerText = data.length;
        updatePaginationControls(data.length);
    }}
    
    function updatePaginationControls(totalRows) {{
        let controls = document.getElementById('paginationControls');
        if (!controls) return;
        let totalPages = Math.max(1, Math.ceil(totalRows / rowsPerPage));
        controls.innerHTML = `
            <button class="btn btn-sm btn-light" onclick="prevPage()" ${{currentPage === 1 ? 'disabled' : ''}}>Anterior</button>
            Página ${{currentPage}} de ${{totalPages}}
            <button class="btn btn-sm btn-light" onclick="nextPage(${{totalPages}})" ${{currentPage === totalPages ? 'disabled' : ''}}>Próxima</button>
        `;
    }}
    
    window.prevPage = function() {{
        if (currentPage > 1) {{
            currentPage--;
            renderTable(filteredData);
            highlightSearch(document.getElementById("globalSearch").value.toLowerCase());
        }}
    }}
    window.nextPage = function(totalPages) {{
        if (currentPage < totalPages) {{
            currentPage++;
            renderTable(filteredData);
            highlightSearch(document.getElementById("globalSearch").value.toLowerCase());
        }}
    }}
    
    function updateActiveFilters() {{
        let filters = getActiveFilterValues();
        let html = '';
        Object.keys(filters).forEach(idx => {{
            let col = columns[idx];
            filters[idx].forEach(val => {{
                html += `<span class="badge bg-info text-dark badge-filter">${{col}}: ${{val}} <a href="#" onclick="removeFilter(${{idx}}, '${{val}}');return false;">&times;</a></span>`;
            }});
        }});
        document.getElementById('activeFilters').innerHTML = html;
    }}
    
    function removeFilter(idx, val) {{
        let menu = document.getElementById('menu_'+idx);
        let cbs = menu.querySelectorAll('.filter-checkbox');
        cbs.forEach(cb => {{
            if(cb.value === val) cb.checked = false;
        }});
        let checked = menu.querySelectorAll('.filter-checkbox:checked:not([value="__ALL__"])');
        if(checked.length === 0) {{
            menu.querySelector('.filter-checkbox[value="__ALL__"]').checked = true;
        }}
        applyFilters();
    }}
    
    function applyFilters() {{
        filteredData = filterData();
        currentPage = 1;
        renderTable(filteredData);
        updateActiveFilters();
        updateFilterOptions();
        let search = document.getElementById("globalSearch").value.toLowerCase();
        highlightSearch(search);
    }}
    
    document.getElementById("globalSearch").addEventListener("input", function(){{
        applyFilters();
    }});
    
    for(let i=0;i<columns.length;i++) {{
        let menu = document.getElementById('menu_'+i);
        menu.addEventListener('change', function(e){{
            if(e.target.classList.contains('filter-checkbox')) {{
                if(e.target.value !== "__ALL__") {{
                    menu.querySelector('.filter-checkbox[value="__ALL__"]').checked = false;
                }}
                let checked = menu.querySelectorAll('.filter-checkbox:checked:not([value="__ALL__"])');
                if(checked.length === 0) {{
                    menu.querySelector('.filter-checkbox[value="__ALL__"]').checked = true;
                }}
                applyFilters();
            }}
        }});
    }}
    
    function filterCheckboxes(idx, search) {{
        let menu = document.getElementById('menu_'+idx);
        search = search.toLowerCase();
        let labels = menu.querySelectorAll('label');
        labels.forEach((label, i) => {{
            if(i === 0) return;
            let text = label.textContent.toLowerCase();
            if(text.includes(search)) {{
                label.style.display = '';
            }} else {{
                label.style.display = 'none';
            }}
        }});
    }}
    
    window.toggleDropdown = function(idx) {{
        let menu = document.getElementById('menu_'+idx);
        let parent = menu.parentElement;
        let isOpen = parent.classList.contains('show');
        document.querySelectorAll('.dropdown-checkbox').forEach(el => {{
            el.classList.remove('show');
            let m = el.querySelector('.dropdown-menu');
            if(m) m.style.display = 'none';
        }});
        if(!isOpen) {{
            parent.classList.add('show');
            menu.style.display = 'block';
            let btn = parent.querySelector('.dropdown-toggle');
            let rect = btn.getBoundingClientRect();
            menu.style.top = (rect.bottom + window.scrollY) + 'px';
            menu.style.left = (rect.left + window.scrollX) + 'px';
            menu.style.width = rect.width + 'px';
            setTimeout(() => {{
                let input = menu.querySelector('input[type="text"]');
                if(input) input.focus();
            }}, 150);
        }} else {{
            menu.style.display = 'none';
        }}
    }}
    
    window.addEventListener('scroll', function() {{
        document.querySelectorAll('.dropdown-checkbox').forEach(el => {{
            el.classList.remove('show');
            let m = el.querySelector('.dropdown-menu');
            if(m) m.style.display = 'none';
        }});
    }});
    
    document.addEventListener('click', function(e){{
        if(!e.target.classList.contains('dropdown-toggle') && !e.target.classList.contains('form-control') && !e.target.closest('.dropdown-menu')) {{
            document.querySelectorAll('.dropdown-checkbox').forEach(el => {{
                el.classList.remove('show');
                let m = el.querySelector('.dropdown-menu');
                if(m) m.style.display = 'none';
            }});
        }}
    }});
    
    window.allCheck = function(idx) {{
        let menu = document.getElementById('menu_'+idx);
        let allCb = menu.querySelector('.filter-checkbox[value="__ALL__"]');
        let cbs = menu.querySelectorAll('.filter-checkbox:not([value="__ALL__"])');
        if(allCb.checked) {{
            cbs.forEach(cb => cb.checked = false);
        }}
        applyFilters();
    }}
    
    function highlightSearch(term) {{
        if(!term) return;
        let body = document.getElementById('tableBody');
        if(!body) return;
        let reg = new RegExp('('+term.replace(/[.*+?^${{}}|[\\]\\\\]/g, '\\$&')+')', 'gi');
        body.innerHTML = body.innerHTML.replace(/<mark class="highlight">(.*?)<\/mark>/g, '$1');
        if(term) {{
            body.innerHTML = body.innerHTML.replace(reg, '<mark class="highlight">$1</mark>');
        }}
    }}
    
    document.querySelectorAll('#dataTable th.sortable').forEach(function(th){{
        th.addEventListener('click', function(){{
            let colIdx = parseInt(th.getAttribute('data-col'));
            if(sortCol === colIdx){{
                sortDir = sortDir === 'asc' ? 'desc' : 'asc';
            }} else {{
                sortCol = colIdx;
                sortDir = 'asc';
            }}
            document.querySelectorAll('#dataTable th.sortable').forEach(h=>h.classList.remove('sorted-asc','sorted-desc'));
            th.classList.add(sortDir === 'asc' ? 'sorted-asc' : 'sorted-desc');
            renderTable(filteredData);
            highlightSearch(document.getElementById("globalSearch").value.toLowerCase());
        }});
    }});
    
    // Botão Limpar Filtros
    document.getElementById("clearFiltersBtn").addEventListener("click", function() {{
        for(let i=0; i<columns.length; i++) {{
            let menu = document.getElementById('menu_'+i);
            menu.querySelectorAll('.filter-checkbox').forEach(cb => cb.checked = false);
            menu.querySelector('.filter-checkbox[value="__ALL__"]').checked = true;
        }}
        document.getElementById("globalSearch").value = "";
        applyFilters();
    }});
    
    applyFilters();
    
    // DUPLICADOS
    const allDuplicates = {all_duplicates_json};
    function renderDuplicatesTable(col) {{
        let html = '';
        if(col === '_ALL_') {{
            let any = false;
            for(const c in allDuplicates) {{
                if(allDuplicates[c].length) {{
                    any = true;
                    html += `<h6>${{c}}</h6><table class="table table-bordered table-sm table-striped"><thead><tr><th>Valor Duplicado</th><th>Quantidade</th></tr></thead><tbody>`;
                    html += allDuplicates[c].map(([val,count]) => `<tr><td>${{val}}</td><td>${{count}}</td></tr>`).join('');
                    html += '</tbody></table>';
                }}
            }}
            if(!any) html = "<div class='alert alert-success'>Nenhum valor duplicado encontrado nas colunas.</div>";
        }} else {{
            let dups = allDuplicates[col];
            if(dups && dups.length) {{
                html += `<table class="table table-bordered table-sm table-striped"><thead><tr><th>Valor Duplicado</th><th>Quantidade</th></tr></thead><tbody>`;
                html += dups.map(([val,count]) => `<tr><td>${{val}}</td><td>${{count}}</td></tr>`).join('');
                html += '</tbody></table>';
            }} else {{
                html = "<div class='alert alert-success'>Nenhum valor duplicado encontrado nesta coluna.</div>";
            }}
        }}
        document.getElementById('duplicatesTableDiv').innerHTML = html;
    }}
    function setupDuplicatesTab() {{
        const sel = document.getElementById('dupColSelect');
        if(sel) {{
            renderDuplicatesTable(sel.value);
            sel.onchange = function() {{ renderDuplicatesTable(this.value); }};
        }}
    }}
    
    {pivot_js}
    
    document.addEventListener('DOMContentLoaded', function(){{
        setupDuplicatesTab();
        const tabBtn = document.getElementById('duplicates-tab');
        if(tabBtn) {{
            tabBtn.addEventListener('shown.bs.tab', function(e){{ setupDuplicatesTab(); }});
        }}
        
        if(typeof setupPivotTable === 'function') {{
            setupPivotTable();
        }}
    }});
</script>
</body>
</html>
"""
        # Verificar se temos tabela dinâmica para adicionar
        pivot_tab = ""
        pivot_pane = ""
        if pivot_html:
            pivot_tab = """
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="pivot-tab" data-bs-toggle="tab" data-bs-target="#pivot-pane" type="button" role="tab" aria-controls="pivot-pane">Dinâmica</button>
            </li>
            """
            
            pivot_pane = """
            <div class="tab-pane fade tab-pane-content" id="pivot-pane" role="tabpanel" aria-labelledby="pivot-tab">
                <div class="p-3">
                    <h5>Tabela Dinâmica</h5>
                    <div class="mb-3">
                        <strong>Configuração:</strong> 
                        Linhas: {rows}, 
                        Colunas: {cols}, 
                        Valores: {vals}, 
                        Agregação: {agg}
                    </div>
                    <div class="table-responsive">
                        {pivot_html}
                    </div>
                </div>
            </div>
            """.format(
                rows=", ".join(self.pivot_rows),
                cols=", ".join(self.pivot_columns) if self.pivot_columns else "Nenhuma",
                vals=self.pivot_values,
                agg=self.pivot_agg,
                pivot_html=pivot_html
            )

        return html.format(
            page_title=page_title,
            filters_html=filters_html,
            fixed_header=' fixed-header' if horizontal_scroll else '',
            table_striped=' table-striped' if alternate_rows else '',
            thead=thead,
            tbody=tbody,
            export_time=datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            df_len=len(df),
            all_data=json.dumps(data_json, ensure_ascii=False),
            columns=json.dumps(columns, ensure_ascii=False),
            duplicates_html=duplicates_html,
            all_duplicates_json=all_duplicates_json,
            pivot_tab=pivot_tab,
            pivot_pane=pivot_pane,
            pivot_config_js=pivot_config_js,
            pivot_js=pivot_js
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToHTMLConverter(root)
    root.mainloop()
