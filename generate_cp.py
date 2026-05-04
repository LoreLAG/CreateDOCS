import pandas as pd
import openpyxl
from copy import deepcopy, copy
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from openpyxl.styles import Alignment, Font, Border, Side
from datetime import datetime
from thefuzz import process, fuzz  # Imported fuzz for cross-scoring


def find_best_matches(phase_info, pure_phases_db, blocks_db):
    """
    Finds suggestions and calculates the absolute best match
    based on Phase, Department, and penalizing Int/Ext inconsistencies.
    """
    gui_phase = str(phase_info['base_phase']).strip().upper()
    gui_dept = str(phase_info.get('department', '')).strip().upper()
    gui_ext_int = str(phase_info.get('ext_int', '')).strip().upper()

    # 1. Find suggested candidates based on the PHASE NAME
    match_uppers = [db_f.upper() for db_f in pure_phases_db if gui_phase in db_f.upper()]
    if not match_uppers and pure_phases_db:
        matches = process.extract(gui_phase, pure_phases_db, limit=15)
        match_uppers = [m[0].upper() for m in matches if m[1] >= 60]

    match_uppers = list(dict.fromkeys(match_uppers))
    suggested = [k for k in blocks_db.keys() if k[0] in match_uppers]

    # 2. Calculate the best match considering DEPARTMENT and INT/EXT
    best_key = None
    best_score = -999

    is_internal_gui = gui_ext_int in ["INT", "INTERNAL", "INT.", "I"]
    is_external_gui = gui_ext_int in ["EXT", "EXTERNAL", "EXT.", "E"]

    for k in suggested:
        db_f, db_t = k[0], k[1]

        # Base scores
        score_phase = fuzz.WRatio(gui_phase, db_f)
        score_dept = fuzz.WRatio(gui_dept, db_t) if gui_dept else 0

        # Int/Ext Penalty Logic
        penalty = 0
        if is_internal_gui and ("EXT" in db_t or "EXTERNAL" in db_t):
            penalty += 50
        if is_external_gui and ("INT" in db_t or "INTERNAL" in db_t):
            penalty += 50

        # Total Calculation (Phase counts 60%, Dept 40%, minus the penalty)
        total_score = (score_phase * 0.6) + (score_dept * 0.4) - penalty

        if total_score > best_score:
            best_score = total_score
            best_key = k

    return suggested, best_key


def ask_global_choices(phases, pure_phases_db, blocks_db, choices_cache):
    """
    Opens a 3-column interface:
    LEFT: Database | CENTER: Current Phase Composition | RIGHT: Phase Bookmarks
    """
    final_choice = {"abort": True, "choices": {}}

    parent = tk._default_root
    if parent:
        fw = parent.focus_get()
        if fw:
            parent = fw.winfo_toplevel()

    root = tk.Toplevel(parent)
    if parent:
        root.transient(parent)
    root.withdraw()
    root.title("Global Control Plan Composer")

    # ─── GLOBAL INTERFACE STATE ───
    state = {
        "active_idx": 0,
        "choices": {},
        "current_key": None,
        "current_suggested_keys": [],
        "visited": set()
    }

    # Pre-calculate matches for all phases with the new smart logic
    for i, p in enumerate(phases):
        gui_phase = str(p['base_phase']).strip()
        key = f"{p['number']}_{gui_phase}"

        if key in choices_cache:
            state["choices"][key] = deepcopy(choices_cache[key])
            state["visited"].add(i)
        else:
            suggested, best_key = find_best_matches(p, pure_phases_db, blocks_db)

            if best_key:
                state["choices"][key] = [{"key": best_key, "alias": None}]
            else:
                state["choices"][key] = []

    # ─── MAIN LAYOUT ───
    main_frame = tk.Frame(root)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # RIGHT FRAME: Phase Bookmarks
    nav_frame = ttk.LabelFrame(main_frame, text="Phases (Use Up/Down arrows)")
    nav_frame.pack(side="right", fill="y", padx=(10, 0))

    scroll_nav = ttk.Scrollbar(nav_frame)
    scroll_nav.pack(side="right", fill="y")
    listbox_nav = tk.Listbox(nav_frame, selectmode=tk.SINGLE, yscrollcommand=scroll_nav.set, width=45, font=("Segoe UI", 10))
    listbox_nav.pack(side="left", fill="both", expand=True, padx=5, pady=5)
    scroll_nav.config(command=listbox_nav.yview)

    # LEFT AND CENTER FRAME
    editor_frame = tk.Frame(main_frame)
    editor_frame.pack(side="left", fill="both", expand=True)

    lbl_current_phase = ttk.Label(editor_frame, text="", font=("Segoe UI", 12, "bold"), foreground="#005599")
    lbl_current_phase.pack(anchor="w", pady=(0, 10))

    columns_frame = tk.Frame(editor_frame)
    columns_frame.pack(fill="both", expand=True)

    columns_frame.columnconfigure(0, weight=1)
    columns_frame.columnconfigure(1, weight=0)
    columns_frame.columnconfigure(2, weight=1)
    columns_frame.columnconfigure(3, weight=0)
    columns_frame.rowconfigure(2, weight=1)

    # --- ROW 0: Headers ---
    top_db_frame = tk.Frame(columns_frame)
    top_db_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
    lbl_title_left = ttk.Label(top_db_frame, text="Suggested variants:", font=("Segoe UI", 9, "bold"))
    lbl_title_left.pack(side="left")

    show_all_var = tk.BooleanVar(value=False)
    cb_show_all = tk.Checkbutton(top_db_frame, text="Show all", variable=show_all_var)
    cb_show_all.pack(side="right")

    top_cp_frame = tk.Frame(columns_frame)
    top_cp_frame.grid(row=0, column=2, sticky="ew", pady=(0, 5))
    ttk.Label(top_cp_frame, text="Blocks for this phase (Right-Click for Alias):", font=("Segoe UI", 9, "bold")).pack(side="left")

    # --- ROW 1: Filter ---
    filter_var = tk.StringVar()
    filter_entry = ttk.Entry(columns_frame, textvariable=filter_var)
    filter_entry.grid(row=1, column=0, sticky="ew", pady=(0, 5))
    filter_entry.insert(0, "Search phase...")

    def on_filter_in(e):
        if filter_var.get() == "Search phase...":
            filter_entry.delete(0, tk.END)

    def on_filter_out(e):
        if not filter_var.get():
            filter_entry.insert(0, "Search phase...")

    filter_entry.bind("<FocusIn>", on_filter_in)
    filter_entry.bind("<FocusOut>", on_filter_out)

    # --- ROW 2: Listboxes ---
    list_frame_db = tk.Frame(columns_frame)
    list_frame_db.grid(row=2, column=0, sticky="nsew")
    scroll_db = ttk.Scrollbar(list_frame_db)
    scroll_db.pack(side="right", fill="y")
    listbox_db = tk.Listbox(list_frame_db, selectmode=tk.EXTENDED, yscrollcommand=scroll_db.set, width=55)
    listbox_db.pack(side="left", fill="both", expand=True)
    scroll_db.config(command=listbox_db.yview)

    list_frame_cp = tk.Frame(columns_frame)
    list_frame_cp.grid(row=2, column=2, sticky="nsew")
    scroll_cp = ttk.Scrollbar(list_frame_cp)
    scroll_cp.pack(side="right", fill="y")
    listbox_cp = tk.Listbox(list_frame_cp, selectmode=tk.SINGLE, yscrollcommand=scroll_cp.set, width=65)
    listbox_cp.pack(side="left", fill="both", expand=True)
    scroll_cp.config(command=listbox_cp.yview)

    # --- CENTRAL AND SIDE BUTTONS ---
    btn_frame = tk.Frame(columns_frame)
    btn_frame.grid(row=2, column=1, sticky="n", padx=10, pady=40)
    ttk.Button(btn_frame, text="Add DB >>", command=lambda: add_selected()).pack(pady=5)
    ttk.Button(btn_frame, text="Add Empty >>", command=lambda: add_empty()).pack(pady=5)

    ord_frame = tk.Frame(columns_frame)
    ord_frame.grid(row=2, column=3, sticky="n", padx=5, pady=40)
    ttk.Button(ord_frame, text="⬆", width=3, command=lambda: move_up()).pack(pady=2)
    ttk.Button(ord_frame, text="⬇", width=3, command=lambda: move_down()).pack(pady=2)
    ttk.Button(ord_frame, text="❌", width=3, command=lambda: remove_sel()).pack(pady=(15, 2))

    # ─── HELPER FUNCTIONS ───
    def get_display_text(item_dict):
        key = item_dict["key"]
        alias = item_dict.get("alias")
        base_text = "[EMPTY] Empty row" if key == "EMPTY" else f"[DB] {key[0]} - {key[1]}"
        if alias:
            return f"[ALIAS: {alias}] (Orig: {key[0] if key != 'EMPTY' else 'Empty'})"
        return base_text

    def render_nav_list():
        listbox_nav.delete(0, tk.END)
        for i, p in enumerate(phases):
            key = f"{p['number']}_{str(p['base_phase']).strip()}"
            has_items = len(state["choices"].get(key, [])) > 0

            if not has_items:
                status = "❌"
                color = "red"
            elif i not in state["visited"]:
                status = "❗"
                color = "#b85c00"
            else:
                status = "✅"
                color = "#008000"

            disp_phase = p.get('phase', p['base_phase'])
            text = f"{status} Op.{p['number']} - {disp_phase}"
            listbox_nav.insert(tk.END, text)

            if color != "black":
                listbox_nav.itemconfig(i, {'fg': color})

        listbox_nav.selection_clear(0, tk.END)
        listbox_nav.selection_set(state["active_idx"])
        listbox_nav.see(state["active_idx"])

    def apply_db_filter(*args):
        q = filter_var.get().lower()
        if q == "search phase...":
            q = ""
        listbox_db.delete(0, tk.END)

        base_dict = blocks_db if show_all_var.get() else state["current_suggested_keys"]
        base_keys = list(base_dict.keys()) if isinstance(base_dict, dict) else base_dict

        current_view_keys = [k for k in base_keys if not q or q in k[0].lower() or q in k[1].lower()]
        listbox_db.current_keys = current_view_keys

        for k in current_view_keys:
            listbox_db.insert(tk.END, f"{k[0]} - {k[1]}")

        if not current_view_keys:
            listbox_db.insert(tk.END, "No results.")
            listbox_db.config(state=tk.DISABLED)
        else:
            listbox_db.config(state=tk.NORMAL)

    filter_var.trace_add("write", apply_db_filter)

    def load_phase(idx):
        if idx < 0 or idx >= len(phases):
            return

        state["active_idx"] = idx
        state["visited"].add(idx)

        p = phases[idx]
        gui_phase = str(p['base_phase']).strip()
        state["current_key"] = f"{p['number']}_{gui_phase}"

        disp_phase = p.get('phase', gui_phase)
        disp_dept = p.get('department', 'N/A')
        lbl_current_phase.config(text=f"⚙️ Editing: Op. {p['number']} - {disp_phase} [{disp_dept}]")

        suggested, _ = find_best_matches(p, pure_phases_db, blocks_db)
        state["current_suggested_keys"] = suggested

        if not suggested:
            show_all_var.set(True)

        lbl_title_left.config(text="All variants (DB):" if show_all_var.get() else "Suggested variants:")

        if filter_var.get() != "Search phase...":
            filter_var.set("")
        apply_db_filter()

        listbox_cp.delete(0, tk.END)
        for item in state["choices"][state["current_key"]]:
            listbox_cp.insert(tk.END, get_display_text(item))

        render_nav_list()

    def on_nav_click(event):
        sel = listbox_nav.curselection()
        if sel:
            load_phase(sel[0])

    listbox_nav.bind("<<ListboxSelect>>", on_nav_click)
    cb_show_all.config(command=lambda: load_phase(state["active_idx"]))

    # ─── EDITING AND KEYBOARD ACTIONS ───
    def nav_up(e):
        load_phase(state["active_idx"] - 1)
        return "break"

    def nav_down(e):
        load_phase(state["active_idx"] + 1)
        return "break"

    def global_nav_up(e):
        if e.widget in (listbox_db, listbox_cp, filter_entry):
            return
        nav_up(e)

    def global_nav_down(e):
        if e.widget in (listbox_db, listbox_cp, filter_entry):
            return
        nav_down(e)

    listbox_nav.bind("<Up>", nav_up)
    listbox_nav.bind("<Down>", nav_down)
    root.bind("<Up>", global_nav_up)
    root.bind("<Down>", global_nav_down)

    def add_selected(event=None):
        if not getattr(listbox_db, 'current_keys', None):
            return
        sel = listbox_db.curselection()
        for i in sel:
            item = {"key": listbox_db.current_keys[i], "alias": None}
            state["choices"][state["current_key"]].append(item)
            listbox_cp.insert(tk.END, get_display_text(item))
        listbox_db.selection_clear(0, tk.END)
        render_nav_list()

    def add_empty():
        item = {"key": "EMPTY", "alias": None}
        state["choices"][state["current_key"]].append(item)
        listbox_cp.insert(tk.END, get_display_text(item))
        render_nav_list()

    def set_alias(event):
        idx = listbox_cp.nearest(event.y)
        if idx < 0 or idx >= listbox_cp.size():
            return
        bbox = listbox_cp.bbox(idx)
        if not bbox or not (bbox[1] <= event.y <= bbox[1] + bbox[3]):
            return

        listbox_cp.selection_clear(0, tk.END)
        listbox_cp.selection_set(idx)

        item = state["choices"][state["current_key"]][idx]
        current_alias = item.get("alias", "") or ""

        new_alias = simpledialog.askstring(
            "Set Alias",
            "Enter the name you want to appear on the Control Plan:",
            initialvalue=current_alias, parent=root
        )

        if new_alias is not None:
            item["alias"] = new_alias.strip() if new_alias.strip() else None
            listbox_cp.delete(idx)
            listbox_cp.insert(idx, get_display_text(item))
            listbox_cp.selection_set(idx)

    def remove_sel(event=None):
        sel = listbox_cp.curselection()
        if not sel:
            return
        idx = sel[0]
        listbox_cp.delete(idx)
        state["choices"][state["current_key"]].pop(idx)
        if listbox_cp.size() > 0:
            listbox_cp.selection_set(max(0, idx - 1))
        render_nav_list()

    def move_up():
        sel = listbox_cp.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx > 0:
            lst = state["choices"][state["current_key"]]
            lst[idx - 1], lst[idx] = lst[idx], lst[idx - 1]
            val = listbox_cp.get(idx)
            listbox_cp.delete(idx)
            listbox_cp.insert(idx - 1, val)
            listbox_cp.selection_set(idx - 1)

    def move_down():
        sel = listbox_cp.curselection()
        if not sel:
            return
        idx = sel[0]
        lst = state["choices"][state["current_key"]]
        if idx < len(lst) - 1:
            lst[idx + 1], lst[idx] = lst[idx], lst[idx + 1]
            val = listbox_cp.get(idx)
            listbox_cp.delete(idx)
            listbox_cp.insert(idx + 1, val)
            listbox_cp.selection_set(idx + 1)

    listbox_db.bind("<Double-1>", add_selected)
    listbox_cp.bind("<Double-1>", remove_sel)
    listbox_cp.bind("<Button-3>", set_alias)

    # ─── BOTTOM: GLOBAL CONFIRMATION ───
    bot_frame = tk.Frame(root)
    bot_frame.pack(fill="x", pady=10, padx=10)

    def confirm_all(event=None):
        empty_phases = [p['number'] for p in phases if not state["choices"][f"{p['number']}_{str(p['base_phase']).strip()}"]]
        if empty_phases:
            msg = f"Warning: the following operations have no assigned blocks:\nOp. {', '.join(map(str, empty_phases))}\n\nDo you want to automatically fill these phases with empty rows?"
            ans = messagebox.askyesnocancel("Incomplete phases", msg, parent=root)
            if ans is None:
                return
            if ans is True:
                for v in empty_phases:
                    for p in phases:
                        if p['number'] == v:
                            ch = f"{p['number']}_{str(p['base_phase']).strip()}"
                            state["choices"][ch].append({"key": "EMPTY", "alias": None})
                            break

        final_choice["abort"] = False
        final_choice["choices"] = state["choices"]
        root.destroy()

    ttk.Button(bot_frame, text="Save and Generate CP", command=confirm_all, style="Accent.TButton").pack()
    root.bind("<Return>", confirm_all)

    if phases:
        load_phase(0)

    root.update_idletasks()
    window_width = 1250
    window_height = 650
    try:
        if parent and parent.winfo_viewable():
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            pos_x = px + (pw // 2) - (window_width // 2)
            pos_y = py + (ph // 2) - (window_height // 2)
        else:
            sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
            pos_x = (sw // 2) - (window_width // 2)
            pos_y = (sh // 2) - (window_height // 2)
    except Exception:
        pos_x, pos_y = 100, 100

    root.geometry(f"{window_width}x{window_height}+{pos_x}+{pos_y}")
    root.deiconify()
    root.grab_set()
    root.focus_force()
    root.wait_window()

    return None if final_choice["abort"] else final_choice["choices"]


def generate_cp(template_path, masters_path, output_path, phases, title=None, rev="00", customer_only=False, choices_cache=None):
    if not os.path.exists(masters_path):
        raise FileNotFoundError(f"MASTERS Database not found: {masters_path}")

    if customer_only:
        phases = [p for p in phases if p.get('is_customer', False)]
        if not phases:
            return

    df = pd.read_excel(masters_path, sheet_name='functional_database')
    df = df.dropna(how='all')
    col_phase = df.columns[0]
    col_type = df.columns[1]

    pure_phases_db = df[col_phase].dropna().astype(str).str.strip().unique().tolist()

    blocks_db = {}
    for (f, t), group in df.groupby([col_phase, col_type], sort=False):
        blocks_db[(str(f).strip().upper(), str(t).strip().upper())] = group

    if choices_cache is None:
        choices_cache = {}

    # OPEN GLOBAL COMPOSER
    new_choices = ask_global_choices(phases, pure_phases_db, blocks_db, choices_cache)
    if new_choices is None:
        raise InterruptedError("Generation cancelled by the user in the Control Plan.")

    # Update cache with global changes
    choices_cache.update(new_choices)

    start_row = 11
    cp_rows_main = []
    merge_info_main = []
    current_row_main = start_row

    for phase_info in phases:
        gui_phase = str(phase_info['base_phase']).strip()
        op_num = phase_info['number']
        disp_phase = phase_info.get('phase', gui_phase)

        phase_key = f"{op_num}_{gui_phase}"
        choices_list = choices_cache.get(phase_key, [])

        for choice_dict in choices_list:
            # Compatibility for safety
            if isinstance(choice_dict, dict):
                choice_item = choice_dict["key"]
                alias = choice_dict.get("alias")
            else:
                choice_item = choice_dict
                alias = None

            if choice_item == "EMPTY":
                chosen_block_df = pd.DataFrame([["", phase_info.get('ext_int', '')] + [""] * (len(df.columns) - 2)], columns=df.columns)
            else:
                chosen_block_df = blocks_db.get(choice_item)
                if chosen_block_df is None:
                    chosen_block_df = pd.DataFrame([["", phase_info.get('ext_int', '')] + [""] * (len(df.columns) - 2)], columns=df.columns)

            num_block_rows = len(chosen_block_df)
            merge_info_main.append((current_row_main, current_row_main + num_block_rows - 1))
            current_row_main += num_block_rows

            for _, row in chosen_block_df.iterrows():
                row_list = ["" if pd.isna(x) else x for x in row.tolist()]

                # Apply the alias or the GUI name
                name_to_print = alias.upper() if alias else disp_phase.upper()
                row_list[0] = name_to_print

                final_row = [op_num] + row_list
                cp_rows_main.append(final_row)

    wb = openpyxl.load_workbook(template_path)
    ws_main = wb.active
    ws_main.title = "Control Plan CUST" if customer_only else "Control Plan"

    if len(cp_rows_main) > 1:
        ws_main.insert_rows(start_row + 1, amount=len(cp_rows_main) - 1)

    today = datetime.now().strftime("%d/%m/%Y")
    ws_main['J7'] = f"Date: {today}"
    ws_main['P7'] = f"Rev: {rev}"
    if title:
        ws_main['C5'] = title

    thin = Side(style='thin', color="000000")
    for r_idx, row_data in enumerate(cp_rows_main):
        r_excel = start_row + r_idx
        ws_main.row_dimensions[r_excel].height = 23
        for c_idx, val in enumerate(row_data):
            c_excel = c_idx + 1
            if c_excel > ws_main.max_column:
                break

            cell = ws_main.cell(row=r_excel, column=c_excel)
            cell.value = val

            if cell.font:
                new_font = copy(cell.font)
                new_font.name = 'Calibri Light'
                new_font.scheme = None
                cell.font = new_font
            else:
                cell.font = Font(name='Calibri Light', size=11, scheme=None)

            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # MERGE CELLS BY BLOCK
    for r_start, r_end in merge_info_main:
        if r_start != r_end:
            for col in [1, 2, 3]:
                ws_main.merge_cells(start_row=r_start, start_column=col, end_row=r_end, end_column=col)
                top_cell = ws_main.cell(row=r_start, column=col)
                top_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # CLEAN UP HEADER AND BODY
    for row in ws_main.iter_rows(min_row=1, max_row=10):
        for cell in row:
            if type(cell).__name__ == 'MergedCell':
                continue
            if cell.font:
                new_font = copy(cell.font)
                new_font.name = 'Calibri Light'
                new_font.scheme = None
                cell.font = new_font
            else:
                cell.font = Font(name='Calibri Light', scheme=None)

    ws_main.sheet_view.topLeftCell = 'A1'
    ws_main.freeze_panes = 'A11'

    wb.save(output_path)
    print(f"Control Plan generated in: {output_path}")
