import os
import sys
import textwrap
import tkinter as tk
from tkinter import filedialog
import webbrowser
import json
from typing import Dict, List, Optional

import xlsxwriter
from pyvis.network import Network

############################################################
#                      BUILD TREE                          #
############################################################

def build_tree(root_dir: str) -> Dict[str, List[str]]:
    """Return {abs_path: [children]} ignoring hidden and temp files."""
    tree: Dict[str, List[str]] = {}
    for dirpath, dirnames, filenames in os.walk(root_dir):
        dirnames[:] = [d for d in dirnames if not d.startswith('.')]
        filenames = [f for f in filenames if not f.startswith('.') and not f.startswith('~$')]
        tree[os.path.normpath(dirpath)] = dirnames + filenames
    return tree

############################################################
#                     EXCEL EXPORT                         #
############################################################

def export_to_excel(tree: Dict[str, List[str]], root_dir: str, save_path: Optional[str] = None) -> Optional[str]:
    """Save *tree* to an Excel file. If *save_path* is None, show a dialog and
    suggest a name that includes the root folder. Always overwrites existing
    file so content is up‑to‑date."""

    root_name = os.path.basename(root_dir.rstrip(os.sep)) or "root"

    if save_path is None:
        default_name = f"folder_structure_{root_name}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить Excel как…",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
        )
    if not save_path:
        return None

    wb = xlsxwriter.Workbook(save_path)
    ws = wb.add_worksheet()
    bold = wb.add_format({"bold": True})

    ws.write(0, 0, os.path.basename(root_dir), bold)
    row = 1
    for dirpath, children in tree.items():
        rel = os.path.relpath(dirpath, root_dir)
        level = 0 if rel == "." else rel.count(os.sep) + 1
        folder = os.path.basename(dirpath)
        if rel != ".":
            ws.write(row, level, folder, bold)
            row += 1
        for item in children:
            if not os.path.isdir(os.path.join(dirpath, item)):
                ws.write(row, level + 1, item)
                row += 1

    wb.close()
    return save_path

############################################################
#                 VISUALISATION (PyVis)                    #
############################################################

def wrap_label(text: str, width: int = 28) -> str:
    """Break long labels so they render on multiple lines in vis.js."""
    return "\n".join(textwrap.wrap(text, width=width, break_long_words=True))


def draw_tree_web(tree: Dict[str, List[str]], root_dir: str, save_path: Optional[str] = None) -> Optional[str]:
    """Create HTML hierarchy with PyVis. File name ends with root folder.
    Auto‑switches to LR orientation if any level has >15 nodes."""

    root_name = os.path.basename(root_dir.rstrip(os.sep)) or "root"
    if save_path is None:
        default_name = f"tree_visualization_{root_name}.html"
        save_path = filedialog.asksaveasfilename(
            title="Сохранить визуализацию как…",
            defaultextension=".html",
            filetypes=[("HTML files", "*.html")],
            initialfile=default_name,
        )
    if not save_path:
        return None

    # determine widest level
    level_counts: Dict[int, int] = {}
    for path in tree.keys():
        depth = os.path.relpath(path, root_dir).count(os.sep)
        level_counts[depth] = level_counts.get(depth, 0) + 1
    max_width = max(level_counts.values(), default=0)

    orientation = "LR" if max_width > 15 else "UD"
    if orientation == "LR":
        level_sep, node_space = 320, 160
        est_height = max_width * 40 + 300
        height_css = f"{min(est_height, 6000)}px"
    else:
        level_sep, node_space = 220, 480
        height_css = "1000px"

    net = Network(height=height_css, width="100%", directed=True, bgcolor="#ffffff")
    net.set_options(json.dumps({
        "layout": {
            "hierarchical": {
                "enabled": True,
                "direction": orientation,
                "sortMethod": "directed",
                "levelSeparation": level_sep,
                "nodeSpacing": node_space,
                "treeSpacing": 500,
                "blockShifting": True,
                "edgeMinimization": True,
                "parentCentralization": True,
                "avoidOverlap": 1
            }
        },
        "physics": {"enabled": False},
        "nodes": {
            "shape": "box",
            "margin": 8,
            "font": {"size": 18, "face": "Arial", "align": "center", "multi": True},
            "color": {"background": "#ffffff", "border": "#ffffff"},
            "borderWidth": 0,
            "widthConstraint": {"maximum": 380}
        },
        "edges": {
            "arrows": {"to": {"enabled": True, "scaleFactor": 0.7}},
            "smooth": {"type": "cubicBezier", "forceDirection": "vertical", "roundness": 0.4}
        },
        "interaction": {"dragNodes": False, "dragView": True, "zoomView": True}
    }))

    root_id = os.path.normpath(root_dir)
    net.add_node(root_id, label=wrap_label(os.path.basename(root_id)), level=0)

    for dirpath, children in tree.items():
        dirpath = os.path.normpath(dirpath)
        for item in children:
            abs_path = os.path.normpath(os.path.join(dirpath, item))
            depth = os.path.relpath(abs_path, root_dir).count(os.sep)
            net.add_node(abs_path, label=wrap_label(item), level=depth)
            net.add_edge(dirpath, abs_path)

    net.save_graph(save_path)

    # make scrollable
    with open(save_path, "r", encoding="utf-8") as f:
        html = f.read()
    if "overflow" not in html:
        html = html.replace("<body>", "<body style=\"overflow:auto;\">", 1)
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(html)

    return save_path

############################################################
#                         HELPERS                          #
############################################################

def _open_file(path: str):
    if sys.platform == "darwin":
        os.system(f"open '{path}'")
    elif sys.platform.startswith("linux"):
        os.system(f"xdg-open '{path}'")
    else:
        os.startfile(path)  # type: ignore[attr-defined]

############################################################
#                         MENU                             #
############################################################

def show_menu(root_dir: str) -> None:
    if tk._default_root is None:
        hidden_root = tk.Tk()
        hidden_root.withdraw()

    menu = tk.Toplevel()
    menu.title("Меню действий")
    menu.geometry("480x350")

    excel_path: Optional[str] = None
    html_path: Optional[str] = None

    # ---------- callbacks ---------- #
    def save_excel():
        nonlocal excel_path
        tree = build_tree(root_dir)
        path = export_to_excel(tree, root_dir)
        if path:
            excel_path = path
            btn_open_excel["state"] = "normal"
            btn_refresh_excel["state"] = "normal"
            _open_file(path)

    def refresh_excel():
        if excel_path:
            tree = build_tree(root_dir)
            export_to_excel(tree, root_dir, save_path=excel_path)
            _open_file(excel_path)

    def open_excel():
        if excel_path:
            _open_file(excel_path)

    def save_html():
        nonlocal html_path
        tree = build_tree(root_dir)
        path = draw_tree_web(tree, root_dir)
        if path:
            html_path = path
            btn_open_html["state"] = "normal"
            btn_refresh_html["state"] = "normal"
            webbrowser.open(f"file://{path}")

    def refresh_html():
        if html_path:
            tree = build_tree(root_dir)
            draw_tree_web(tree, root_dir, save_path=html_path)
            webbrowser.open(f"file://{html_path}")

    def open_html():
        if html_path:
            webbrowser.open(f"file://{html_path}")

    def choose_new():
        menu.destroy()
        main()

    def exit_app():
        menu.destroy()
        sys.exit()

    # ---------- widgets ---------- #
    tk.Label(menu, text="Выберите действие", font=("Arial", 14)).pack(pady=15)

    # Excel frame
    frame_excel = tk.Frame(menu)
    frame_excel.pack(pady=6)
    tk.Label(frame_excel, text="Excel", font=("Arial", 12)).grid(row=0, columnspan=3, pady=(0, 4))
    tk.Button(frame_excel, text="Сохранить", width=14, command=save_excel).grid(row=1, column=0, padx=2)
    btn_open_excel = tk.Button(frame_excel, text="Открыть", width=14, state="disabled", command=open_excel)
    btn_open_excel.grid(row=1, column=1, padx=2)
    btn_refresh_excel = tk.Button(frame_excel, text="Обновить", width=14, state="disabled", command=refresh_excel)
    btn_refresh_excel.grid(row=1, column=2, padx=2)

    # HTML frame
    frame_html = tk.Frame(menu)
    frame_html.pack(pady=10)
    tk.Label(frame_html, text="HTML-дерево", font=("Arial", 12)).grid(row=0, columnspan=3, pady=(0, 4))
    tk.Button(frame_html, text="Сохранить", width=14, command=save_html).grid(row=1, column=0, padx=2)
    btn_open_html = tk.Button(frame_html, text="Открыть", width=14, state="disabled", command=open_html)
    btn_open_html.grid(row=1, column=1, padx=2)
    btn_refresh_html = tk.Button(frame_html, text="Обновить", width=14, state="disabled", command=refresh_html)
    btn_refresh_html.grid(row=1, column=2, padx=2)

    # bottom buttons
    tk.Button(menu, text="Выбрать другую папку", width=24, command=choose_new).pack(pady=4)
    tk.Button(menu, text="Выход", width=24, command=exit_app).pack(pady=4)


########################################################################
#                                MAIN                                  #
########################################################################

def main() -> None:
    def select_and_proceed():
        start.destroy()
        folder = filedialog.askdirectory(title="Выберите папку")
        if not folder:
            sys.exit()
        show_menu(folder)

    def close_program():
        start.destroy()
        sys.exit()

    start = tk.Tk()
    start.title("Стартовое окно")
    start.geometry("420x220")

    tk.Label(start, text="Выберите действие", font=("Arial", 14)).pack(pady=20)
    tk.Button(start, text="Выбрать папку", width=22, command=select_and_proceed).pack(pady=5)
    tk.Button(start, text="Выход", width=22, command=close_program).pack(pady=5)

    start.mainloop()


if __name__ == "__main__":
    main()