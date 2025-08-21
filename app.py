import os
import sys
import webbrowser
from pathlib import Path

# --- UI ---
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception as e:
    print("Failed to import UI libraries. Make sure tkinter and tkinterdnd2 are installed.")
    print(e)
    sys.exit(1)

# --- Data / Graph ---
import pandas as pd
import networkx as nx
from pyvis.network import Network

SUPPORTED = {".ods", ".xlsx", ".xls", ".csv"}


def read_table(path: str) -> pd.DataFrame:
    """Read a spreadsheet (ODS/XLSX/XLS/CSV) and return a DataFrame.
    Tries to be forgiving with separators and header capitalization.
    """
    p = Path(path)
    ext = p.suffix.lower()

    if ext not in SUPPORTED:
        raise ValueError(f"Unsupported file type: {ext}. Supported: {', '.join(sorted(SUPPORTED))}")

    if ext == ".ods":
        # Requires odfpy
        df = pd.read_excel(p, engine="odf")
    elif ext in {".xlsx", ".xls"}:
        df = pd.read_excel(p)
    else:  # .csv
        # Try common delimiters; fall back to default
        df = None
        for sep in [",", ";", " ", "|"]:
            try:
                tmp = pd.read_csv(p, sep=sep)
                if tmp.shape[1] >= 2:
                    df = tmp
                    break
            except Exception:
                continue
        if df is None:
            df = pd.read_csv(p)

    # Normalize column names
    df.columns = [str(c).strip() for c in df.columns]

    # Fuzzy locate From/To/Amount columns
    def find_col(candidates):
        for name in df.columns:
            key = name.lower().replace(" ", "")
            for c in candidates:
                if key == c:
                    return name
        # partial match fallback
        for name in df.columns:
            key = name.lower().replace(" ", "")
            if any(c in key for c in candidates):
                return name
        return None

    col_from = find_col(["from"])  # expects exact 'From' or close
    col_to = find_col(["to"])      # expects exact 'To' or close
    col_amount = find_col(["amount", "amt"])  # optional

    if not col_from or not col_to:
        raise ValueError(
            "Could not find 'From' and 'To' columns. Make sure your sheet has those headers.")

    # Keep only relevant columns to lighten memory
    keep = [col_from, col_to]
    if col_amount:
        keep.append(col_amount)
    df = df[keep].copy()

    # Standardize types
    df[col_from] = df[col_from].astype(str).str.strip()
    df[col_to] = df[col_to].astype(str).str.strip()
    if col_amount:
        df[col_amount] = pd.to_numeric(df[col_amount], errors='coerce')

    # Rename to canonical names that the rest of the code expects
    rename_map = {col_from: "From", col_to: "To"}
    if col_amount:
        rename_map[col_amount] = "Amount"
    return df.rename(columns=rename_map)


def to_graph(df: pd.DataFrame) -> nx.DiGraph:
    """Build a directed graph aggregated by (From, To). Adds edge count and total amount."""
    G = nx.DiGraph()

    has_amount = "Amount" in df.columns
    if has_amount:
        grouped = df.groupby(["From", "To"], dropna=False)["Amount"].agg(["count", "sum"]).reset_index()
        grouped.rename(columns={"count": "tx_count", "sum": "amount_sum"}, inplace=True)
    else:
        grouped = df.groupby(["From", "To"], dropna=False).size().reset_index(name="tx_count")
        grouped["amount_sum"] = None

    if grouped.empty:
        return G

    for _, row in grouped.iterrows():
        src = row["From"]; dst = row["To"]
        if pd.isna(src) or pd.isna(dst) or str(src).strip() == '' or str(dst).strip() == '':
            continue
        G.add_node(src)
        G.add_node(dst)
        # Hover tooltip for the EDGE (HTML)
        title_html = f"From: {src}->To: {dst}; Transfers: {int(row['tx_count'])}"
        if pd.notna(row["amount_sum"]):
            title_html += f"; Total amount: {row['amount_sum']:.2f}"
        # Do NOT set an edge label -> we want hover-only link data
        G.add_edge(
            src, dst,
            weight=float(row["tx_count"]),
            title=title_html,
            tx_count=int(row["tx_count"]),
            amount_sum=(None if pd.isna(row["amount_sum"]) else float(row["amount_sum"]))
        )

    # Node attributes: degree (for sizing)
    for n in G.nodes:
        deg = G.in_degree(n, weight='weight') + G.out_degree(n, weight='weight')
        G.nodes[n]['degree'] = deg

    return G


def render_pyvis(G: nx.DiGraph, out_html: Path) -> None:
    """Render an interactive network with PyVis and open it in a browser."""
    if G.number_of_nodes() == 0 or G.number_of_edges() == 0:
        raise ValueError("The network is empty. Check that your file has 'From' and 'To' columns and at least one row.")

    net = Network(height="750px", width="100%", directed=True, notebook=False)
    net.barnes_hut()

    net.from_nx(G)

    # Show only IBANs on nodes; other info stays in hover tooltips
    for node in net.nodes:
        deg = node.get('degree', 1) or 1
        node['value'] = deg
        node['title'] = node.get('id')  # tooltip shows the IBAN only
        node['label'] = str(node['id'])  # on-canvas: IBAN only
        node['shape'] = 'dot'
        node['scaling'] = {"min": 10, "max": 50}

    # Edge arrows and NO edge labels (hover-only details via 'title')
    for edge in net.edges:
        edge['arrows'] = 'to'
        if 'label' in edge:
            del edge['label']

    net.set_options('''{
      "interaction": {"hover": true, "navigationButtons": true},
      "physics": {"stabilization": true}
    }''')

    net.write_html(str(out_html), notebook=False)
    try:
        webbrowser.open(out_html.as_uri())
    except Exception:
        pass


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Drag a file to visualize transfers -> social network")
        self.geometry("800x400")

        self.style = ttk.Style(self)
        self.style.theme_use('clam')

        instructions = (
            "You know what to do..."
        )
        self.drop_area = ttk.Label(self, text=instructions, relief=tk.RIDGE, anchor=tk.CENTER)
        self.drop_area.pack(expand=True, fill=tk.BOTH, padx=16, pady=16)

        # Enable Drops
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.on_drop)

        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=16, pady=(0, 16))
        ttk.Button(btn_frame, text="Browse...", command=self.on_browse).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="Quit", command=self.destroy).pack(side=tk.RIGHT)

        self.status = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status, anchor=tk.W).pack(fill=tk.X, padx=16, pady=(0, 8))

    def on_browse(self):
        path = filedialog.askopenfilename(title="Select data file",
                                          filetypes=[
                                              ("Spreadsheets", "*.ods *.xlsx *.xls *.csv"),
                                              ("All Files", "*.*")
                                          ])
        if path:
            self.process(path)

    def on_drop(self, event):
        # On Windows, event.data may have braces if spaces exist; handle multi-file drops – take first
        raw = event.data
        paths = self._split_paths(raw)
        if not paths:
            messagebox.showerror("Drop error", "No valid file path detected.")
            return
        self.process(paths[0])

    @staticmethod
    def _split_paths(raw: str):
        # Handles formats like '{C:/path with spaces/file.ods}' or multiple files
        out = []
        buf = ''
        in_brace = False
        for ch in raw:
            if ch == '{':
                in_brace = True
                buf = ''
            elif ch == '}':
                in_brace = False
                out.append(buf)
                buf = ''
            elif ch == ' ' and not in_brace:
                if buf:
                    out.append(buf)
                    buf = ''
            else:
                buf += ch
        if buf:
            out.append(buf)
        return out

    def process(self, path: str):
        try:
            self.status.set(f"Reading: {path}")
            df = read_table(path)
            self.status.set(f"Rows loaded: {len(df):,}. Building graph...")
            G = to_graph(df)
            if G.number_of_nodes() == 0 or G.number_of_edges() == 0:
                raise ValueError("No edges to display. Ensure there are valid 'From' and 'To' rows.")
            out_html = Path(path)
            out_html = out_html.with_name(out_html.stem + "_graph.html")
            self.status.set(f"Rendering: {out_html.name}")
            render_pyvis(G, out_html)
            self.status.set(f"Done -> opened {out_html.name} in your browser")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status.set("Error - see message")


if __name__ == "__main__":
    # Improve DPI scaling on Windows for sharper UI
    if sys.platform.startswith("win"):
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)  # Per-monitor DPI Aware
        except Exception:
            pass

    app = App()
    app.mainloop()
