# !/usr/bin/env python3
"""
Amazon Product Analysis Dashboard
A professional desktop application for analyzing Amazon product data.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.backends.backend_tkagg as tkagg
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.gridspec as gridspec
import re
import os
import webbrowser
import threading
from datetime import datetime

# ── ReportLab ──────────────────────────────────────────────────────────────
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                     Paragraph, Spacer, Image as RLImage,
                                     HRFlowable, PageBreak)
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    pdfmetrics.registerFont(TTFont('CustomFont', 'E:\DATA_SCIENCE\proj\dejavu-sans\DejaVuSans.ttf'))
    pdfmetrics.registerFont(TTFont('CustomFont-Bold', 'E:\DATA_SCIENCE\proj\dejavu-sans\DejaVuSans-Bold.ttf'))
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

# ═══════════════════════════════════════════════════════════════════════════
# THEME DEFINITIONS
# ═══════════════════════════════════════════════════════════════════════════
THEMES = {
    "Amazon Dark": {
        "bg": "#0F1111",
        "surface": "#1A1F2E",
        "card": "#232F3E",
        "accent": "#FF9900",
        "accent2": "#146EB4",
        "text": "#FFFFFF",
        "text2": "#AAAAAA",
        "success": "#067D62",
        "warning": "#FF9900",
        "danger": "#CC0C39",
        "border": "#37475A",
        "highlight": "#FF9900",
        "button_bg": "#FF9900",
        "button_fg": "#0F1111",
        "header_bg": "#232F3E",
        "treeview_bg": "#1A1F2E",
        "treeview_fg": "#FFFFFF",
        "treeview_sel": "#37475A",
        "sidebar": "#131A22",
        "plot_bg": "#1A1F2E",
        "plot_fg": "#FFFFFF",
        "plot_grid": "#37475A",
    },
    "Ocean Blue": {
        "bg": "#0D1B2A",
        "surface": "#1B2838",
        "card": "#1E3A5F",
        "accent": "#00B4D8",
        "accent2": "#0077B6",
        "text": "#FFFFFF",
        "text2": "#90E0EF",
        "success": "#06D6A0",
        "warning": "#FFB703",
        "danger": "#EF233C",
        "border": "#2D4A6E",
        "highlight": "#00B4D8",
        "button_bg": "#00B4D8",
        "button_fg": "#0D1B2A",
        "header_bg": "#1E3A5F",
        "treeview_bg": "#1B2838",
        "treeview_fg": "#FFFFFF",
        "treeview_sel": "#2D4A6E",
        "sidebar": "#0D1B2A",
        "plot_bg": "#1B2838",
        "plot_fg": "#FFFFFF",
        "plot_grid": "#2D4A6E",
    },
    "Forest Green": {
        "bg": "#0A1628",
        "surface": "#132A1E",
        "card": "#1A3A28",
        "accent": "#52B788",
        "accent2": "#2D6A4F",
        "text": "#FFFFFF",
        "text2": "#B7E4C7",
        "success": "#52B788",
        "warning": "#F4A261",
        "danger": "#E76F51",
        "border": "#2D6A4F",
        "highlight": "#52B788",
        "button_bg": "#52B788",
        "button_fg": "#0A1628",
        "header_bg": "#1A3A28",
        "treeview_bg": "#132A1E",
        "treeview_fg": "#FFFFFF",
        "treeview_sel": "#2D6A4F",
        "sidebar": "#0A1628",
        "plot_bg": "#132A1E",
        "plot_fg": "#FFFFFF",
        "plot_grid": "#2D6A4F",
    },
    "Light Mode": {
        "bg": "#F8F9FA",
        "surface": "#FFFFFF",
        "card": "#FFFFFF",
        "accent": "#FF9900",
        "accent2": "#146EB4",
        "text": "#111111",
        "text2": "#555555",
        "success": "#067D62",
        "warning": "#FF9900",
        "danger": "#CC0C39",
        "border": "#DDDDDD",
        "highlight": "#FF9900",
        "button_bg": "#FF9900",
        "button_fg": "#111111",
        "header_bg": "#232F3E",
        "treeview_bg": "#FFFFFF",
        "treeview_fg": "#111111",
        "treeview_sel": "#FFE8B3",
        "sidebar": "#232F3E",
        "plot_bg": "#FFFFFF",
        "plot_fg": "#111111",
        "plot_grid": "#DDDDDD",
    },
}

# ═══════════════════════════════════════════════════════════════════════════
# DATA PROCESSING
# ═══════════════════════════════════════════════════════════════════════════

def clean_price(val):
    if pd.isna(val): return np.nan
    s = str(val).replace("₹","").replace(",","").strip()
    m = re.search(r"[\d.]+", s)
    return float(m.group()) if m else np.nan

def clean_rating(val):
    if pd.isna(val): return np.nan
    m = re.search(r"[\d.]+", str(val))
    return float(m.group()) if m else np.nan

def clean_reviews(val):
    if pd.isna(val): return np.nan
    s = str(val).replace(",","").replace("(","").replace(")","").strip()
    if "K" in s.upper():
        m = re.search(r"[\d.]+", s)
        return float(m.group())*1000 if m else np.nan
    m = re.search(r"[\d.]+", s)
    return float(m.group()) if m else np.nan

def load_and_clean(path):
    df = pd.read_excel(path, engine="openpyxl")
    # Normalize column names
    df.columns = [c.strip() for c in df.columns]
    rename = {}
    for c in df.columns:
        cl = c.lower()
        if "product" in cl or "name" in cl:
            rename[c] = "Product Name"
        elif "rating" in cl:
            rename[c] = "Rating"
        elif "review" in cl or "count" in cl:
            rename[c] = "ReviewsCount"
        elif "price" in cl:
            rename[c] = "Price"
        elif "link" in cl or "url" in cl:
            rename[c] = "Link"
    df.rename(columns=rename, inplace=True)
    for col in ["Product Name","Rating","ReviewsCount","Price","Link"]:
        if col not in df.columns:
            df[col] = np.nan
    df = df[df["Product Name"].notna()].copy()
    df["Price_num"]   = df["Price"].apply(clean_price)
    df["Rating_num"]  = df["Rating"].apply(clean_rating)
    df["Reviews_num"] = df["ReviewsCount"].apply(clean_reviews)
    df.dropna(subset=["Price_num","Rating_num"], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df

def get_top10(df):
    df = df.copy()

    # IMDb part
    C = df["Rating_num"].mean()
    m = df["Reviews_num"].quantile(0.75)

    df["imdb_score"] = (
        (df["Reviews_num"] / (df["Reviews_num"] + m)) * df["Rating_num"] +
        (m / (df["Reviews_num"] + m)) * C
    )

    # Price normalization (cheap = better)
    p_min, p_max = df["Price_num"].min(), df["Price_num"].max()
    df["price_norm"] = 1 - ((df["Price_num"] - p_min) / (p_max - p_min + 1e-9))

    # Final score
    df["Score"] = (
        0.8 * df["imdb_score"] +
        0.2 * df["price_norm"]
    )

    return df.sort_values("Score", ascending=False).head(10)

def get_decent_deals(df):
    lo, hi = df["Price_num"].quantile(0.3), df["Price_num"].quantile(0.7)
    return df[(df["Price_num"]>=lo)&(df["Price_num"]<=hi)&(df["Rating_num"]>=4.0)]

def summary_stats(df):
    return {
        "Total Products": len(df),
        "Avg Rating": round(df["Rating_num"].mean(),2),
        "Avg Price (₹)": round(df["Price_num"].mean(),2),
        "Min Price (₹)": round(df["Price_num"].min(),2),
        "Max Price (₹)": round(df["Price_num"].max(),2),
        "Avg Reviews": int(df["Reviews_num"].fillna(0).mean()),
        "Top Rated": round(df["Rating_num"].max(),1),
    }

# ═══════════════════════════════════════════════════════════════════════════
# PDF GENERATION
# ═══════════════════════════════════════════════════════════════════════════

def generate_pdf(df, top10, deals, stats, fig, out_path):
    if not REPORTLAB_OK:
        messagebox.showerror("Error","ReportLab not installed.")
        return False
    doc = SimpleDocTemplate(out_path, pagesize=A4,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    story = []

    # Title
    title_style = ParagraphStyle("Title2", parent=styles["Title"],
        fontSize=22, textColor=colors.HexColor("#FF9900"), spaceAfter=4)
    sub_style = ParagraphStyle("Sub", parent=styles["Normal"],
        fontSize=10, textColor=colors.HexColor("#555555"), spaceAfter=12)
    story.append(Paragraph("📊 Amazon Product Analysis Report", title_style))
    story.append(Paragraph(f"Generated on {datetime.now().strftime('%d %b %Y, %H:%M')}", sub_style))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#FF9900")))
    story.append(Spacer(1, 0.3*cm))

    # Summary stats table
    story.append(Paragraph("Summary Statistics", styles["Heading2"]))
    stat_data = [["Metric","Value"]] + [[k,str(v)] for k,v in stats.items()]
    t = Table(stat_data, colWidths=[8*cm,6*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#232F3E")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'CustomFont'),
        ("FONTSIZE",(0,0),(-1,-1),10),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.HexColor("#F8F9FA"),colors.white]),
        ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#DDDDDD")),
        ("ALIGN",(1,1),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),8),
        ("RIGHTPADDING",(0,0),(-1,-1),8),
        ("TOPPADDING",(0,0),(-1,-1),6),
        ("BOTTOMPADDING",(0,0),(-1,-1),6),
    ]))
    story.append(t)
    story.append(Spacer(1,0.4*cm))

    # Save chart to temp image
    import tempfile
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    img_path = tmp.name
    tmp.close()
    
    fig.savefig(img_path, dpi=100, bbox_inches="tight",
                facecolor="#FFFFFF", edgecolor="none")
    
    story.append(Paragraph("Data Visualizations", styles["Heading2"]))
    story.append(RLImage(img_path, width=20*cm, height=14*cm))
    story.append(Spacer(0,5,0.4*cm))

    # Top 10 table
    story.append(PageBreak())
    story.append(Paragraph("Top 10 Products", styles["Heading2"]))
    cols = ["Product Name","Rating_num","Price_num","Reviews_num"]
    heads = ["Product","Rating","Price (₹)","Reviews"]
    rows = [heads]
    for _, r in top10.iterrows():
        name = str(r["Product Name"])[:65]+"…" if len(str(r["Product Name"]))>65 else str(r["Product Name"])
        rows.append([name, f"{r['Rating_num']:.1f}",
                     f"₹{r['Price_num']:.0f}", f"{int(r['Reviews_num']):,}"])
    t2 = Table(rows, colWidths=[10*cm, 2*cm, 2.5*cm, 2.5*cm])
    t2.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#FF9900")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'CustomFont'),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.HexColor("#FFF8EC"),colors.white]),
        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#DDDDDD")),
        ("ALIGN",(1,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),5),
        ("RIGHTPADDING",(0,0),(-1,-1),5),
        ("TOPPADDING",(0,0),(-1,-1),5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),
    ]))
    story.append(t2)
    story.append(Spacer(1,0.4*cm))

    # Decent deals table
    story.append(Paragraph("Decent Deals (Best Value Products)", styles["Heading2"]))
    rows2 = [heads]
    for _, r in deals.head(15).iterrows():
        name = str(r["Product Name"])[:65]+"…" if len(str(r["Product Name"]))>65 else str(r["Product Name"])
        rows2.append([name, f"{r['Rating_num']:.1f}",
                      f"₹{r['Price_num']:.0f}", f"{int(r['Reviews_num']):,}" if not pd.isna(r['Reviews_num']) else "N/A"])
    t3 = Table(rows2, colWidths=[10*cm, 2*cm, 2.5*cm, 2.5*cm])
    t3.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#067D62")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'CustomFont'),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.HexColor("#F0FFF8"),colors.white]),
        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#DDDDDD")),
        ("ALIGN",(1,0),(-1,-1),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),5),
        ("RIGHTPADDING",(0,0),(-1,-1),5),
        ("TOPPADDING",(0,0),(-1,-1),5),
        ("BOTTOMPADDING",(0,0),(-1,-1),5),
    ]))
    story.append(t3)

    doc.build(story)
    return True


# ═══════════════════════════════════════════════════════════════════════════
# MAIN APPLICATION
# ═══════════════════════════════════════════════════════════════════════════

class AmazonDashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Amazon Product Analysis Dashboard")
        self.geometry("1280x820")
        self.minsize(1100, 700)
        self.configure(bg="#0F1111")

        self.df = None
        self.top10 = None
        self.deals = None
        self.stats = None
        self.current_theme = tk.StringVar(value="Amazon Dark")
        self.T = THEMES["Amazon Dark"]
        self.fig = None  # current chart figure (white bg for PDF)

        self._build_ui()
        self.apply_theme()

    # ── UI BUILDER ────────────────────────────────────────────────────────

    def _build_ui(self):
        # Top header bar
        self.header = tk.Frame(self, height=60)
        self.header.pack(fill="x", side="top")
        self.header.pack_propagate(False)

        tk.Label(self.header, text="🛒 Amazon Product Analysis Dashboard",
                 font=("Segoe UI", 16, "bold")).pack(side="left", padx=16, pady=10)

        # Theme selector
        right_bar = tk.Frame(self.header)
        right_bar.pack(side="right", padx=12)
        tk.Label(right_bar, text="Theme:", font=("Segoe UI",9)).pack(side="left", padx=(0,4))
        self.theme_cb = ttk.Combobox(right_bar, textvariable=self.current_theme,
                                     values=list(THEMES.keys()), state="readonly", width=14)
        self.theme_cb.pack(side="left")
        self.theme_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_theme())

        # Toolbar
        self.toolbar = tk.Frame(self, height=48)
        self.toolbar.pack(fill="x", side="top")
        self.toolbar.pack_propagate(False)

        self.btn_load = self._tb_btn(self.toolbar, "📂  Load Excel", self.load_file)
        self.btn_pdf  = self._tb_btn(self.toolbar, "📄  Export PDF", self.export_pdf)
        self.btn_ref  = self._tb_btn(self.toolbar, "🔄  Refresh", self.refresh)
        self.status_var = tk.StringVar(value="No file loaded — click 'Load Excel' to begin")
        self.status_lbl = tk.Label(self.toolbar, textvariable=self.status_var,
                                   font=("Segoe UI",9), anchor="e")
        self.status_lbl.pack(side="right", padx=14)

        # Main pane: sidebar + content
        self.paned = tk.PanedWindow(self, orient="horizontal", sashwidth=4)
        self.paned.pack(fill="both", expand=True, padx=0, pady=0)

        self.sidebar = tk.Frame(self.paned, width=220)
        self.sidebar.pack_propagate(False)
        self.paned.add(self.sidebar, minsize=180)

        self.content = tk.Frame(self.paned)
        self.paned.add(self.content, minsize=600)

        self._build_sidebar()
        self._build_notebook()

        # Status bar bottom
        self.statusbar = tk.Frame(self, height=24)
        self.statusbar.pack(fill="x", side="bottom")
        self.statusbar.pack_propagate(False)
        self.sb_text = tk.Label(self.statusbar, text="Ready", font=("Segoe UI",8), anchor="w")
        self.sb_text.pack(side="left", padx=8)

    def _tb_btn(self, parent, text, cmd):
        b = tk.Button(parent, text=text, command=cmd, relief="flat",
                      font=("Segoe UI",10,"bold"), cursor="hand2",
                      bd=0, padx=14, pady=8)
        b.pack(side="left", padx=4, pady=6)
        return b

    def _build_sidebar(self):
        tk.Label(self.sidebar, text="📊 Analysis Panels",
                 font=("Segoe UI",11,"bold")).pack(pady=(16,8), padx=12, anchor="w")

        self.nav_buttons = {}
        pages = [
            ("🏠  Overview",     "overview"),
            ("🏆  Top 10 Products","top10"),
            ("💚  Decent Deals",  "deals"),
            ("📈  Charts",        "charts"),
            ("📋  All Products",  "all"),
        ]
        for label, key in pages:
            b = tk.Button(self.sidebar, text=label, anchor="w",
                          font=("Segoe UI",10), relief="flat", bd=0,
                          cursor="hand2", padx=12, pady=8,
                          command=lambda k=key: self.show_tab(k))
            b.pack(fill="x", padx=8, pady=2)
            self.nav_buttons[key] = b

        tk.Frame(self.sidebar, height=1).pack(fill="x", pady=12, padx=12)
        tk.Label(self.sidebar, text="File Info",
                 font=("Segoe UI",10,"bold")).pack(anchor="w", padx=12)
        self.file_info = tk.Label(self.sidebar, text="No file loaded",
                                  font=("Segoe UI",8), wraplength=190,
                                  justify="left", anchor="w")
        self.file_info.pack(anchor="w", padx=12, pady=4)

    def _build_notebook(self):
        self.pages = {}
        self.page_container = tk.Frame(self.content)
        self.page_container.pack(fill="both", expand=True)

        for key in ["overview","top10","deals","charts","all"]:
            f = tk.Frame(self.page_container)
            f.place(x=0, y=0, relwidth=1, relheight=1)
            self.pages[key] = f

        self._build_overview()
        self._build_top10()
        self._build_deals()
        self._build_charts()
        self._build_all()
        self.show_tab("overview")

    # ── PAGES ─────────────────────────────────────────────────────────────

    def _build_overview(self):
        p = self.pages["overview"]
        tk.Label(p, text="Dashboard Overview",
                 font=("Segoe UI",18,"bold")).pack(pady=(20,4), padx=20, anchor="w")
        tk.Label(p, text="Load an Excel file to see analysis results",
                 font=("Segoe UI",10)).pack(padx=20, anchor="w")

        self.kpi_frame = tk.Frame(p)
        self.kpi_frame.pack(fill="x", padx=20, pady=16)

        self.overview_chart_frame = tk.Frame(p)
        self.overview_chart_frame.pack(fill="both", expand=True, padx=20, pady=(0,16))

    def _build_kpis(self):
        for w in self.kpi_frame.winfo_children():
            w.destroy()
        T = self.T
        kpis = [
            ("Total Products", str(self.stats["Total Products"]), "📦", T["accent"]),
            ("Avg Rating", f"⭐ {self.stats['Avg Rating']}", "⭐", T["warning"]),
            ("Avg Price", f"₹{self.stats['Avg Price (₹)']}", "💰", T["accent2"]),
            ("Top Rated", str(self.stats["Top Rated"]), "🏅", T["success"]),
            ("Decent Deals", str(len(self.deals)), "💚", T["success"]),
            ("Price Range", f"₹{self.stats['Min Price (₹)']}–{self.stats['Max Price (₹)']}", "📊", T["accent"]),
        ]
        for i, (label, val, icon, color) in enumerate(kpis):
            card = tk.Frame(self.kpi_frame, bd=0, relief="flat",
                            bg=T["card"], padx=14, pady=12)
            card.grid(row=0, column=i, padx=6, sticky="ew")
            self.kpi_frame.columnconfigure(i, weight=1)
            tk.Label(card, text=icon, font=("Segoe UI",20), bg=T["card"],
                     fg=color).pack()
            tk.Label(card, text=val, font=("Segoe UI",13,"bold"),
                     bg=T["card"], fg=color).pack()
            tk.Label(card, text=label, font=("Segoe UI",8),
                     bg=T["card"], fg=T["text2"]).pack()

    def _build_top10(self):
        p = self.pages["top10"]
        tk.Label(p, text="🏆 Top 10 Products",
                 font=("Segoe UI",16,"bold")).pack(pady=(16,2), padx=16, anchor="w")
        tk.Label(p, text="Ranked by rating, reviews and value score",
                 font=("Segoe UI",9)).pack(padx=16, anchor="w", pady=(0,8))

        cols = ("Rank","Product Name","Rating","Price","Reviews","Link")
        self.tree_top = self._make_tree(p, cols, [40,380,70,80,80,80])

    def _build_deals(self):
        p = self.pages["deals"]
        tk.Label(p, text="💚 Decent Deals — Best Value",
                 font=("Segoe UI",16,"bold")).pack(pady=(16,2), padx=16, anchor="w")
        tk.Label(p, text="Mid-price products with rating ≥ 4.0 stars",
                 font=("Segoe UI",9)).pack(padx=16, anchor="w", pady=(0,8))

        cols = ("Rank","Product Name","Rating","Price","Reviews","Link")
        self.tree_deals = self._make_tree(p, cols, [40,380,70,80,80,80])

    def _build_all(self):
        p = self.pages["all"]
        hdr = tk.Frame(p)
        hdr.pack(fill="x", padx=16, pady=12)
        tk.Label(hdr, text="📋 All Products",
                 font=("Segoe UI",16,"bold")).pack(side="left")
        tk.Label(hdr, text="Search:", font=("Segoe UI",10)).pack(side="left", padx=(20,4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._filter_all())
        tk.Entry(hdr, textvariable=self.search_var, font=("Segoe UI",10),
                 width=28, relief="flat", bd=4).pack(side="left")

        cols = ("Product Name","Rating","Price","Reviews","Open","Link")
        self.tree_all = self._make_tree(p, cols, [420,70,80,80,100])
        self.tree_all.column("Link", width=0, stretch=False)
        self.tree_all.unbind("<Double-1>")   # remove default binding
        self.tree_all.bind("<Double-1>", self._open_link_all)

    def _build_charts(self):
        p = self.pages["charts"]
        tk.Label(p, text="📈 Data Visualizations",
                 font=("Segoe UI",16,"bold")).pack(pady=(16,2), padx=16, anchor="w")

        self.chart_container = tk.Frame(p)
        self.chart_container.pack(fill="both", expand=True, padx=8, pady=8)

    def _make_tree(self, parent, cols, widths):
        frame = tk.Frame(parent)
        frame.pack(fill="both", expand=True, padx=12, pady=4)

        style_name = f"Custom{id(frame)}.Treeview"
        style = ttk.Style()
        T = self.T
        style.theme_use("clam")
        style.configure(style_name,
            background=T["treeview_bg"], foreground=T["treeview_fg"],
            fieldbackground=T["treeview_bg"], rowheight=30,
            font=("Segoe UI",9))
        style.configure(f"{style_name}.Heading",
            background=T["card"], foreground=T["accent"],
            font=("Segoe UI",9,"bold"), relief="flat")
        style.map(style_name, background=[("selected",T["treeview_sel"])])

        vsb = tk.Scrollbar(frame, orient="vertical")
        hsb = tk.Scrollbar(frame, orient="horizontal")
        tree = ttk.Treeview(frame, columns=cols, show="headings",
                            style=style_name,
                            yscrollcommand=vsb.set,
                            xscrollcommand=hsb.set)
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)

        for col, w in zip(cols, widths):
            tree.heading(col, text=col,
                         command=lambda c=col, t=tree: self._sort_tree(t, c))
            tree.column(col, width=w, minwidth=40, anchor="center" if col!="Product Name" else "w")

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        tree.bind("<Double-1>", self._open_link)
        tree.bind("<Return>", self._open_link)
        return tree

    # ── DATA LOADING ──────────────────────────────────────────────────────

    def load_file(self, path=None):
        if path is None:
            path = filedialog.askopenfilename(
                title="Select Excel File",
                filetypes=[("Excel files","*.xlsx *.xls *.xlsm"),("All","*.*")])
        if not path:
            return
        try:
            self.sb_text.config(text="Loading…")
            self.update_idletasks()
            self.df    = load_and_clean(path)
            self.top10 = get_top10(self.df)
            self.deals = get_decent_deals(self.df)
            self.stats = summary_stats(self.df)
            fname = os.path.basename(path)
            self.status_var.set(f"✅ Loaded: {fname}  ({len(self.df)} products)")
            self.file_info.config(text=f"{fname}\n{len(self.df)} products\n"
                                       f"{self.df['Price_num'].nunique()} price points")
            self._populate_all()
            self._populate_top10()
            self._populate_deals()
            self._build_kpis()
            self._draw_overview_mini()
            self._draw_charts()
            self.apply_theme()
            self.sb_text.config(text=f"Loaded {len(self.df)} products from {fname}")
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            self.sb_text.config(text="Error loading file")

    def refresh(self):
        if self.df is not None:
            self.load_file.__func__(self)  # re-show same data (theme refresh)
            self.apply_theme()

    # ── POPULATE TREES ────────────────────────────────────────────────────

    def _populate_top10(self):
        self.tree_top.delete(*self.tree_top.get_children())
        for i, (_, r) in enumerate(self.top10.iterrows(), 1):
            link = str(r.get("Link",""))[:60] if pd.notna(r.get("Link")) else ""
            rev = f"{int(r['Reviews_num']):,}" if not pd.isna(r['Reviews_num']) else "N/A"
            self.tree_top.insert("", "end", iid=str(r.get("Link",i)),
                values=(i, r["Product Name"], f"⭐ {r['Rating_num']:.1f}",
                        f"₹{r['Price_num']:.0f}", rev, "🔗 Open"),
                tags=("link",))
        self.tree_top.tag_configure("link", foreground=self.T["accent2"])

    def _populate_deals(self):
        self.tree_deals.delete(*self.tree_deals.get_children())
        for i, (_, r) in enumerate(self.deals.iterrows(), 1):
            rev = f"{int(r['Reviews_num']):,}" if not pd.isna(r['Reviews_num']) else "N/A"
            self.tree_deals.insert("", "end", iid=f"d{i}{r.get('Link',i)}",
                values=(i, r["Product Name"], f"⭐ {r['Rating_num']:.1f}",
                        f"₹{r['Price_num']:.0f}", rev, "🔗 Open"),
                tags=("link",))
        self.tree_deals.tag_configure("link", foreground=self.T["success"])

    def _populate_all(self):
        self.tree_all.delete(*self.tree_all.get_children())
    
        for _, r in self.df.iterrows():
            rev = f"{int(r['Reviews_num']):,}" if not pd.isna(r['Reviews_num']) else "N/A"
    
            link = r.get("Link", "")
            if pd.isna(link):
                link = ""
            else:
                link = str(link).strip()
    
            self.tree_all.insert(
                "", "end",
                values=(
                    r["Product Name"],
                    f"⭐ {r['Rating_num']:.1f}",
                    f"₹{r['Price_num']:.0f}",
                    rev,
                    "🔗 Open",
                    link   # 👈 hidden data
                )
            )
        self.tree_all.tag_configure(link, foreground=self.T["danger"])   
            

    def _filter_all(self):
        q = self.search_var.get().lower()
        self.tree_all.delete(*self.tree_all.get_children())
        subset = self.df if self.df is not None else pd.DataFrame()
        for _, r in subset.iterrows():
            if q in str(r["Product Name"]).lower():
                rev = f"{int(r['Reviews_num']):,}" if not pd.isna(r['Reviews_num']) else "N/A"
                self.tree_all.insert("", "end",
                    values=(r["Product Name"], f"⭐ {r['Rating_num']:.1f}",
                            f"₹{r['Price_num']:.0f}", rev, "🔗 Open"))

    def _sort_tree(self, tree, col):
        data = [(tree.set(k, col), k) for k in tree.get_children("")]
        try:
            data.sort(key=lambda x: float(x[0].replace("₹","").replace("⭐","").replace(",","").strip()))
        except:
            data.sort(key=lambda x: x[0].lower())
        for i, (_, k) in enumerate(data):
            tree.move(k, "", i)

    def _open_link(self, event):
        widget = event.widget
        item = widget.focus()
        if not item:
            return
        vals = widget.item(item,"values")
        if not vals:
            return
        # Link is in the iid for top10/deals, last column shows open
        # For all products tree, link col is last
        # We stored actual link in iid for top10
        if self.df is None:
            return
        # Find matching row by product name
        name = vals[1] if len(vals)>1 else vals[0]
        match = self.df[self.df["Product Name"]==name]
        if match.empty:
            match = self.df[self.df["Product Name"].str.startswith(str(name)[:30])]
        if not match.empty:
            link = match.iloc[0].get("Link","")
            if pd.notna(link) and str(link).startswith("http"):
                webbrowser.open(str(link))
                
    def _open_link_all(self, event):
        tree = event.widget
        item = tree.identify_row(event.y)
    
        if not item:
            return
    
        values = tree.item(item, "values")
    
        if not values or len(values) < 6:
            print("No link data")
            return
    
        link = values[5]  # hidden column
    
        if link:
            if not str(link).startswith("http"):
                link = "https://" + str(link)
            webbrowser.open(link)
        else:
            print("Empty link")            

    # ── CHARTS ────────────────────────────────────────────────────────────

    def _draw_charts(self):
        for w in self.chart_container.winfo_children():
            w.destroy()
        if self.df is None:
            return
        T = self.T
        fig = Figure(figsize=(13, 8), facecolor=T["plot_bg"], tight_layout=True)
        gs = gridspec.GridSpec(2, 3, figure=fig, hspace=0.45, wspace=0.35)

        accent = T["accent"]
        accent2 = T["accent2"]
        fg = T["plot_fg"]
        grid_c = T["plot_grid"]
        success = T["success"]

        # 1. Price Distribution
        ax1 = fig.add_subplot(gs[0,0])
        prices = self.df["Price_num"].dropna()
        ax1.hist(prices, bins=20, color=accent, edgecolor=T["plot_bg"], alpha=0.85)
        ax1.set_title("Price Distribution", color=fg, fontsize=9, fontweight="bold")
        ax1.set_xlabel("Price (₹)", color=fg, fontsize=7)
        ax1.set_ylabel("Count", color=fg, fontsize=7)
        self._style_ax(ax1, T)

        # 2. Rating Distribution
        ax2 = fig.add_subplot(gs[0,1])
        ratings = self.df["Rating_num"].dropna()
        bins = np.arange(1,5.6,0.5)
        ax2.hist(ratings, bins=bins, color=accent2, edgecolor=T["plot_bg"], alpha=0.85)
        ax2.set_title("Rating Distribution", color=fg, fontsize=9, fontweight="bold")
        ax2.set_xlabel("Rating (Stars)", color=fg, fontsize=7)
        ax2.set_ylabel("Count", color=fg, fontsize=7)
        self._style_ax(ax2, T)

        # 3. Top 10 by Reviews (bar)
        ax3 = fig.add_subplot(gs[0,2])
        t10 = self.top10.nlargest(10,"Reviews_num")
        names = [n[:18]+"…" if len(n)>18 else n for n in t10["Product Name"]]
        bars = ax3.barh(range(len(names)), t10["Reviews_num"]/1000,
                        color=[success]*len(names), alpha=0.85)
        ax3.set_yticks(range(len(names)))
        ax3.set_yticklabels(names, fontsize=6, color=fg)
        ax3.set_title("Top 10 by Reviews (K)", color=fg, fontsize=9, fontweight="bold")
        ax3.set_xlabel("Reviews (thousands)", color=fg, fontsize=7)
        self._style_ax(ax3, T)

        # 4. Price vs Rating scatter
        ax4 = fig.add_subplot(gs[1,0])
        sc = ax4.scatter(self.df["Price_num"], self.df["Rating_num"],
                         alpha=0.6, c=self.df["Reviews_num"].fillna(0),
                         cmap="YlOrRd", s=25, edgecolors="none")
        fig.colorbar(sc, ax=ax4, label="Reviews").ax.yaxis.set_tick_params(color=fg)
        ax4.set_title("Price vs Rating", color=fg, fontsize=9, fontweight="bold")
        ax4.set_xlabel("Price (₹)", color=fg, fontsize=7)
        ax4.set_ylabel("Rating", color=fg, fontsize=7)
        self._style_ax(ax4, T)

        # 5. Top 10 ratings bar
        ax5 = fig.add_subplot(gs[1,1])
        t10r = self.df.nlargest(10,"Rating_num")
        names2 = [n[:16]+"…" if len(n)>16 else n for n in t10r["Product Name"]]
        bars2 = ax5.bar(range(len(names2)), t10r["Rating_num"],
                        color=accent, alpha=0.85, width=0.6)
        ax5.set_xticks(range(len(names2)))
        ax5.set_xticklabels(names2, rotation=45, ha="right", fontsize=5.5, color=fg)
        ax5.set_ylim(0,5.5)
        ax5.set_title("Top 10 Ratings", color=fg, fontsize=9, fontweight="bold")
        ax5.set_ylabel("Rating", color=fg, fontsize=7)
        self._style_ax(ax5, T)

        # 6. Pie: price segments
        ax6 = fig.add_subplot(gs[1,2])
        bins_p = [0,100,300,600,1000,99999]
        labs = ["<₹100","₹100-300","₹300-600","₹600-1K",">₹1K"]
        counts = pd.cut(self.df["Price_num"], bins=bins_p, labels=labs).value_counts()
        ax6.pie(counts, labels=counts.index, autopct="%1.0f%%",
                colors=[accent,accent2,success,T["warning"],T["danger"]],
                textprops={"color":fg,"fontsize":7}, startangle=140)
        ax6.set_title("Price Segments", color=fg, fontsize=9, fontweight="bold")

        canvas = FigureCanvasTkAgg(fig, master=self.chart_container)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Store white-bg version for PDF
        self.fig = fig

    def _draw_overview_mini(self):
        for w in self.overview_chart_frame.winfo_children():
            w.destroy()
        if self.df is None:
            return
        T = self.T
        fig = Figure(figsize=(10,3.5), facecolor=T["plot_bg"], tight_layout=True)
        gs = gridspec.GridSpec(1,3,figure=fig,wspace=0.35)
        fg = T["plot_fg"]

        ax1 = fig.add_subplot(gs[0,0])
        ax1.hist(self.df["Price_num"].dropna(), bins=18, color=T["accent"],
                 edgecolor=T["plot_bg"], alpha=0.85)
        ax1.set_title("Price Distribution", color=fg, fontsize=9, fontweight="bold")
        self._style_ax(ax1, T)

        ax2 = fig.add_subplot(gs[0,1])
        ax2.hist(self.df["Rating_num"].dropna(), bins=np.arange(1,5.6,0.5),
                 color=T["accent2"], edgecolor=T["plot_bg"], alpha=0.85)
        ax2.set_title("Ratings Spread", color=fg, fontsize=9, fontweight="bold")
        self._style_ax(ax2, T)

        ax3 = fig.add_subplot(gs[0,2])
        t = self.top10.head(8)
        names = [n[:14]+"…" if len(n)>14 else n for n in t["Product Name"]]
        ax3.barh(range(len(names)), t["Rating_num"], color=T["success"], alpha=0.85)
        ax3.set_yticks(range(len(names)))
        ax3.set_yticklabels(names, fontsize=6, color=fg)
        ax3.set_xlim(0,5.5)
        ax3.set_title("Top Products Rating", color=fg, fontsize=9, fontweight="bold")
        self._style_ax(ax3, T)

        canvas = FigureCanvasTkAgg(fig, master=self.overview_chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _style_ax(self, ax, T):
        ax.set_facecolor(T["plot_bg"])
        ax.tick_params(colors=T["plot_fg"], labelsize=6)
        ax.spines[:].set_color(T["plot_grid"])
        for label in ax.get_xticklabels() + ax.get_yticklabels():
            label.set_color(T["plot_fg"])
        ax.grid(True, color=T["plot_grid"], alpha=0.3, linewidth=0.5)

    # ── NAVIGATION ────────────────────────────────────────────────────────

    def show_tab(self, key):
        for k, p in self.pages.items():
            p.lower()
        self.pages[key].lift()
        T = self.T
        for k, b in self.nav_buttons.items():
            if k == key:
                b.config(bg=T["accent"], fg=T["button_fg"])
            else:
                b.config(bg=T["sidebar"], fg=T["text"])

    # ── THEME ─────────────────────────────────────────────────────────────

    def apply_theme(self):
        name = self.current_theme.get()
        self.T = THEMES.get(name, THEMES["Amazon Dark"])
        T = self.T
        self.configure(bg=T["bg"])

        def recolor(w):
            try:
                cls = w.winfo_class()
                if cls in ("Frame","Labelframe"):
                    w.configure(bg=T["bg"])
                elif cls == "Label":
                    try: w.configure(bg=T["bg"], fg=T["text"])
                    except: pass
                elif cls == "Button":
                    w.configure(bg=T["button_bg"], fg=T["button_fg"],
                                activebackground=T["accent"], activeforeground=T["button_fg"])
                elif cls == "Entry":
                    w.configure(bg=T["surface"], fg=T["text"], insertbackground=T["text"])
                for child in w.winfo_children():
                    recolor(child)
            except: pass

        recolor(self)
        # Special frames
        self.header.configure(bg=T["header_bg"])
        self.toolbar.configure(bg=T["surface"])
        self.sidebar.configure(bg=T["sidebar"])
        self.statusbar.configure(bg=T["surface"])
        self.sb_text.configure(bg=T["surface"], fg=T["text2"])
        self.status_lbl.configure(bg=T["surface"], fg=T["text2"])

        for label in self.header.winfo_children():
            if isinstance(label, tk.Label):
                label.configure(bg=T["header_bg"], fg=T["accent"])

        # Nav buttons
        for k, b in self.nav_buttons.items():
            b.configure(bg=T["sidebar"], fg=T["text"],
                        activebackground=T["accent"], activeforeground=T["button_fg"])

        # Toolbar buttons
        for b in [self.btn_load, self.btn_pdf, self.btn_ref]:
            b.configure(bg=T["accent"], fg=T["button_fg"])

        # Page frames
        for f in self.pages.values():
            f.configure(bg=T["bg"])
            for w in f.winfo_children():
                try:
                    w.configure(bg=T["bg"])
                    if isinstance(w, tk.Label):
                        w.configure(fg=T["text"])
                except: pass

        # KPIs
        if hasattr(self,'kpi_frame'):
            for card in self.kpi_frame.winfo_children():
                card.configure(bg=T["card"])
                for lbl in card.winfo_children():
                    try: lbl.configure(bg=T["card"])
                    except: pass

        # Redraw charts with new theme
        if self.df is not None:
            self._draw_charts()
            self._draw_overview_mini()
            self._build_kpis()

        self.show_tab(self.show_tab.__defaults__[0] if hasattr(self.show_tab,'__defaults__') and self.show_tab.__defaults__ else "overview")

    # ── EXPORT PDF ────────────────────────────────────────────────────────

    def export_pdf(self):
        if self.df is None:
            messagebox.showwarning("No Data","Please load an Excel file first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF","*.pdf")],
            initialfile="amazon_analysis_report.pdf")
        if not path:
            return

        # Generate white-bg fig for PDF
        T_white = {
            "plot_bg":"#FFFFFF","plot_fg":"#111111","plot_grid":"#DDDDDD",
            "accent":"#FF9900","accent2":"#146EB4","success":"#067D62",
            "warning":"#FF9900","danger":"#CC0C39"
        }
        fig_pdf = Figure(figsize=(13,8), facecolor="white", tight_layout=True)
        gs = gridspec.GridSpec(2,3,figure=fig_pdf,hspace=0.45,wspace=0.35)
        self._draw_pdf_charts(fig_pdf, gs, T_white)

        self.sb_text.config(text="Generating PDF…")
        self.update_idletasks()
        try:
            ok = generate_pdf(self.df, self.top10, self.deals, self.stats, fig_pdf, path)
            if ok:
                messagebox.showinfo("Success",f"PDF saved to:\n{path}")
                self.sb_text.config(text=f"PDF saved: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("PDF Error", str(e))
            self.sb_text.config(text="PDF generation failed")

    def _draw_pdf_charts(self, fig, gs, T):
        fg = T["plot_fg"]
        ax1 = fig.add_subplot(gs[0,0])
        ax1.hist(self.df["Price_num"].dropna(), bins=20, color=T["accent"], alpha=0.8)
        ax1.set_title("Price Distribution", fontsize=9, fontweight="bold")
        ax1.set_xlabel("Price (₹)", fontsize=7); ax1.set_ylabel("Count", fontsize=7)

        ax2 = fig.add_subplot(gs[0,1])
        ax2.hist(self.df["Rating_num"].dropna(), bins=np.arange(1,5.6,0.5),
                 color=T["accent2"], alpha=0.8)
        ax2.set_title("Rating Distribution", fontsize=9, fontweight="bold")

        ax3 = fig.add_subplot(gs[0,2])
        t10 = self.top10.nlargest(10,"Reviews_num")
        names = [n[:18]+"…" if len(n)>18 else n for n in t10["Product Name"]]
        ax3.barh(range(len(names)), t10["Reviews_num"]/1000, color=T["success"], alpha=0.8)
        ax3.set_yticks(range(len(names))); ax3.set_yticklabels(names, fontsize=6)
        ax3.set_title("Top 10 by Reviews (K)", fontsize=9, fontweight="bold")

        ax4 = fig.add_subplot(gs[1,0])
        ax4.scatter(self.df["Price_num"], self.df["Rating_num"],
                    alpha=0.5, c=self.df["Reviews_num"].fillna(0), cmap="YlOrRd", s=20)
        ax4.set_title("Price vs Rating", fontsize=9, fontweight="bold")
        ax4.set_xlabel("Price (₹)", fontsize=7); ax4.set_ylabel("Rating", fontsize=7)

        ax5 = fig.add_subplot(gs[1,1])
        t10r = self.df.nlargest(10,"Rating_num")
        names2 = [n[:14]+"…" if len(n)>14 else n for n in t10r["Product Name"]]
        ax5.bar(range(len(names2)), t10r["Rating_num"], color=T["accent"], alpha=0.8, width=0.6)
        ax5.set_xticks(range(len(names2)))
        ax5.set_xticklabels(names2, rotation=45, ha="right", fontsize=5.5)
        ax5.set_ylim(0,5.5); ax5.set_title("Top 10 Ratings", fontsize=9, fontweight="bold")

        ax6 = fig.add_subplot(gs[1,2])
        bins_p = [0,100,300,600,1000,99999]
        labs = ["<₹100","₹100-300","₹300-600","₹600-1K",">₹1K"]
        counts = pd.cut(self.df["Price_num"], bins=bins_p, labels=labs).value_counts()
        ax6.pie(counts, labels=counts.index, autopct="%1.0f%%",
                colors=[T["accent"],T["accent2"],T["success"],T["warning"],T["danger"]],
                textprops={"fontsize":7}, startangle=140)
        ax6.set_title("Price Segments", fontsize=9, fontweight="bold")


# ═══════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import sys
    app = AmazonDashboard()

    # Auto-load sample if path given as arg or hardcoded for demo
    sample = None
    if len(sys.argv) > 1:
        sample = sys.argv[1]
    elif os.path.exists("/mnt/user-data/uploads/sample1.xlsx"):
        sample = "/mnt/user-data/uploads/sample1.xlsx"

    if sample and os.path.exists(sample):
        app.after(200, lambda: app.load_file(sample))

    app.mainloop()