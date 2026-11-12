"""GradAspal - Perencanaan Gradasi Agregat, Kadar Aspal Estimasi & Kadar Aspal Optimum (Marshall)
Latasir, Lataston, Laston.  Jalankan: streamlit run app.py
"""
import csv
import hashlib
import io
import json
import os
import re
from datetime import datetime, timedelta, timezone

import importlib.util as _ilu

import numpy as np
import pandas as pd
import streamlit as st

# Pengaman: jika paket tambahan belum terpasang, tampilkan petunjuk yang jelas (bukan traceback merah)
_BELUM = [m for m in ("plotly", "openpyxl") if _ilu.find_spec(m) is None]
if _BELUM:
    st.set_page_config(page_title="GradAspal", page_icon="🛣️")
    st.error("Aplikasi belum bisa dijalankan: paket **" + ", ".join(_BELUM) + "** belum terpasang di server.")
    st.markdown("**Cara memperbaiki:** buka repo GitHub aplikasi ini, lalu isi file `requirements.txt` "
                "(di folder yang sama dengan `app.py`) dengan baris berikut. Setelah itu klik **Commit changes**, "
                "buka **Manage app**, pilih **Reboot app**, dan tunggu 1-3 menit.")
    st.code("streamlit\npandas\nnumpy\nopenpyxl\nplotly", language="text")
    st.stop()

import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from plotly.subplots import make_subplots

# Saringan spesifikasi (mm) dan batas gradasi % lolos (BB, BA)
S = [37.5, 25, 19, 12.5, 9.5, 4.75, 2.36, 1.18, 0.6, 0.3, 0.15, 0.075]
SPEC = {
    "Latasir Kelas A": {19: (100, 100), 9.5: (90, 100), 0.075: (10, 15)},
    "Latasir Kelas B": {19: (100, 100), 2.36: (75, 100), 0.075: (8, 13)},
    "Lataston WC (Senjang)": {19: (100, 100), 12.5: (90, 100), 9.5: (75, 85), 2.36: (50, 72), 0.6: (35, 60), 0.075: (6, 10)},
    "Lataston Base (Senjang)": {19: (100, 100), 12.5: (90, 100), 9.5: (65, 90), 2.36: (35, 55), 0.6: (15, 35), 0.075: (2, 9)},
    "Lataston WC (Semi Senjang)": {19: (100, 100), 12.5: (87, 100), 9.5: (55, 88), 2.36: (50, 62), 0.6: (20, 45), 0.3: (15, 35), 0.075: (6, 10)},
    "Lataston Base (Semi Senjang)": {19: (100, 100), 12.5: (90, 100), 9.5: (55, 70), 2.36: (32, 44), 0.6: (15, 35), 0.3: (5, 35), 0.075: (4, 8)},
    "Laston AC-WC": {19: (100, 100), 12.5: (90, 100), 9.5: (77, 90), 4.75: (53, 69), 2.36: (33, 53), 1.18: (21, 40), 0.6: (14, 30), 0.3: (9, 22), 0.15: (6, 15), 0.075: (4, 9)},
    "Laston AC-BC": {25: (100, 100), 19: (90, 100), 12.5: (75, 90), 9.5: (66, 82), 4.75: (46, 64), 2.36: (30, 49), 1.18: (18, 38), 0.6: (12, 28), 0.3: (7, 20), 0.15: (5, 13), 0.075: (4, 8)},
    "Laston AC-Base": {37.5: (100, 100), 25: (90, 100), 19: (76, 90), 12.5: (60, 78), 9.5: (52, 71), 4.75: (35, 54), 2.36: (23, 41), 1.18: (13, 30), 0.6: (10, 22), 0.3: (6, 15), 0.15: (4, 10), 0.075: (3, 7)},
}
K_DEF = {"Latasir": 2.5, "Lataston": 2.5, "Laston": 0.75}
K_RNG = {"Latasir": (2.0, 3.0), "Lataston": (2.0, 3.0), "Laston": (0.5, 1.0)}

# Syarat Marshall Laston (Spesifikasi Umum 2018, Tabel 12.2): (min, maks); None = tidak dibatasi.
# MQ tidak ada di Tabel 12.2, jadi tidak dibatasi (boleh diisi di aplikasi jika modul mensyaratkan, mis. min 250).
_MQ = (None, None)
MSPEC = {
    "Laston AC-WC": {"VIM (%)": (3.0, 5.0), "VMA (%)": (15.0, None), "VFA (%)": (65.0, None),
                     "Stabilitas (kg)": (800.0, None), "Flow (mm)": (2.0, 4.0), "MQ (kg/mm)": _MQ},
    "Laston AC-BC": {"VIM (%)": (3.0, 5.0), "VMA (%)": (14.0, None), "VFA (%)": (65.0, None),
                     "Stabilitas (kg)": (800.0, None), "Flow (mm)": (2.0, 4.0), "MQ (kg/mm)": _MQ},
    "Laston AC-Base": {"VIM (%)": (3.0, 5.0), "VMA (%)": (13.0, None), "VFA (%)": (65.0, None),
                       "Stabilitas (kg)": (1800.0, None), "Flow (mm)": (3.0, 6.0), "MQ (kg/mm)": _MQ},
}
MPARAMS = ["VIM (%)", "VMA (%)", "VFA (%)", "Stabilitas (kg)", "Flow (mm)", "MQ (kg/mm)"]

# Angka korelasi beban stabilitas (SNI 06-2489-1991): (isi min, isi maks cm3, angka korelasi)
KORELASI = [(200, 213, 5.56), (214, 225, 5.00), (226, 237, 4.55), (238, 250, 4.17), (251, 264, 3.85),
            (265, 276, 3.57), (277, 289, 3.33), (290, 301, 3.03), (302, 316, 2.78), (317, 328, 2.50),
            (329, 340, 2.27), (341, 353, 2.08), (354, 367, 1.92), (368, 379, 1.79), (380, 392, 1.67),
            (393, 405, 1.56), (406, 420, 1.47), (421, 431, 1.39), (432, 443, 1.32), (444, 456, 1.25),
            (457, 470, 1.19), (471, 482, 1.14), (483, 495, 1.09), (496, 508, 1.04), (509, 522, 1.00),
            (523, 535, 0.96), (536, 546, 0.93), (547, 559, 0.89), (560, 573, 0.86), (574, 585, 0.83),
            (586, 598, 0.81), (599, 610, 0.78), (611, 625, 0.76)]


def f_korelasi(vol):
    """Angka korelasi dengan interpolasi linear antar titik tengah rentang isi benda uji."""
    xs = [(a + b) / 2 for a, b, _ in KORELASI]
    ys = [c for _, _, c in KORELASI]
    return round(float(np.interp(vol, xs, ys)), 2)


# Angka korelasi berdasarkan tebal benda uji (mm), SNI 06-2489-1991 (urutan sama dengan KORELASI)
KORELASI_TEBAL = [25.4, 27.0, 28.6, 30.2, 31.8, 33.3, 34.9, 36.5, 38.1, 39.7, 41.3, 42.9, 44.4, 46.0, 47.6, 49.2, 50.8,
                  52.4, 54.0, 55.6, 57.2, 58.7, 60.3, 61.9, 63.5, 65.1, 66.7, 68.3, 69.8, 71.4, 73.0, 74.6, 76.2]
assert len(KORELASI_TEBAL) == len(KORELASI)


def f_korelasi_tebal(tebal):
    """Angka korelasi dari tebal benda uji (mm), interpolasi linear antar baris tabel SNI."""
    return round(float(np.interp(tebal, KORELASI_TEBAL, [c for _, _, c in KORELASI])), 2)


def tabel_lembar_kerja(rm):
    """Hasil Marshall dalam format lembar kerja praktikum: tiap benda uji + baris rata-rata per kadar aspal."""
    kol = [("Kadar Aspal (%)", "Kadar Aspal (%)"), ("Berat di udara (F) gr", "Berat Kering (gr)"),
           ("Berat dalam air (G) gr", "Berat dalam Air (gr)"), ("Berat SSD (H) gr", "Berat SSD (gr)"),
           ("Volume (I = H - G) cm3", "Volume (cm3)"), ("BJ Bulk Gmb (J = F / I)", "BJ Bulk Campuran"),
           ("VIM (K) %", "VIM (%)"), ("VMA (L) %", "VMA (%)"), ("VFB (M) %", "VFA (%)"),
           ("Stabilitas bacaan (N)", "Pembacaan"), ("Tebal (mm)", "Tebal (mm)"), ("Angka korelasi", "Angka Korelasi"),
           ("Stabilitas terkoreksi (O) kg", "Stabilitas (kg)"), ("Flow (P) mm", "Flow (mm)"), ("MQ (Q = O / P) kg/mm", "MQ (kg/mm)")]
    out = []
    for kd, g in rm.sort_values(["Kadar Aspal (%)", "No"]).groupby("Kadar Aspal (%)", sort=True):
        for _, r in g.iterrows():
            row = {"Benda uji": str(int(r["No"]))}
            for lab, src in kol:
                row[lab] = float(r[src]) if src in rm.columns and not pd.isna(r.get(src)) else np.nan
            out.append(row)
        mrow = {"Benda uji": "rata-rata"}
        for lab, src in kol:
            mrow[lab] = float(g[src].mean()) if src in g.columns and g[src].notna().any() else np.nan
        mrow["Kadar Aspal (%)"] = float(kd)
        out.append(mrow)
    return pd.DataFrame(out, columns=["Benda uji"] + [k for k, _ in kol])


# Saringan hasil praktikum (mm); baris terakhir tabel input = pan
LAB = [76.2, 63.5, 50.8, 37.5, 19, 12.5, 9.5, 4.75, 2.36, 1.18, 0.6, 0.3, 0.15, 0.075]


# ------------------------------------------------ penyimpanan & riwayat akses
DATA_DIR = os.environ.get("GRADASPAL_DATA", "data")
ADMIN_CODE = os.environ.get("GRADASPAL_ADMIN", "gradaspal-admin")
try:  # di Streamlit Community Cloud, kode admin bisa disimpan di Secrets
    ADMIN_CODE = st.secrets["GRADASPAL_ADMIN"]
except Exception:
    pass
os.makedirs(os.path.join(DATA_DIR, "users"), exist_ok=True)
VISITS = os.path.join(DATA_DIR, "visits.csv")


def now_wib():
    return datetime.now(timezone(timedelta(hours=7))).strftime("%Y-%m-%d %H:%M:%S")


def user_path(uid):
    return os.path.join(DATA_DIR, "users", f"{uid}.json")


def load_user(uid):
    try:
        with open(user_path(uid), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_user(uid, data):
    try:
        with open(user_path(uid), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception:
        pass


def log_visit(name, event):
    try:
        ip = st.context.headers.get("X-Forwarded-For", "").split(",")[0].strip()
    except Exception:
        ip = ""
    try:
        new = not os.path.exists(VISITS)
        with open(VISITS, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["waktu (WIB)", "nama", "aktivitas", "ip"])
            w.writerow([now_wib(), name, event, ip])
    except Exception:
        pass


def saran(sz, v, lo, hi):
    """Saran perbaikan untuk satu saringan yang di luar batas."""
    if lo == hi == 100:
        return (f"**{sz:g} mm**: masih ada butir tertahan (lolos {v:.2f}%, seharusnya 100%). "
                "Ukuran butir maksimum melebihi spesifikasi. Ayak dan buang butir yang terlalu besar, atau pilih jenis campuran dengan ukuran maksimum lebih besar.")
    if v > hi:
        d = v - hi
        if sz == 0.075:
            return (f"**0,075 mm (filler)**: lolos {v:.2f}%, melebihi BA {hi:g}% sebesar {d:.2f}%. "
                    "Filler berlebih. Kurangi filler/abu batu atau kurangi debu pada agregat halus.")
        return (f"**{sz:g} mm**: lolos {v:.2f}%, melebihi BA {hi:g}% sebesar {d:.2f}%. "
                f"Terlalu banyak butir halus. Kurangi fraksi halus (pasir/abu batu) dan tambah fraksi kasar yang tertahan di saringan {sz:g} mm.")
    d = lo - v
    if sz == 0.075:
        return (f"**0,075 mm (filler)**: lolos {v:.2f}%, kurang dari BB {lo:g}% sebesar {d:.2f}%. "
                "Filler kurang. Tambahkan filler (semen/abu batu).")
    return (f"**{sz:g} mm**: lolos {v:.2f}%, kurang dari BB {lo:g}% sebesar {d:.2f}%. "
            f"Terlalu sedikit butir halus. Tambah fraksi halus yang lolos saringan {sz:g} mm, atau kurangi fraksi kasar.")


C_BB, C_BA, C_MID, C_SMP = "#1F3A5F", "#6FA8DC", "#7A8CA3", "#F97316"
BG_PAPER, BG_PLOT = "#FFFFFF", "#F8FBFE"


def plot(mix, xs=None, ys=None, mids=False):
    sp = SPEC[mix]
    x = [s for s in S if s in sp]
    lo_v, hi_v = [sp[s][0] for s in x], [sp[s][1] for s in x]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=x + x[::-1], y=hi_v + lo_v[::-1], fill="toself", mode="none",
                             fillcolor="rgba(111,168,220,0.22)", name="Daerah spesifikasi", hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=x, y=lo_v, mode="lines+markers", name="BB", line=dict(color=C_BB, width=2.5),
                             marker=dict(size=8), hovertemplate="BB<br>%{x:g} mm: %{y:g}%<extra></extra>"))
    fig.add_trace(go.Scatter(x=x, y=hi_v, mode="lines+markers", name="BA", line=dict(color=C_BA, width=2.5),
                             marker=dict(size=8), hovertemplate="BA<br>%{x:g} mm: %{y:g}%<extra></extra>"))
    if mids:
        fig.add_trace(go.Scatter(x=x, y=[sum(sp[s]) / 2 for s in x], mode="lines+markers", name="Nilai Tengah",
                                 line=dict(color=C_MID, width=2.5, dash="dash"),
                                 marker=dict(size=9, symbol="diamond"),
                                 hovertemplate="Nilai tengah<br>%{x:g} mm: %{y:.2f}%<extra></extra>"))
    if xs is not None:
        fig.add_trace(go.Scatter(x=list(xs), y=[float(v) for v in ys], mode="lines+markers",
                                 name="Presentase Butir Lolos (%)", line=dict(color=C_SMP, width=3),
                                 marker=dict(size=8),
                                 hovertemplate="Sampel<br>%{x:g} mm: %{y:.2f}%<extra></extra>"))
    axis = dict(showline=True, mirror=True, linecolor="#222", gridcolor="rgba(0,0,0,0.12)", zeroline=False)
    fig.update_xaxes(type="log", range=[2, -2], tickvals=[100, 10, 1, 0.1, 0.01],
                     ticktext=["100", "10", "1", "0,1", "0,01"], title="Ukuran Lubang Saringan (mm)",
                     minor=dict(showgrid=False), **axis)
    fig.update_yaxes(range=[0, 120], tickvals=list(range(0, 121, 20)),
                     title="Presentase Lolos Kumulatif (%)", **axis)
    fig.update_layout(
        title=dict(text=f"Hasil Perbandingan Presentase Lolos Kumulatif Terhadap Persen<br>"
                        f"Berat Butir yang Lolos pada {mix}", font=dict(size=14), x=0.5, xanchor="center"),
        paper_bgcolor=BG_PAPER, plot_bgcolor=BG_PLOT, height=520, margin=dict(l=60, r=20, t=80, b=20),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5),
        hovermode="closest")
    return fig


def plot_marshall(ra, specs, kao=None):
    params = ["Stabilitas (kg)", "Flow (mm)", "MQ (kg/mm)", "VMA (%)", "VIM (%)", "VFA (%)"]
    fig = make_subplots(rows=2, cols=3, subplot_titles=params, vertical_spacing=0.16, horizontal_spacing=0.07)
    for i, p in enumerate(params):
        r, c = i // 3 + 1, i % 3 + 1
        d = ra[["Kadar Aspal (%)", p]].dropna()
        if not d.empty:
            fig.add_trace(go.Scatter(x=d["Kadar Aspal (%)"], y=d[p], mode="lines+markers", name="Hasil uji",
                                     legendgroup="uji", showlegend=(i == 0),
                                     line=dict(color=C_SMP, width=2.5), marker=dict(size=8),
                                     hovertemplate="Kadar aspal %{x:g}%<br>" + p + ": %{y:.2f}<extra></extra>"),
                          row=r, col=c)
        lo, hi = specs.get(p, (np.nan, np.nan))
        if not pd.isna(lo):
            fig.add_hline(y=lo, line=dict(color=C_BB, dash="dash", width=1.5), row=r, col=c,
                          annotation_text=f"min {lo:g}", annotation_font_size=10)
        if not pd.isna(hi):
            fig.add_hline(y=hi, line=dict(color=C_BA, dash="dash", width=1.5), row=r, col=c,
                          annotation_text=f"maks {hi:g}", annotation_font_size=10)
        if kao is not None:
            fig.add_vline(x=kao, line=dict(color="#0E9F6E", dash="dot", width=2), row=r, col=c,
                          annotation_text="KAO", annotation_font_size=10)
        fig.update_xaxes(title_text="Kadar Aspal (%)", row=r, col=c)
    axis = dict(showline=True, mirror=True, linecolor="#999", gridcolor="rgba(0,0,0,0.12)")
    fig.update_xaxes(**axis)
    fig.update_yaxes(**axis)
    fig.update_layout(paper_bgcolor=BG_PAPER, plot_bgcolor=BG_PLOT, height=640,
                      margin=dict(l=50, r=20, t=50, b=40),
                      legend=dict(orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5))
    return fig


def show_fig(fig, name):
    st.plotly_chart(fig, width="stretch", key=f"pl_{name}",
                    config={"displaylogo": False, "responsive": True,
                            "modeBarButtonsToRemove": ["lasso2d", "select2d"],
                            "toImageButtonOptions": {"format": "png", "filename": name, "scale": 3}})
    st.caption("Arahkan kursor ke titik untuk melihat nilai. Klik legenda untuk menyembunyikan garis. "
               "Klik ikon kamera di pojok grafik untuk mengunduh PNG. Klik dua kali untuk reset zoom.")


# ------------------------------------------------ ekspor Excel (rumus hidup + grafik)
_HDR_FILL = PatternFill("solid", fgColor="243B53")
_HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF")
_IN_FILL = PatternFill("solid", fgColor="FFF2CC")
_IN_FONT = Font(name="Arial", color="0000FF")
_Q = {"G": "'Gradasi Rencana'", "A": "'Analisis Saringan'", "K": "'Kadar Aspal'", "B": "'Berat per Saringan'",
      "V": "'Variasi Kadar Aspal'", "T": "'Tabel Korelasi'", "P": "'Per Benda Uji'", "R": "'Rata-rata'",
      "C": "'Pengecekan Syarat'", "I": "Input"}


def _hdr(ws, row, ncol, c0=1):
    for c in range(c0, c0 + ncol):
        cell = ws.cell(row, c)
        cell.fill, cell.font = _HDR_FILL, _HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _inp(cell):
    cell.fill, cell.font = _IN_FILL, _IN_FONT


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[CL(i)].width = w


def _fmt(ws, rng, f):
    for row in ws[rng]:
        for c in row:
            c.number_format = f


def _num(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return float(v)


def _legend(ws, cell, text="Sel berwarna kuning dengan huruf biru = isian yang boleh diubah. Sel lain berisi rumus."):
    ws[cell] = text
    ws[cell].font = Font(name="Arial", italic=True, color="7F7F7F", size=9)


def _ser(xref, yref, title, color, dash=None, marker="circle", width=28575):
    s = Series(yref, xref, title=title)
    s.smooth = False
    s.graphicalProperties.line.solidFill = color
    s.graphicalProperties.line.width = width
    if dash:
        s.graphicalProperties.line.dashStyle = dash
    if marker == "none":
        s.marker = Marker(symbol="none")
    else:
        s.marker = Marker(symbol=marker, size=7)
        s.marker.graphicalProperties = GraphicalProperties(solidFill=color)
        s.marker.graphicalProperties.line.solidFill = color
    return s


def _scatter(title, xt, yt, w, h):
    ch = ScatterChart()
    ch.scatterStyle = "lineMarker"
    ch.title = title
    ch.x_axis.title, ch.y_axis.title = xt, yt
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.majorGridlines = ChartLines()
    ch.legend.position = "b"
    ch.width, ch.height = w, h
    return ch


def _grad_chart(mix, wsg, n, ws_s=None, mid=False):
    """Grafik gradasi: BB, BA (dan nilai tengah / sampel) pada sumbu x logaritmik terbalik."""
    ch = _scatter(f"Hasil Perbandingan Presentase Lolos Kumulatif Terhadap Persen Berat Butir yang Lolos pada {mix}",
                  "Ukuran Lubang Saringan (mm)", "Presentase Lolos Kumulatif (%)", 24, 13)
    ch.x_axis.scaling.logBase = 10
    ch.x_axis.scaling.min, ch.x_axis.scaling.max = 0.01, 100
    ch.x_axis.scaling.orientation = "maxMin"
    ch.y_axis.scaling.min, ch.y_axis.scaling.max, ch.y_axis.majorUnit = 0, 120, 20
    ch.y_axis.crosses = "max"  # sumbu y tetap di kiri walau sumbu x dibalik
    xr = Reference(wsg, min_col=7, min_row=2, max_row=n + 1)
    ch.series.append(_ser(xr, Reference(wsg, min_col=8, min_row=2, max_row=n + 1), "BB", "1F3A5F"))
    ch.series.append(_ser(xr, Reference(wsg, min_col=9, min_row=2, max_row=n + 1), "BA", "6FA8DC"))
    if mid:
        ch.series.append(_ser(xr, Reference(wsg, min_col=10, min_row=2, max_row=n + 1), "Nilai Tengah", "7A8CA3",
                              dash="dash", marker="diamond"))
    if ws_s is not None:
        ch.series.append(_ser(Reference(ws_s, min_col=1, min_row=2, max_row=len(LAB) + 1),
                              Reference(ws_s, min_col=5, min_row=2, max_row=len(LAB) + 1),
                              "Presentase Butir Lolos (%)", "F97316", width=34925))
    return ch


# ------------------------------------------------------------------ berat jenis dari berat hasil pengujian
def bj_kasar(bk, bj, ba):
    """Agregat kasar (SNI 1969): Bk kering oven, Bj SSD di udara, Ba SSD dalam air."""
    if not (bk and bj and ba):
        return None
    d = bj - ba
    if d <= 0 or bk - ba <= 0:
        return None
    return {"bulk": bk / d, "ssd": bj / d, "semu": bk / (bk - ba), "serap": (bj - bk) / bk * 100}


def bj_halus(bk, b, bt, s):
    """Agregat halus (SNI 1970): Bk kering oven, B piknometer+air, Bt piknometer+benda uji+air, S benda uji SSD."""
    if not (bk and b and bt and s):
        return None
    d, d2 = b + s - bt, b + bk - bt
    if d <= 0 or d2 <= 0:
        return None
    return {"bulk": bk / d, "ssd": s / d, "semu": bk / d2, "serap": (s - bk) / bk * 100}


def bj_filler(bk, b, bt):
    """Filler (piknometer): satu nilai berat jenis, dipakai untuk bulk dan semu."""
    if not (bk and b and bt):
        return None
    d = b + bk - bt
    return bk / d if d > 0 else None


def bj_aspal(a, b, c, d):
    """Aspal (SNI 2441): A piknometer kosong, B +air, C +aspal, D +aspal+air."""
    if not (a and b and c and d):
        return None
    den = (b - a) - (d - c)
    return (c - a) / den if den > 0 and c > a else None



def xlsx_gradasi(mix, W, K, kadar, Pbr, m, kadar_list, labels):
    """Workbook Tab 1-3: semua angka turunan berupa rumus yang saling terhubung + grafik Excel."""
    sp = SPEC[mix]
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Analisis Saringan"
    wsc = wb.create_sheet("Pengecekan Batas")
    wsg = wb.create_sheet("Gradasi Rencana")
    wsk = wb.create_sheet("Kadar Aspal")
    wsb = wb.create_sheet("Berat per Saringan")
    wsr = wb.create_sheet("Koreksi")
    wsv = wb.create_sheet("Variasi Kadar Aspal")
    wsw = wb.create_sheet("Variasi Berat Saringan")
    nL = len(LAB)
    rpan, rjum = nL + 2, nL + 3

    # ---- Analisis Saringan (input: massa tertahan)
    ws1.append(["Saringan (mm)", "Massa Tertahan (gr)", "Jumlah Tertahan (gr)", "% Kumulatif Tertahan", "% Lewat"])
    _hdr(ws1, 1, 5)
    for i, sz in enumerate(LAB + ["Pan"]):
        r = i + 2
        ws1.cell(r, 1, sz)
        mv = float(m[i])
        _inp(ws1.cell(r, 2, mv if mv > 0 else None))
        ws1.cell(r, 3, f"=SUM($B$2:B{r})")
        ws1.cell(r, 4, f'=IF($B${rjum}=0,"",C{r}/$B${rjum}*100)')
        ws1.cell(r, 5, f'=IF($B${rjum}=0,"",100-D{r})')
    ws1.cell(rjum, 1, "Jumlah")
    ws1.cell(rjum, 2, f"=SUM(B2:B{rpan})")
    ws1.cell(rjum, 3, f"=SUM(C2:C{rpan})")
    ws1.cell(rjum, 4, f'=IF($B${rjum}=0,"",SUM(D2:D{rpan - 1}))')
    for c in range(1, 6):
        ws1.cell(rjum, c).font = Font(name="Arial", bold=True)
    ws1["G1"], ws1["H1"] = "MHB", f'=IF($B${rjum}=0,"",D{rjum}/100)'
    ws1["G2"] = "Keterangan"
    ws1["H2"] = "MHB campuran gabungan (informasi). Syarat pasir 1,5-3,8 dan kerikil 5,0-8,0 berlaku untuk agregat tunggal."
    ws1["G1"].font = ws1["G2"].font = Font(name="Arial", bold=True)
    _fmt(ws1, f"B2:B{rjum}", "0.0")
    _fmt(ws1, f"C2:E{rjum}", "0.00")
    ws1["H1"].number_format = "0.00"
    _widths(ws1, [15, 20, 21, 22, 12, 3, 12, 34])
    _legend(ws1, f"A{rjum + 2}")

    # ---- Gradasi Rencana
    wsg.append(["Saringan (mm)", "BB", "BA", "% Lolos (Nilai Tengah)", "% Tertahan"])
    _hdr(wsg, 1, 5)
    for i, s in enumerate(S):
        r = i + 2
        wsg.cell(r, 1, s)
        if s in sp:
            wsg.cell(r, 2, sp[s][0])
            wsg.cell(r, 3, sp[s][1])
            wsg.cell(r, 4, f"=AVERAGE(B{r},C{r})")
        else:  # tanpa batas: mengikuti nilai tengah sebelumnya (awal = 100)
            wsg.cell(r, 4, f"=D{r - 1}" if i else 100)
        wsg.cell(r, 5, f"=100-D{r}" if i == 0 else f"=D{r - 1}-D{r}")
    rg_pan, rg_jum = len(S) + 2, len(S) + 3
    wsg.cell(rg_pan, 1, "Pan")
    wsg.cell(rg_pan, 4, 0)
    wsg.cell(rg_pan, 5, f"=D{rg_pan - 1}")
    wsg.cell(rg_jum, 1, "Jumlah")
    wsg.cell(rg_jum, 5, f"=SUM(E2:E{rg_pan})")
    for c in range(1, 6):
        wsg.cell(rg_jum, c).font = Font(name="Arial", bold=True)
    spec_s = [s for s in S if s in sp]
    for c, h in enumerate(["Saringan (mm)", "BB", "BA", "Nilai Tengah"], 7):
        wsg.cell(1, c, h)
    _hdr(wsg, 1, 4, 7)
    wsg["G1"].value = "Saringan (mm)"
    for j, s in enumerate(spec_s):
        r, rs = j + 2, S.index(s) + 2
        for c, col in zip(range(7, 11), "ABCD"):
            wsg.cell(r, c, f"={col}{rs}")
    wsg.cell(len(spec_s) + 3, 7, "Blok G:J = data grafik (hanya saringan yang diatur spesifikasi).").font = \
        Font(name="Arial", italic=True, color="7F7F7F", size=9)
    _fmt(wsg, f"B2:E{rg_jum}", "0.00")
    _fmt(wsg, f"H2:J{len(spec_s) + 1}", "0.00")
    _widths(wsg, [15, 9, 9, 22, 13, 3, 15, 9, 9, 13])
    wsg.add_chart(_grad_chart(mix, wsg, len(spec_s), mid=True), "L2")
    ws1.add_chart(_grad_chart(mix, wsg, len(spec_s), ws_s=ws1), "J4")

    # ---- Pengecekan Batas
    wsc.append(["Saringan (mm)", "BB", "BA", "% Lolos Sampel", "Selisih ke batas (%)", "Status"])
    _hdr(wsc, 1, 6)
    r = 2
    for i, sz in enumerate(LAB):
        if sz not in sp:
            continue
        rg = S.index(sz) + 2
        wsc.cell(r, 1, sz)
        wsc.cell(r, 2, f"={_Q['G']}!B{rg}")
        wsc.cell(r, 3, f"={_Q['G']}!C{rg}")
        wsc.cell(r, 4, f"={_Q['A']}!E{i + 2}")
        ok = f"AND(D{r}>=B{r}-0.005,D{r}<=C{r}+0.005)"
        wsc.cell(r, 5, f'=IF(D{r}="","",IF({ok},0,IF(D{r}>C{r},D{r}-C{r},D{r}-B{r})))')
        wsc.cell(r, 6, f'=IF(D{r}="","",IF({ok},"Masuk batas","Di luar batas"))')
        r += 1
    _fmt(wsc, f"D2:E{max(r, 3)}", "0.00")
    _widths(wsc, [15, 9, 9, 16, 20, 16])

    # ---- Kadar Aspal
    r236, r075 = S.index(2.36) + 2, S.index(0.075) + 2
    wsk.append(["Parameter", "Nilai", "Keterangan"])
    _hdr(wsk, 1, 3)
    rows_k = [
        ("Jenis campuran", mix, "Pilihan di aplikasi"),
        ("Berat campuran (gr)", float(W), "Isian"),
        ("Konstanta K", float(K), "Isian"),
        ("AK (tertahan 2,36 mm) (%)", f"=100-{_Q['G']}!D{r236}", "100 - % lolos 2,36"),
        ("AH (lolos 2,36 - lolos 0,075) (%)", f"={_Q['G']}!D{r236}-{_Q['G']}!D{r075}", ""),
        ("F (lolos 0,075) (%)", f"={_Q['G']}!D{r075}", ""),
        ("Pb estimasi (%)", "=0.035*B5+0.045*B6+0.18*B7+B4", "Pb = 0,035 AK + 0,045 AH + 0,18 F + K"),
        ("Pb dibulatkan 0,5% (%)", "=INT(B8*2+0.5)/2", "Pembulatan ke 0,5% terdekat"),
        ("Kadar aspal yang dipakai (%)", "=B9" if abs(kadar - Pbr) < 1e-9 else float(kadar),
         "Isian. Boleh diganti angka lain"),
        ("Berat aspal (gr)", "=B10/100*B3", ""),
        ("Berat agregat (gr)", "=B3-B11", ""),
    ]
    for i, (a, b, c) in enumerate(rows_k, 2):
        wsk.cell(i, 1, a)
        wsk.cell(i, 2, b)
        wsk.cell(i, 3, c)
    for ref in ("B3", "B4", "B10"):
        _inp(wsk[ref])
    _fmt(wsk, "B5:B12", "0.000")
    _widths(wsk, [36, 16, 40])
    _legend(wsk, "A14")

    # ---- Berat per Saringan
    wsb.append(["Saringan (mm)", "% Tertahan", "Berat Tertahan (gr)"])
    _hdr(wsb, 1, 3)
    for i, s in enumerate(S + ["Pan"]):
        r = i + 2
        wsb.cell(r, 1, s)
        wsb.cell(r, 2, f"={_Q['G']}!E{r}")
        wsb.cell(r, 3, f"=B{r}/100*{_Q['K']}!$B$12")
    rj = len(S) + 3
    wsb.cell(rj, 1, "Jumlah")
    wsb.cell(rj, 2, f"=SUM(B2:B{rj - 1})")
    wsb.cell(rj, 3, f"=SUM(C2:C{rj - 1})")
    for c in range(1, 4):
        wsb.cell(rj, c).font = Font(name="Arial", bold=True)
    _fmt(wsb, f"B2:C{rj}", "0.00")
    _widths(wsb, [15, 14, 20])

    # ---- Koreksi
    wsr.append(["Komponen", "Sebelum Koreksi (%)", "Setelah Koreksi (%)", "Berat (gr)"])
    _hdr(wsr, 1, 4)
    for i, (nm, src) in enumerate([("Agregat kasar (AK)", "B5"), ("Agregat halus (AH)", "B6"), ("Filler (F)", "B7")], 2):
        wsr.cell(i, 1, nm)
        wsr.cell(i, 2, f"={_Q['K']}!{src}")
        wsr.cell(i, 3, f"=B{i}*(1-{_Q['K']}!$B$10/100)")
        wsr.cell(i, 4, f"=C{i}/100*{_Q['K']}!$B$3")
    wsr.cell(5, 1, "Aspal")
    wsr.cell(5, 3, f"={_Q['K']}!B10")
    wsr.cell(5, 4, f"=C5/100*{_Q['K']}!$B$3")
    wsr.cell(6, 1, "Total")
    wsr.cell(6, 2, "=SUM(B2:B4)")
    wsr.cell(6, 3, "=SUM(C2:C5)")
    wsr.cell(6, 4, "=SUM(D2:D5)")
    for c in range(1, 5):
        wsr.cell(6, c).font = Font(name="Arial", bold=True)
    _fmt(wsr, "B2:D6", "0.00")
    _widths(wsr, [22, 20, 20, 14])

    # ---- Variasi kadar aspal
    wsv.append(["Variasi", "Kadar Aspal (%)", "Berat Aspal (gr)", "Berat Agregat (gr)"])
    _hdr(wsv, 1, 4)
    for i, (lb, k) in enumerate(zip(labels, kadar_list), 2):
        off = round(k - kadar, 2)
        wsv.cell(i, 1, lb)
        wsv.cell(i, 2, f"=ROUND({_Q['K']}!$B$10+({off}),2)")
        wsv.cell(i, 3, f"=B{i}/100*{_Q['K']}!$B$3")
        wsv.cell(i, 4, f"={_Q['K']}!$B$3-C{i}")
    _fmt(wsv, f"B2:D{len(kadar_list) + 1}", "0.00")
    _widths(wsv, [14, 16, 18, 20])

    # ---- Variasi berat saringan
    wsw.cell(1, 1, "Saringan (mm)")
    for j in range(len(kadar_list)):
        wsw.cell(1, 2 + j, f'={_Q["V"]}!B{2 + j}&"% (gr)"')
    _hdr(wsw, 1, 1 + len(kadar_list))
    for i, s in enumerate(S + ["Pan"]):
        r = i + 2
        wsw.cell(r, 1, s)
        for j in range(len(kadar_list)):
            wsw.cell(r, 2 + j, f"={_Q['B']}!$B{r}/100*({_Q['K']}!$B$3-{_Q['V']}!$B${2 + j}/100*{_Q['K']}!$B$3)")
    rj = len(S) + 3
    wsw.cell(rj, 1, "Jumlah").font = Font(name="Arial", bold=True)
    for j in range(len(kadar_list)):
        c = wsw.cell(rj, 2 + j, f"=SUM({CL(2 + j)}2:{CL(2 + j)}{rj - 1})")
        c.font = Font(name="Arial", bold=True)
    _fmt(wsw, f"B2:{CL(1 + len(kadar_list))}{rj}", "0.00")
    _widths(wsw, [15] + [14] * len(kadar_list))

    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def xlsx_marshall(mix, specs, bj, kal_in, unit, props, rows_in, kadar_vals, raw=None):
    """Workbook Tab 4: VMA/VIM/VFA/stabilitas/MQ, rata-rata, cek syarat, KAO berupa rumus + 6 grafik Excel."""
    wb = Workbook()
    wsi = wb.active
    wsi.title = "Input"
    wst = wb.create_sheet("Tabel Korelasi")
    wsp = wb.create_sheet("Per Benda Uji")
    wsr = wb.create_sheet("Rata-rata")
    wsc = wb.create_sheet("Pengecekan Syarat")
    wsx = wb.create_sheet("Grafik Marshall")
    wsd = wb.create_sheet("Data Grafik")

    # ---- Input
    wsi.append(["Parameter", "Nilai", "Keterangan"])
    _hdr(wsi, 1, 3)
    pCA, pFA, pFF = props
    rows_i = [
        ("BJ bulk agregat kasar", _num(bj["kasar_bulk"]), "Isian"),
        ("BJ bulk agregat halus", _num(bj["halus_bulk"]), "Isian"),
        ("BJ bulk filler", _num(bj["filler_bulk"]), "Isian"),
        ("BJ semu agregat kasar", _num(bj["kasar_semu"]), "Isian"),
        ("BJ semu agregat halus", _num(bj["halus_semu"]), "Isian"),
        ("BJ semu filler", _num(bj["filler_semu"]), "Isian"),
        ("BJ aspal", _num(bj["aspal"]), "Isian"),
        ("Faktor kalibrasi alat (per skala)", _num(kal_in), "Isian"),
        ("Satuan faktor kalibrasi", unit, "Isian: kg atau lbs"),
        ("Faktor kalibrasi (kg per skala)", '=IF(B9="","",IF(B10="lbs",B9*0.45359237,B9))', "1 lbs = 0,45359237 kg"),
        ("Proporsi agregat kasar (%)", float(pCA), f"Dari gradasi rencana {mix} (tertahan 4,75 mm)"),
        ("Proporsi agregat halus (%)", float(pFA), "Dari gradasi rencana"),
        ("Proporsi filler (%)", float(pFF), "Dari gradasi rencana"),
        ("Jenis campuran", mix, ""),
        ("BJ bulk agregat gabungan (Gsb)", '=IF(COUNT(B2:B8)<7,"",100/(B12/B2+B13/B3+B14/B4))', ""),
        ("BJ semu agregat gabungan (Gsa)", '=IF(COUNT(B2:B8)<7,"",100/(B12/B5+B13/B6+B14/B7))', ""),
        ("BJ efektif agregat gabungan (Gse)",
         '=IF(COUNT(B2:B8)<7,"",100/(B12/((B2+B5)/2)+B13/((B3+B6)/2)+B14/((B4+B7)/2)))',
         "Rata-rata BJ bulk dan semu tiap fraksi"),
    ]
    for i, (a, b, c) in enumerate(rows_i, 2):
        wsi.cell(i, 1, a)
        wsi.cell(i, 2, b)
        wsi.cell(i, 3, c)
    for r in list(range(2, 11)) + [12, 13, 14]:
        _inp(wsi.cell(r, 2))
    dv = DataValidation(type="list", formula1='"kg,lbs"', allow_blank=False)
    wsi.add_data_validation(dv)
    dv.add("B10")
    _fmt(wsi, "B2:B9", "0.0000")
    _fmt(wsi, "B11:B18", "0.0000")
    _widths(wsi, [38, 16, 48])
    _legend(wsi, "A20")

    # ---- Berat jenis dari berat hasil pengujian (rumus), menggantikan isian langsung di sheet Input
    if raw is not None:
        wsb = wb.create_sheet("Berat Jenis", 1)
        blk = [
            (1, "AGREGAT KASAR (SNI 1969)", [("Berat benda uji kering oven, Bk (gr)", "k_bk"), ("Berat benda uji SSD di udara, Bj (gr)", "k_bj"),
                                              ("Berat benda uji SSD dalam air, Ba (gr)", "k_ba")]),
            (10, "AGREGAT HALUS (SNI 1970)", [("Berat benda uji kering oven, Bk (gr)", "h_bk"), ("Berat piknometer + air, B (gr)", "h_b"),
                                               ("Berat piknometer + benda uji + air, Bt (gr)", "h_bt"), ("Berat benda uji SSD, S (gr)", "h_s")]),
            (20, "FILLER (piknometer)", [("Berat filler kering, Bk (gr)", "f_bk"), ("Berat piknometer + air, B (gr)", "f_b"),
                                         ("Berat piknometer + filler + air, Bt (gr)", "f_bt")]),
            (26, "ASPAL (SNI 2441)", [("Berat piknometer kosong, A (gr)", "a_a"), ("Berat piknometer + air, B (gr)", "a_b"),
                                      ("Berat piknometer + aspal, C (gr)", "a_c"), ("Berat piknometer + aspal + air, D (gr)", "a_d")]),
        ]
        for r0, ttl, items in blk:
            wsb.cell(r0, 1, ttl)
            _hdr(wsb, r0, 2)
            for j, (lb, key) in enumerate(items, 1):
                wsb.cell(r0 + j, 1, lb)
                _inp(wsb.cell(r0 + j, 2, _num(raw.get(key))))
        res = [
            (5, "BJ bulk", '=IFERROR(B2/(B3-B4),"")'), (6, "BJ SSD", '=IFERROR(B3/(B3-B4),"")'),
            (7, "BJ semu", '=IFERROR(B2/(B2-B4),"")'), (8, "Penyerapan (%)", '=IFERROR((B3-B2)/B2*100,"")'),
            (15, "BJ bulk", '=IFERROR(B11/(B12+B14-B13),"")'), (16, "BJ SSD", '=IFERROR(B14/(B12+B14-B13),"")'),
            (17, "BJ semu", '=IFERROR(B11/(B12+B11-B13),"")'), (18, "Penyerapan (%)", '=IFERROR((B14-B11)/B11*100,"")'),
            (24, "BJ filler (dipakai sebagai bulk dan semu)", '=IFERROR(B21/(B22+B21-B23),"")'),
            (31, "BJ aspal", '=IFERROR((B29-B27)/((B28-B27)-(B30-B29)),"")'),
        ]
        for r_, lb, fx in res:
            wsb.cell(r_, 1, lb).font = Font(name="Arial", bold=True)
            wsb.cell(r_, 2, fx).number_format = "0.0000"
        _widths(wsb, [46, 16])
        _legend(wsb, "A33")
        # sheet Input mengambil nilai dari sheet Berat Jenis
        for r_, ref in {2: "B5", 3: "B15", 4: "B24", 5: "B7", 6: "B17", 7: "B24", 8: "B31"}.items():
            c_ = wsi.cell(r_, 2, f"='Berat Jenis'!{ref}")
            c_.fill = PatternFill(fill_type=None)
            c_.font = Font(name="Arial", color="000000")
            c_.number_format = "0.0000"
            wsi.cell(r_, 3, "Dihitung otomatis dari sheet Berat Jenis")

    # ---- Tabel korelasi
    wst.append(["Isi min (cm3)", "Isi maks (cm3)", "Titik tengah", "Angka korelasi", "Tebal benda uji (mm)"])
    _hdr(wst, 1, 5)
    for i, (a, b, c) in enumerate(KORELASI, 2):
        wst.cell(i, 1, a)
        wst.cell(i, 2, b)
        wst.cell(i, 3, f"=(A{i}+B{i})/2")
        wst.cell(i, 4, c)
        wst.cell(i, 5, KORELASI_TEBAL[i - 2])
    k1, k2 = 2, len(KORELASI) + 1
    wst.cell(k2 + 2, 1, "Sumber: SNI 06-2489-1991. Angka korelasi untuk volume di antara titik tengah diinterpolasi linear. Kolom tebal dipakai bila angka korelasi ditentukan dari tebal benda uji.").font = \
        Font(name="Arial", italic=True, color="7F7F7F", size=9)
    _widths(wst, [14, 14, 14, 16, 20])

    # ---- Per benda uji
    cols_m = list(rows_in.columns)
    data = [[_num(rw[c]) for c in cols_m] for _, rw in rows_in.iterrows()]
    data = [d for d in data if any(v is not None for v in d)]
    last = len(data) + 5 + 1  # + 5 baris cadangan
    wsp.append(["No"] + cols_m + ["Volume (cm3)", "BJ Bulk Campuran", "BJ Teoritis Maks", "VMA (%)", "VIM (%)",
                                  "VFA (%)", "Angka Korelasi", "Stabilitas (kg)", "Flow (mm)", "MQ (kg/mm)"])
    _hdr(wsp, 1, 17)
    wsp.row_dimensions[1].height = 45
    tk_x, tk_y = f"{_Q['T']}!$C${k1}:$C${k2}", f"{_Q['T']}!$D${k1}:$D${k2}"
    for r in range(2, last + 1):
        wsp.cell(r, 1, r - 1)
        vals = data[r - 2] if r - 2 < len(data) else [None] * 7
        for c, v in enumerate(vals, 2):
            _inp(wsp.cell(r, c, v))
        mt = f"MATCH(I{r},{tk_x},1)"
        interp = (f"ROUND(INDEX({tk_y},{mt})+(I{r}-INDEX({tk_x},{mt}))*(INDEX({tk_y},{mt}+1)-INDEX({tk_y},{mt}))"
                  f"/(INDEX({tk_x},{mt}+1)-INDEX({tk_x},{mt})),2)")
        f = {
            "I": f'=IF(COUNT(B{r}:E{r})<4,"",IF(OR(D{r}-E{r}<=0,B{r}<=0,B{r}>=100),"",D{r}-E{r}))',
            "J": f'=IF(I{r}="","",C{r}/I{r})',
            "K": f'=IF(OR(I{r}="",Input!$B$18="",Input!$B$8=""),"",100/((100-B{r})/Input!$B$18+B{r}/Input!$B$8))',
            "L": f'=IF(OR(K{r}="",Input!$B$16=""),"",100-J{r}*(100-B{r})/Input!$B$16)',
            "M": f'=IF(K{r}="","",100*(K{r}-J{r})/K{r})',
            "N": f'=IF(OR(L{r}="",M{r}=""),"",100*(L{r}-M{r})/L{r})',
            "O": f'=IF(I{r}="","",IF(ISNUMBER(H{r}),H{r},IF(I{r}<={tk_x.split(":")[0]},{tk_y.split(":")[0]},'
                 f'IF(I{r}>=INDEX({tk_x},{k2 - k1 + 1}),INDEX({tk_y},{k2 - k1 + 1}),{interp}))))',
            "P": f'=IF(OR(I{r}="",F{r}="",Input!$B$11=""),"",F{r}*Input!$B$11*O{r})',
            "Q": f'=IF(I{r}="","",IF(ISNUMBER(G{r}),IF(G{r}>0,G{r},""),""))',
            "R": f'=IF(OR(P{r}="",Q{r}=""),"",P{r}/Q{r})',
        }
        for col, fx in f.items():
            wsp[f"{col}{r}"] = fx
    _fmt(wsp, f"I2:K{last}", "0.000")
    _fmt(wsp, f"L2:N{last}", "0.00")
    _fmt(wsp, f"O2:O{last}", "0.00")
    _fmt(wsp, f"P2:R{last}", "0.00")
    _widths(wsp, [5] + [13] * 7 + [12, 13, 13, 10, 10, 10, 12, 14, 10, 12])
    _legend(wsp, f"A{last + 2}", "Isi / ubah kolom kuning (B-H). Kolom I-R berisi rumus. Baris cadangan sudah berisi rumus: "
                                  "untuk menambah baris, salin rumus kolom I-R ke bawah dan perluas rentang di sheet Rata-rata.")

    # ---- Rata-rata per kadar aspal
    n = len(kadar_vals)
    prng = lambda col: f"{_Q['P']}!${col}$2:${col}${last}"
    src = {"VIM (%)": "M", "VMA (%)": "L", "VFA (%)": "N", "Stabilitas (kg)": "P", "Flow (mm)": "Q", "MQ (kg/mm)": "R"}
    wsr.append(["Kadar Aspal (%)", "Jumlah Benda Uji"] + MPARAMS)
    _hdr(wsr, 1, 2 + len(MPARAMS))
    for i, kv in enumerate(kadar_vals, 2):
        wsr.cell(i, 1, float(kv))
        wsr.cell(i, 2, f'=COUNTIFS({prng("B")},A{i},{prng("I")},">0")')
        for j, p in enumerate(MPARAMS):
            wsr.cell(i, 3 + j, f'=IFERROR(AVERAGEIFS({prng(src[p])},{prng("B")},$A{i}),"")')
    _fmt(wsr, f"C2:H{n + 1}", "0.000")
    _widths(wsr, [16, 16] + [13] * 6)

    # ---- Pengecekan syarat + KAO
    wsc.append(["Parameter", "Min", "Maks"])
    _hdr(wsc, 1, 3)
    for j, p in enumerate(MPARAMS, 2):
        wsc.cell(j, 1, p)
        lo, hi = specs.get(p, (None, None))
        _inp(wsc.cell(j, 2, _num(lo)))
        _inp(wsc.cell(j, 3, _num(hi)))
    e = 12 + n
    wsc["A9"] = "KAO (%)"
    wsc["A9"].font = Font(name="Arial", bold=True)
    wsc["B9"] = (f'=IF(COUNT(I13:I{e})>0,INDEX(A13:A{e},MATCH(MAX(I13:I{e}),I13:I{e},0)),'
                 f'IF(COUNT(J13:J{e})>0,MEDIAN(J13:J{e}),"Belum ada yang memenuhi"))')
    wsc["C9"] = "Kadar yang memenuhi semua syarat dengan stabilitas tertinggi"
    wsc["B9"].font = Font(name="Arial", bold=True, color="0E9F6E")
    wsc.append([])
    hdr_row = 12
    for c, h in enumerate(["Kadar Aspal (%)"] + MPARAMS + ["Semua memenuhi", "Stabilitas jika memenuhi",
                                                          "Kadar jika memenuhi"], 1):
        wsc.cell(hdr_row, c, h)
    _hdr(wsc, hdr_row, 10)
    wsc.row_dimensions[hdr_row].height = 32
    for r in range(13, e + 1):
        rr = r - 11
        wsc.cell(r, 1, f"={_Q['R']}!A{rr}")
        for j in range(6):
            v, lo, hi = f"{_Q['R']}!{CL(3 + j)}{rr}", f"$B${2 + j}", f"$C${2 + j}"
            wsc.cell(r, 2 + j, f'=IF(AND({lo}="",{hi}=""),"-",IF({v}="","Tidak",IF(AND(OR({lo}="",{v}>={lo}-1E-9),'
                               f'OR({hi}="",{v}<={hi}+1E-9)),"Memenuhi","Tidak")))')
        wsc.cell(r, 8, f'=IF(AND(COUNTIF(B{r}:G{r},"Memenuhi")>0,COUNTIF(B{r}:G{r},"Memenuhi")+COUNTIF(B{r}:G{r},"-")=6),"Ya","Tidak")')
        wsc.cell(r, 9, f'=IF(H{r}="Ya",{_Q["R"]}!F{rr},"")')
        wsc.cell(r, 10, f'=IF(H{r}="Ya",A{r},"")')
    _fmt(wsc, f"I13:I{e}", "0.00")
    _widths(wsc, [18, 12, 12, 12, 17, 12, 12, 16, 20, 18])
    _legend(wsc, f"A{e + 2}", "Syarat min/maks (kuning) boleh diubah. KAO dan status ikut berubah otomatis.")

    # ---- Data grafik (memakai NA() agar titik kosong tidak digambar)
    wsd.cell(1, 1, "Kadar Aspal (%)")
    rA, rB = n + 3, n + 4
    wsd.cell(rA - 1, 1, "KAO (garis vertikal)")
    wsd.cell(rA, 1, f'=IF(ISNUMBER({_Q["C"]}!$B$9),{_Q["C"]}!$B$9,NA())')
    wsd.cell(rB, 1, f"=A{rA}")
    for i in range(2, n + 2):
        wsd.cell(i, 1, f"={_Q['R']}!A{i}")
    for j, p in enumerate(MPARAMS):
        c0 = 2 + 4 * j
        for k, h in enumerate([f"{p} hasil", f"{p} min", f"{p} maks", f"{p} KAO"]):
            wsd.cell(1, c0 + k, h)
        rc = CL(3 + j)
        for i in range(2, n + 2):
            wsd.cell(i, c0, f'=IF({_Q["R"]}!{rc}{i}="",NA(),{_Q["R"]}!{rc}{i})')
            wsd.cell(i, c0 + 1, f'=IF({_Q["C"]}!$B${2 + j}="",NA(),{_Q["C"]}!$B${2 + j})')
            wsd.cell(i, c0 + 2, f'=IF({_Q["C"]}!$C${2 + j}="",NA(),{_Q["C"]}!$C${2 + j})')
        rngc = f"{_Q['R']}!{rc}2:{rc}{n + 1}"
        wsd.cell(rA, c0 + 3, f'=MIN(MIN({rngc}),IF({_Q["C"]}!$B${2 + j}="",MIN({rngc}),{_Q["C"]}!$B${2 + j}))')
        wsd.cell(rB, c0 + 3, f'=MAX(MAX({rngc}),IF({_Q["C"]}!$C${2 + j}="",MAX({rngc}),{_Q["C"]}!$C${2 + j}))')
    wsd.cell(rB + 2, 1, "Sheet pendukung grafik. Sel berisi NA() itu disengaja: Excel tidak menggambar titik tersebut "
                        "(syarat yang tidak diisi atau data yang kosong).").font = Font(name="Arial", italic=True, color="7F7F7F", size=9)
    _hdr(wsd, 1, 1 + 4 * len(MPARAMS))
    wsd.row_dimensions[1].height = 32
    _widths(wsd, [16] + [14] * 24)

    # ---- Grafik Marshall (2 x 3)
    order = ["Stabilitas (kg)", "Flow (mm)", "MQ (kg/mm)", "VMA (%)", "VIM (%)", "VFA (%)"]
    anchors = ["A1", "J1", "S1", "A18", "J18", "S18"]
    xr = Reference(wsd, min_col=1, min_row=2, max_row=n + 1)
    xk = Reference(wsd, min_col=1, min_row=rA, max_row=rB)
    for p, anc in zip(order, anchors):
        c0 = 2 + 4 * MPARAMS.index(p)
        ch = _scatter(p, "Kadar Aspal (%)", p, 14, 8.5)
        ch.series.append(_ser(xr, Reference(wsd, min_col=c0, min_row=2, max_row=n + 1), "Hasil uji", "F97316"))
        ch.series.append(_ser(xr, Reference(wsd, min_col=c0 + 1, min_row=2, max_row=n + 1), "Syarat min", "1F3A5F",
                              dash="dash", marker="none", width=19050))
        ch.series.append(_ser(xr, Reference(wsd, min_col=c0 + 2, min_row=2, max_row=n + 1), "Syarat maks", "6FA8DC",
                              dash="dash", marker="none", width=19050))
        ch.series.append(_ser(xk, Reference(wsd, min_col=c0 + 3, min_row=rA, max_row=rB), "KAO", "0E9F6E",
                              dash="sysDot", marker="none", width=25400))
        wsx.add_chart(ch, anc)

    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


st.set_page_config(page_title="GradAspal", page_icon="🛣️", layout="wide")

import random as _rnd


def _hero_html():
    """Hero dengan ilustrasi 3D lapisan perkerasan (isometrik, melayang)."""
    cx, W2, H2, gap = 150, 104, 52, 16
    layers = [("AC-WC", 14, "#3A4048", "#20242A", "#2B3036", "Lapis aus"),
              ("AC-BC", 16, "#4B535D", "#2D333A", "#393F48", "Lapis antara"),
              ("AC-Base", 20, "#5E6874", "#38404A", "#47505B", "Pondasi aspal"),
              ("LPA", 26, "#C2AE82", "#8E7E5C", "#A69468", "Pondasi agregat"),
              ("Tanah", 34, "#A0704A", "#6F4A2D", "#86593A", "Tanah dasar")]
    ys, y = [], 24
    for _, t, *_r in layers:
        ys.append(y)
        y += t + gap
    rnd = _rnd.Random(11)
    parts = []
    for i in reversed(range(len(layers))):  # gambar dari bawah ke atas
        nm, t, top, left, right, lab = layers[i]
        cy = ys[i]
        T = f"{cx},{cy} {cx + W2},{cy + H2} {cx},{cy + 2 * H2} {cx - W2},{cy + H2}"
        L = f"{cx - W2},{cy + H2} {cx},{cy + 2 * H2} {cx},{cy + 2 * H2 + t} {cx - W2},{cy + H2 + t}"
        R = f"{cx},{cy + 2 * H2} {cx + W2},{cy + H2} {cx + W2},{cy + H2 + t} {cx},{cy + 2 * H2 + t}"
        g = [f'<g class="slab" style="animation-delay:{i * 0.35:.2f}s">',
             f'<polygon points="{L}" fill="{left}"/>', f'<polygon points="{R}" fill="{right}"/>',
             f'<polygon points="{T}" fill="{top}" stroke="rgba(255,255,255,.18)" stroke-width="1"/>']
        if i <= 2:  # tekstur butir agregat
            for _ in range(26):
                px, py = rnd.uniform(-0.85, 0.85), rnd.uniform(-0.85, 0.85)
                if abs(px) + abs(py) <= 0.88:
                    g.append(f'<circle cx="{cx + px * W2:.1f}" cy="{cy + H2 + py * H2:.1f}" r="{rnd.uniform(.9, 1.7):.1f}" '
                             f'fill="rgba(255,255,255,{rnd.uniform(.10, .30):.2f})"/>')
        if i == 0:  # marka jalan putus-putus
            g.append(f'<line x1="{cx - W2 / 2}" y1="{cy + H2 / 2}" x2="{cx + W2 / 2}" y2="{cy + 1.5 * H2}" '
                     f'stroke="#FBBF24" stroke-width="3.2" stroke-dasharray="11 8" stroke-linecap="round"/>')
        ly = cy + H2 + t / 2
        g.append(f'<line x1="{cx + W2 + 4}" y1="{ly}" x2="{cx + W2 + 22}" y2="{ly}" stroke="#FDE68A" stroke-width="1.2"/>')
        g.append(f'<text x="{cx + W2 + 27}" y="{ly + 4}" fill="#E8F0F8" font-size="12" font-weight="700" '
                 f'font-family="Arial,sans-serif">{nm}</text>')
        g.append(f'<text x="{cx + W2 + 27}" y="{ly + 17}" fill="#9FB3C8" font-size="9.5" '
                 f'font-family="Arial,sans-serif">{lab}</text>')
        g.append('</g>')
        parts.append("".join(g))
    base_y = ys[-1] + 2 * H2 + layers[-1][1] + 18
    svg = (f'<svg viewBox="0 0 420 {base_y + 18}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Lapisan perkerasan jalan">'
           f'<defs><radialGradient id="shd" cx="50%" cy="50%" r="50%"><stop offset="0%" stop-color="#000" stop-opacity=".55"/>'
           f'<stop offset="100%" stop-color="#000" stop-opacity="0"/></radialGradient></defs>'
           f'<ellipse cx="{cx}" cy="{base_y}" rx="132" ry="15" fill="url(#shd)"/>' + "".join(parts) + '</svg>')
    return ('<div class="hero"><div class="hero-txt"><h1>🛣️ GradAspal</h1>'
            '<p>Perencanaan gradasi agregat, kadar aspal estimasi, dan kadar aspal optimum (uji Marshall) untuk Latasir, Lataston, dan Laston</p>'
            '<div class="chips"><span>Latasir</span><span>Lataston</span><span>Laston</span><span>Uji Marshall</span><span>Tanpa login</span></div>'
            '<div class="tagline">Atur campuran, isi data, dan temukan KAO secara otomatis.</div></div>'
            f'<div class="hero-art">{svg}</div></div>')


def _idn(x, nd=2):
    return f"{x:,.{nd}f}".replace(",", "_").replace(".", ",").replace("_", ".")


def _dash_html(mix, K, W, tot, mhb, ok_n, ok_c, Pb, Pbr, kadar, rm_n, kao):
    """Dasbor ringkasan: 5 kartu 3D + stepper progres."""
    def card(cls, ic, lb, vl, sb):
        return f'<div class="kpi {cls}"><div class="ic">{ic}</div><div class="lb">{lb}</div><div class="vl">{vl}</div><div class="sb">{sb}</div></div>'
    c1 = card("k1", "🧱", "Jenis campuran", mix, f"Konstanta K = {_idn(K, 2)}")
    if W > 0:
        wa = W * kadar / 100
        c2 = card("k2", "⚖️", "Berat campuran", f"{_idn(W, 0)} gr", f"Aspal {_idn(wa, 1)} gr · Agregat {_idn(W - wa, 1)} gr")
    else:
        c2 = card("k2", "⚖️", "Berat campuran", "Belum diisi", "Isi di sidebar")
    if tot > 0 and ok_c:
        c3 = card("k3", "📊", "Gradasi sampel", f"{ok_n}/{ok_c} masuk batas", f"MHB {_idn(mhb, 2)}")
    else:
        c3 = card("k3", "📊", "Gradasi sampel", "Belum diisi", "Isi massa di Tab 1")
    c4 = card("k4", "🛢️", "Pb estimasi", f"{_idn(Pb, 2)}%", f"Dibulatkan {_idn(Pbr, 1)}% · dipakai {_idn(kadar, 1)}%")
    if kao is not None:
        c5 = card("k5", "🎯", "Kadar aspal optimum", f"KAO {_idn(kao, 1)}%", "Dari uji Marshall")
    elif rm_n:
        c5 = card("k6", "🎯", "Kadar aspal optimum", "Belum ditentukan", "Cek syarat di Tab 4")
    else:
        c5 = card("k6", "🎯", "Kadar aspal optimum", "Belum ada data", "Isi uji Marshall di Tab 4")
    steps = [("Atur campuran", W > 0), ("Isi massa saringan", tot > 0), ("Kadar aspal & variasi", W > 0 and tot > 0),
             ("Data Marshall", rm_n > 0), ("KAO ditemukan", kao is not None)]
    now = next((i for i, (_, d) in enumerate(steps) if not d), None)
    done_n = sum(d for _, d in steps)
    sp_html = ""
    for i, (nm, d) in enumerate(steps):
        cls = "done" if d else ("now" if i == now else "")
        sp_html += f'<div class="stp {cls}"><div class="dot">{"✓" if d else i + 1}</div>{nm}</div>'
    return (f'<div class="dash">{c1}{c2}{c3}{c4}{c5}</div>'
            f'<div class="steps-wrap"><div class="steps-ttl">Progres pekerjaan: {done_n} dari {len(steps)} langkah</div>'
            f'<div class="steps">{sp_html}</div></div>')



def _mini_logo():
    cx, W2, H2 = 32, 26, 13
    L = [(6, 4, "#3A4048", "#20242A", "#2B3036"), (20, 5, "#5E6874", "#38404A", "#47505B"), (34, 7, "#C2AE82", "#8E7E5C", "#A69468")]
    out = ""
    for cy, t, top, lf, rt in reversed(L):
        out += (f'<polygon points="{cx - W2},{cy + H2} {cx},{cy + 2 * H2} {cx},{cy + 2 * H2 + t} {cx - W2},{cy + H2 + t}" fill="{lf}"/>'
                f'<polygon points="{cx},{cy + 2 * H2} {cx + W2},{cy + H2} {cx + W2},{cy + H2 + t} {cx},{cy + 2 * H2 + t}" fill="{rt}"/>'
                f'<polygon points="{cx},{cy} {cx + W2},{cy + H2} {cx},{cy + 2 * H2} {cx - W2},{cy + H2}" fill="{top}" stroke="rgba(255,255,255,.2)"/>')
    out += f'<line x1="{cx - 13}" y1="{4 + 6.5}" x2="{cx + 13}" y2="{4 + 19.5}" stroke="#FBBF24" stroke-width="2" stroke-dasharray="5 4" stroke-linecap="round"/>'
    return f'<svg viewBox="0 0 64 62" xmlns="http://www.w3.org/2000/svg">{out}</svg>'


def _brand_html():
    return (f'<div class="sb-brand">{_mini_logo()}<div><b>GradAspal</b><small>Gradasi, Pb, dan KAO<br>untuk perkerasan lentur</small></div></div>')


def _side_html(mix, W, Pb, kadar, tot, ok_n, ok_c, rm_n, kao):
    def row(a, b, cls=""):
        return f'<div class="sbrow"><span>{a}</span><b class="{cls}">{b}</b></div>'
    done = [W > 0, tot > 0, W > 0 and tot > 0, rm_n > 0, kao is not None]
    pct = int(sum(done) / len(done) * 100)
    r = row("Campuran", mix)
    r += row("Berat", f"{_idn(W, 0)} gr" if W > 0 else "belum diisi", "ok" if W > 0 else "no")
    r += row("Gradasi", f"{ok_n}/{ok_c} masuk" if tot > 0 and ok_c else "belum diisi",
             ("ok" if ok_n == ok_c else "no") if tot > 0 and ok_c else "no")
    r += row("Pb estimasi", f"{_idn(Pb, 2)}%")
    r += row("KAO", f"{_idn(kao, 1)}%" if kao is not None else "belum ada", "ok" if kao is not None else "no")
    return (f'<div class="sbcard"><div class="sbc-t">📌 Ringkasan cepat</div>{r}'
            f'<div class="sbbar"><i style="width:{pct}%"></i></div><div class="sbpct">Progres {pct}%</div></div>')


if "started" not in st.session_state:
    st.session_state["started"] = st.query_params.get("mulai") == "1"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@600;700&display=swap');
:root { --asp9:#14202E; --asp7:#243B53; --asp5:#486581; --asp3:#9FB3C8; --asp1:#E6EDF5;
        --amber:#F59E0B; --amber2:#F97316; --ok:#0E9F6E; }
.stApp { background: linear-gradient(180deg, #F5F8FC 0%, #E7EEF6 100%); }
[data-testid="stHeader"] { background: transparent; }

/* ---------- sidebar */
[data-testid="stSidebar"] { background: linear-gradient(180deg, #E9F0F8 0%, #D5E1EE 100%);
        border-right: 1px solid #C7D6E6; box-shadow: 2px 0 14px rgba(20,32,46,.08); }
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: var(--asp9); }

/* ---------- hero */
.hero { position: relative; overflow: hidden; display: grid; grid-template-columns: 1.15fr .95fr; align-items: center; gap: .5rem;
        background: linear-gradient(120deg, #101A26 0%, #1C3047 50%, #3E5C7E 100%);
        padding: 1.6rem 2rem 2.2rem 2rem; border-radius: 20px; margin-bottom: 1.1rem;
        box-shadow: 0 10px 0 -3px #0B131D, 0 24px 40px rgba(20,32,46,.38); }
.hero-txt { position: relative; z-index: 2; }
.hero-art { position: relative; z-index: 2; display: flex; justify-content: center; }
.hero-art svg { width: 100%; max-width: 430px; height: auto; filter: drop-shadow(0 14px 16px rgba(0,0,0,.45)); }
.slab { animation: floaty 4.8s ease-in-out infinite; }
@keyframes floaty { 0%, 100% { transform: translateY(0); } 50% { transform: translateY(-7px); } }
.tagline { position: relative; margin-top: .55rem; color: #FDE68A; font-weight: 600; font-size: .95rem; }
@keyframes roadmove { from { background-position: 0 0; } to { background-position: 58px 0; } }
.hero::after { content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 7px;
        background: repeating-linear-gradient(90deg, #F59E0B 0 34px, transparent 34px 58px); animation: roadmove 2s linear infinite; }
.hero::before { content: ""; position: absolute; right: -60px; top: -70px; width: 260px; height: 260px; border-radius: 50%;
        background: radial-gradient(circle, rgba(245,158,11,.30) 0%, rgba(245,158,11,0) 70%); }
.hero h1 { position: relative; color: #FFFFFF; margin: 0; font-size: 2.4rem; letter-spacing: .4px;
        font-family: 'Poppins', 'Source Sans Pro', sans-serif; }
.hero p { position: relative; color: #C9D8E8; margin: .35rem 0 .8rem 0; font-size: 1.02rem; }
.chips span { display: inline-block; background: rgba(255,255,255,.12); color: #FDE68A; border: 1px solid rgba(253,230,138,.45);
        border-radius: 999px; padding: .12rem .8rem; margin: 0 .4rem .3rem 0; font-size: .82rem; font-weight: 600; }

/* ---------- tab */
.stTabs [data-baseweb="tab-list"], .stTabs [role="tablist"] { gap: 6px; border-bottom: none; padding-bottom: 2px; }
.stTabs [data-baseweb="tab"], [data-testid="stTab"] { background: #FFFFFF; border: 1px solid #D5E1EE; border-bottom: none;
        border-radius: 12px 12px 0 0; padding: .55rem 1rem; font-weight: 600; color: var(--asp5);
        box-shadow: 0 -1px 4px rgba(20,32,46,.05); }
.stTabs [data-baseweb="tab"]:hover, [data-testid="stTab"]:hover { background: #FFF7E0; color: var(--asp9); }
.stTabs [aria-selected="true"], [data-testid="stTab"][aria-selected="true"] { background: var(--asp7) !important; color: #FFFFFF !important; border-color: var(--asp7); }
.stTabs [data-baseweb="tab-highlight"] { background-color: var(--amber) !important; height: 4px; }
.stTabs [data-baseweb="tab-border"] { background-color: #D5E1EE; }

/* ---------- judul bagian */
[data-testid="stMainBlockContainer"] h2, [data-testid="stMainBlockContainer"] h3 { color: var(--asp9); border-left: 7px solid var(--amber);
        padding-left: .7rem; font-family: 'Poppins', 'Source Sans Pro', sans-serif; }

/* ---------- kartu metrik */
[data-testid="stMetric"] { background: #FFFFFF; border: 1px solid #DCE6F0; border-left: 6px solid var(--amber);
        border-radius: 14px; padding: .75rem 1rem; box-shadow: 0 3px 12px rgba(20,32,46,.09); }
[data-testid="stMetricLabel"] { color: var(--asp5); font-weight: 600; }
[data-testid="stMetricValue"] { color: var(--asp9); }

/* ---------- tombol */
.stButton button, .stDownloadButton button { border-radius: 10px; font-weight: 700; transition: all .15s ease; }
.stDownloadButton button { background: linear-gradient(135deg, #F59E0B, #F97316); color: #14202E; border: none;
        box-shadow: 0 3px 10px rgba(245,158,11,.40); }
.stDownloadButton button:hover { background: linear-gradient(135deg, #D98206, #EA580C); color: #FFFFFF; transform: translateY(-1px); }
.stButton button { background: #FFFFFF; color: var(--asp7); border: 1.5px solid var(--asp3); }
.stButton button:hover { border-color: var(--amber); color: var(--asp9); background: #FFF7E0; }

/* ---------- input, tabel, kotak pesan */
[data-baseweb="input"]:focus-within, [data-baseweb="select"] > div:focus-within, [data-baseweb="base-input"]:focus-within
        { border-color: var(--amber) !important; box-shadow: 0 0 0 1px var(--amber) !important; }
[data-baseweb="input"], [data-baseweb="select"] > div { border-radius: 10px !important; }
[data-testid="stDataFrame"] { border-radius: 14px; overflow: hidden; box-shadow: 0 3px 14px rgba(20,32,46,.10);
        border: 1px solid #DCE6F0; }
[data-testid="stAlert"] { border-radius: 14px; border: 1px solid rgba(20,32,46,.08); box-shadow: 0 2px 8px rgba(20,32,46,.06); }
[data-testid="stFileUploaderDropzone"] { border: 2px dashed var(--asp3); border-radius: 14px; background: rgba(255,255,255,.7); }
[data-testid="stFileUploaderDropzone"]:hover { border-color: var(--amber); }
[data-testid="stPlotlyChart"] { background: #FFFFFF; border-radius: 14px; padding: 6px;
        box-shadow: 0 3px 14px rgba(20,32,46,.10); border: 1px solid #DCE6F0; }
[data-testid="stExpander"] { border-radius: 12px; background: rgba(255,255,255,.75); border: 1px solid #DCE6F0; }

/* ---------- dasbor 3D */
.dash { display: grid; grid-template-columns: repeat(5, 1fr); gap: 14px; margin: .3rem 0 1.1rem 0; perspective: 1200px; }
.kpi { position: relative; overflow: hidden; border-radius: 18px; padding: .85rem 1rem 1rem 1rem; color: #fff; min-height: 134px;
       transform: rotateX(5deg); transform-origin: 50% 100%; transition: transform .25s ease, box-shadow .25s ease;
       box-shadow: 0 1px 0 rgba(255,255,255,.35) inset, 0 8px 0 -1px rgba(0,0,0,.28), 0 18px 28px rgba(20,32,46,.30); }
.kpi::after { content: ""; position: absolute; inset: 0; pointer-events: none;
       background: linear-gradient(155deg, rgba(255,255,255,.28) 0%, rgba(255,255,255,0) 46%); }
.kpi:hover { transform: rotateX(0deg) translateY(-9px) scale(1.03);
       box-shadow: 0 1px 0 rgba(255,255,255,.4) inset, 0 12px 0 -1px rgba(0,0,0,.28), 0 28px 38px rgba(20,32,46,.38); }
.kpi .ic { position: absolute; right: .75rem; top: .55rem; font-size: 2rem; filter: drop-shadow(0 5px 4px rgba(0,0,0,.35)); }
.kpi .lb { font-size: .72rem; text-transform: uppercase; letter-spacing: .9px; font-weight: 700; opacity: .88; max-width: 75%; }
.kpi .vl { font-family: 'Poppins', 'Source Sans Pro', sans-serif; font-size: 1.22rem; font-weight: 700; line-height: 1.2; margin: 1.25rem 0 .25rem 0; }
.kpi .sb { font-size: .78rem; opacity: .92; }
.k1 { background: linear-gradient(135deg, #14202E, #34506E); }
.k2 { background: linear-gradient(135deg, #1D4E89, #3B8BD9); }
.k3 { background: linear-gradient(135deg, #0B7A75, #2BB3A6); }
.k4 { background: linear-gradient(135deg, #D97706, #F5A524); }
.k5 { background: linear-gradient(135deg, #0E7C57, #22C58B); }
.k6 { background: linear-gradient(135deg, #5B6B7D, #8CA0B5); }

.steps-wrap { background: #FFFFFF; border-radius: 18px; padding: .8rem 1.2rem 1rem 1.2rem; margin-bottom: 1.2rem;
       box-shadow: 0 8px 0 -3px #C7D6E6, 0 16px 28px rgba(20,32,46,.12); }
.steps-ttl { font-weight: 700; color: #243B53; font-size: .88rem; margin-bottom: .6rem; }
.steps { display: flex; }
.stp { flex: 1; position: relative; text-align: center; font-size: .78rem; font-weight: 600; color: #7B8CA0; padding: 0 .2rem; }
.stp .dot { width: 36px; height: 36px; margin: 0 auto .4rem auto; border-radius: 50%; display: flex; align-items: center; justify-content: center;
       background: #E6EDF5; color: #7B8CA0; font-weight: 800; box-shadow: 0 4px 0 #C7D6E6; position: relative; z-index: 2; }
.stp:not(:last-child)::after { content: ""; position: absolute; top: 16px; left: calc(50% + 24px); right: calc(-50% + 24px); height: 5px;
       border-radius: 3px; background: #E6EDF5; z-index: 1; }
.stp.done { color: #14202E; }
.stp.done .dot { background: linear-gradient(135deg, #F59E0B, #F97316); color: #fff; box-shadow: 0 4px 0 #B45309, 0 8px 12px rgba(245,158,11,.45); }
.stp.done:not(:last-child)::after { background: linear-gradient(90deg, #F59E0B, #F97316); }
.stp.now { color: #14202E; }
.stp.now .dot { background: #243B53; color: #fff; box-shadow: 0 4px 0 #14202E; animation: pulse 1.7s infinite; }
@keyframes pulse { 0% { box-shadow: 0 4px 0 #14202E, 0 0 0 0 rgba(36,59,83,.5); } 70% { box-shadow: 0 4px 0 #14202E, 0 0 0 12px rgba(36,59,83,0); } 100% { box-shadow: 0 4px 0 #14202E, 0 0 0 0 rgba(36,59,83,0); } }

/* ---------- efek 3D umum */
.stDownloadButton button { box-shadow: 0 5px 0 #B45309, 0 12px 18px rgba(245,158,11,.38); }
.stDownloadButton button:active { transform: translateY(4px); box-shadow: 0 1px 0 #B45309, 0 4px 8px rgba(245,158,11,.3); }
.stButton button { box-shadow: 0 4px 0 #B8C8D9, 0 8px 14px rgba(20,32,46,.12); }
.stButton button:active { transform: translateY(3px); box-shadow: 0 1px 0 #B8C8D9; }
[data-testid="stMetric"] { transition: transform .2s ease, box-shadow .2s ease; box-shadow: 0 6px 0 -2px #D5E1EE, 0 12px 20px rgba(20,32,46,.10); }
[data-testid="stMetric"]:hover { transform: translateY(-5px) rotateX(3deg); box-shadow: 0 10px 0 -2px #D5E1EE, 0 22px 30px rgba(20,32,46,.16); }
[data-testid="stDataFrame"], [data-testid="stPlotlyChart"] { box-shadow: 0 7px 0 -2px #D5E1EE, 0 16px 26px rgba(20,32,46,.12); }
.stTabs [aria-selected="true"], [data-testid="stTab"][aria-selected="true"] { box-shadow: 0 -3px 0 #F59E0B inset, 0 6px 14px rgba(20,32,46,.25); }

/* ---------- sidebar hidup */
.sb-brand { position: relative; overflow: hidden; display: flex; align-items: center; gap: .7rem; padding: .8rem .9rem 1rem .9rem;
        border-radius: 16px; color: #fff; margin: .2rem 0 1rem 0; background: linear-gradient(135deg, #101A26, #2A4764);
        box-shadow: 0 6px 0 -1px #0B131D, 0 14px 22px rgba(20,32,46,.32); }
.sb-brand::after { content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 5px;
        background: repeating-linear-gradient(90deg, #F59E0B 0 18px, transparent 18px 32px); animation: roadmove 2s linear infinite; }
.sb-brand svg { width: 58px; height: auto; flex: none; filter: drop-shadow(0 5px 5px rgba(0,0,0,.45)); }
.sb-brand b { font-family: 'Poppins', 'Source Sans Pro', sans-serif; font-size: 1.2rem; display: block; line-height: 1.1; }
.sb-brand small { color: #C9D8E8; font-size: .72rem; line-height: 1.25; display: block; margin-top: .2rem; }
.sb-title { display: inline-block; font-family: 'Poppins', 'Source Sans Pro', sans-serif; font-weight: 700; font-size: .92rem; color: #fff;
        background: linear-gradient(135deg, #243B53, #3E5C7E); padding: .3rem .9rem; border-radius: 999px;
        box-shadow: 0 4px 0 #14202E, 0 8px 12px rgba(20,32,46,.25); margin: .5rem 0 .8rem 0; }
.sbcard { background: #FFFFFF; border-radius: 16px; border-top: 4px solid #F59E0B; padding: .7rem .95rem .85rem .95rem; margin: .1rem 0 1rem 0;
        box-shadow: 0 6px 0 -2px #C7D6E6, 0 12px 20px rgba(20,32,46,.14); transition: transform .2s ease; }
.sbcard:hover { transform: translateY(-3px); }
.sbc-t { font-weight: 800; color: #14202E; font-size: .86rem; margin-bottom: .35rem; }
.sbrow { display: flex; justify-content: space-between; gap: .6rem; padding: .3rem 0; border-bottom: 1px dashed #DCE6F0; font-size: .79rem; color: #486581; }
.sbrow b { color: #14202E; text-align: right; }
.sbrow b.ok { color: #0E7C57; } .sbrow b.no { color: #B45309; }
.sbbar { height: 9px; background: #E6EDF5; border-radius: 6px; overflow: hidden; margin-top: .65rem; box-shadow: inset 0 1px 2px rgba(0,0,0,.18); }
.sbbar i { display: block; height: 100%; border-radius: 6px; background: linear-gradient(90deg, #F59E0B, #F97316); transition: width .4s ease; }
.sbpct { font-size: .7rem; color: #7B8CA0; text-align: right; margin-top: .2rem; font-weight: 600; }
.sbtips { background: linear-gradient(135deg, #FFF7E0, #FDE9B3); border-radius: 16px; padding: .75rem .95rem; font-size: .78rem; color: #5B3A06;
        box-shadow: 0 5px 0 -2px #F0C766, 0 10px 18px rgba(245,158,11,.22); margin-top: 1rem; }
.sbtips b { display: block; margin-bottom: .3rem; font-size: .84rem; }
.sbtips ul { margin: .2rem 0 0 1.05rem; padding: 0; } .sbtips li { margin-bottom: .2rem; }
.sbfoot { text-align: center; color: #7B8CA0; font-size: .7rem; margin: .9rem 0 .3rem 0; }
@media (max-width: 1050px) { .dash { grid-template-columns: repeat(3, 1fr); } }
@media (max-width: 820px) { .hero { grid-template-columns: 1fr; } .hero-art { display: none; } .dash { grid-template-columns: repeat(2, 1fr); }
        .stp { font-size: .66rem; } .stp .dot { width: 30px; height: 30px; } }
@media (max-width: 520px) { .dash { grid-template-columns: 1fr; } }
@media (max-width: 640px) { .hero { padding: 1rem 1.2rem 1.5rem 1.2rem; } .hero h1 { font-size: 1.7rem; }
        .stTabs [data-baseweb="tab"], [data-testid="stTab"] { padding: .45rem .7rem; } }
</style>
__HERO__
""".replace("__HERO__", _hero_html() if st.session_state.get("started") else ""), unsafe_allow_html=True)

# ------------------------------------------------------------------ halaman sambutan (3D jalan aspal)
_LANDING_CSS = """
<style>
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], [data-testid="collapsedControl"], [data-testid="stHeader"],
[data-testid="stToolbar"] { display: none !important; }
[data-testid="stMainBlockContainer"] { padding-top: 1.2rem !important; }
.lp { position: relative; height: 560px; border-radius: 24px; overflow: hidden;
      background: linear-gradient(180deg, #060B16 0%, #0F2036 28%, #2A3F66 45%, #D9803A 58%, #F9C26B 62%);
      box-shadow: 0 12px 0 -3px #0B131D, 0 30px 48px rgba(20,32,46,.5); animation: lpin .9s ease both; }
@keyframes lpin { from { opacity: 0; transform: translateY(18px) scale(.98); } to { opacity: 1; transform: none; } }
.lp-stars { position: absolute; inset: 0 0 50% 0; opacity: .9;
      background-image: radial-gradient(1.6px 1.6px at 8% 12%, #fff 50%, transparent 52%), radial-gradient(1.2px 1.2px at 22% 30%, #fff 50%, transparent 52%),
        radial-gradient(1.8px 1.8px at 35% 8%, #fff 50%, transparent 52%), radial-gradient(1.2px 1.2px at 48% 22%, #fff 50%, transparent 52%),
        radial-gradient(1.6px 1.6px at 61% 10%, #fff 50%, transparent 52%), radial-gradient(1.2px 1.2px at 74% 26%, #fff 50%, transparent 52%),
        radial-gradient(1.8px 1.8px at 86% 14%, #fff 50%, transparent 52%), radial-gradient(1.2px 1.2px at 94% 34%, #fff 50%, transparent 52%),
        radial-gradient(1.2px 1.2px at 15% 40%, #fff 50%, transparent 52%), radial-gradient(1.2px 1.2px at 68% 40%, #fff 50%, transparent 52%);
      animation: twinkle 3.2s ease-in-out infinite; }
@keyframes twinkle { 0%, 100% { opacity: .95; } 50% { opacity: .55; } }
.lp-sun { position: absolute; left: 50%; top: 62%; width: 190px; height: 190px; margin: -112px 0 0 -95px; border-radius: 50%;
      background: radial-gradient(circle, #FFF6CF 0%, #FFC857 38%, rgba(255,160,60,.55) 56%, rgba(255,160,60,0) 72%); }
.lp-mount { position: absolute; left: 0; width: 100%; top: calc(62% - 122px); height: 124px; }
.lp-ground { position: absolute; left: 0; right: 0; bottom: 0; top: 62%; background: linear-gradient(180deg, #2A4233 0%, #14231B 100%); }
.lp-roadwrap { position: absolute; left: 0; right: 0; top: 62%; bottom: 0; overflow: hidden; perspective: 300px; perspective-origin: 50% 0%; }
.lp-road { position: absolute; left: 50%; bottom: 0; width: 1500px; height: 4200px; margin-left: -750px; transform-origin: 50% 100%;
      transform: rotateX(68deg); background-color: #272C33;
      background-image: radial-gradient(rgba(255,255,255,.10) 1.2px, transparent 1.6px), radial-gradient(rgba(0,0,0,.28) 1.4px, transparent 1.8px),
        linear-gradient(90deg, #1E2329 0%, #2B3139 50%, #1E2329 100%);
      background-size: 11px 11px, 17px 17px, 100% 100%; background-position: 0 0, 6px 8px, 0 0; }
.lp-shoulder { position: absolute; top: 0; bottom: 0; width: 8%; background: linear-gradient(90deg, #4B525C, #39404A); }
.lp-shoulder.l { left: 0; } .lp-shoulder.r { right: 0; }
.lp-edge { position: absolute; top: 0; bottom: 0; width: 26px; background: #F2F5F9; box-shadow: 0 0 22px rgba(255,255,255,.55); }
.lp-edge.l { left: 11%; } .lp-edge.r { right: 11%; }
.lp-center { position: absolute; top: 0; bottom: 0; left: 50%; width: 30px; margin-left: -15px;
      background: repeating-linear-gradient(to bottom, #FBBF24 0 120px, transparent 120px 300px); animation: lprun .75s linear infinite; }
@keyframes lprun { from { background-position: 0 0; } to { background-position: 0 300px; } }
.lp-fog { position: absolute; left: 0; right: 0; top: 62%; height: 120px; pointer-events: none;
      background: linear-gradient(180deg, rgba(249,194,107,.75) 0%, rgba(249,194,107,0) 100%); }
.lp-txt { position: absolute; left: 0; right: 0; top: 6%; text-align: center; padding: 0 1rem; z-index: 3; }
.lp-hi { color: #FDE68A; letter-spacing: 5px; font-weight: 700; font-size: .95rem; animation: lpup .9s .15s ease both; }
.lp-title { font-family: 'Poppins', 'Source Sans Pro', sans-serif; font-weight: 800; font-size: 4.6rem; line-height: 1.05; color: #fff; margin: .15rem 0 .3rem 0;
      text-shadow: 0 1px 0 #dbe4ee, 0 2px 0 #c3cfdd, 0 3px 0 #aab9cb, 0 4px 0 #91a3b9, 0 5px 0 #788ca7, 0 6px 0 #5f7594, 0 7px 0 #475f81,
                   0 8px 0 #34496a, 0 22px 26px rgba(0,0,0,.55);
      animation: lpup .9s .3s ease both, lpfloat 4.5s 1.2s ease-in-out infinite; }
.lp-sub { color: #E4ECF6; font-size: 1.05rem; max-width: 640px; margin: 0 auto .8rem auto; text-shadow: 0 2px 6px rgba(0,0,0,.5); animation: lpup .9s .5s ease both; }
.lp-chips span { display: inline-block; margin: 0 .3rem .3rem .3rem; padding: .15rem .85rem; border-radius: 999px; font-size: .82rem; font-weight: 700; color: #FDE68A;
      background: rgba(10,20,34,.55); border: 1px solid rgba(253,230,138,.55); animation: lpup .9s .7s ease both; }
@keyframes lpup { from { opacity: 0; transform: translateY(22px); } to { opacity: 1; transform: none; } }
@keyframes lpfloat { 0%, 100% { transform: translateY(0); } 50% { transform: translateY(-7px); } }
[data-testid="stHorizontalBlock"] { margin-top: -92px; position: relative; z-index: 10; }
[data-testid="stBaseButton-primary"], .stButton button[kind="primary"] { font-size: 1.45rem; font-weight: 800; padding: .85rem 1.2rem; border-radius: 18px; border: none;
      background: linear-gradient(135deg, #FBBF24, #F97316); color: #14202E;
      box-shadow: 0 8px 0 #B45309, 0 20px 30px rgba(249,115,22,.5); animation: lpglow 1.8s ease-in-out infinite; }
[data-testid="stBaseButton-primary"]:hover, .stButton button[kind="primary"]:hover { background: linear-gradient(135deg, #F59E0B, #EA580C); color: #fff; transform: translateY(-3px); }
[data-testid="stBaseButton-primary"]:active, .stButton button[kind="primary"]:active { transform: translateY(6px); box-shadow: 0 2px 0 #B45309, 0 6px 10px rgba(249,115,22,.4); }
[data-testid="stBaseButton-primary"] p, .stButton button[kind="primary"] p { font-size: 1.55rem !important; font-weight: 800 !important; letter-spacing: 1px; }
@keyframes lpglow { 0%, 100% { box-shadow: 0 8px 0 #B45309, 0 18px 28px rgba(249,115,22,.40); } 50% { box-shadow: 0 8px 0 #B45309, 0 24px 44px rgba(251,191,36,.75); } }
.lp-note { text-align: center; color: #486581; font-size: .85rem; margin-top: .9rem; font-weight: 600; }
@media (max-width: 700px) { .lp { height: 500px; } .lp-title { font-size: 2.7rem; } .lp-sub { font-size: .9rem; } .lp-hi { letter-spacing: 3px; font-size: .78rem; } }
</style>
<div class="lp">
  <div class="lp-stars"></div>
  <div class="lp-sun"></div>
  <svg class="lp-mount" viewBox="0 0 1000 124" preserveAspectRatio="none"><path d="M0,124 L0,80 L90,42 L170,84 L260,32 L360,90 L450,54 L560,94 L650,38 L760,88 L850,48 L940,82 L1000,60 L1000,124 Z" fill="#0F1B2B"/><path d="M0,124 L0,98 L120,72 L220,100 L330,68 L450,102 L580,74 L700,104 L820,72 L920,100 L1000,82 L1000,124 Z" fill="#17304A"/></svg>
  <div class="lp-ground"></div>
  <div class="lp-roadwrap"><div class="lp-road"><div class="lp-shoulder l"></div><div class="lp-shoulder r"></div><div class="lp-edge l"></div><div class="lp-edge r"></div><div class="lp-center"></div></div></div>
  <div class="lp-fog"></div>
  <div class="lp-txt"><div class="lp-hi">SELAMAT DATANG DI APLIKASI</div><div class="lp-title">GradAspal</div>
  <div class="lp-sub">Rancang gradasi agregat, hitung kadar aspal, dan temukan kadar aspal optimum (KAO) secara otomatis.</div>
  <div class="lp-chips"><span>Latasir</span><span>Lataston</span><span>Laston</span><span>Uji Marshall</span><span>Ekspor Excel</span></div></div>
</div>
"""
if not st.session_state["started"]:
    st.markdown(_LANDING_CSS, unsafe_allow_html=True)
    _c1, _c2, _c3 = st.columns([1, 1.15, 1])
    with _c2:
        if st.button("🚧  Mulai", key="btn_mulai", type="primary", use_container_width=True):
            st.session_state["started"] = True
            st.rerun()
    st.markdown('<div class="lp-note">Gratis · tanpa login · data tidak disimpan di server</div>', unsafe_allow_html=True)
    st.stop()



# ---------------------------------------------------------------- sidebar (tanpa login)
ver = st.session_state.get("ver", 0)
sv = st.session_state.get("arsip") or {}
tag = f"app_{ver}"
if not st.session_state.get("visited"):
    log_visit("anonim", "buka aplikasi")
    st.session_state["visited"] = True
st.sidebar.markdown(_brand_html(), unsafe_allow_html=True)
side_sum = st.sidebar.container()
st.sidebar.caption("Langsung pakai, tanpa login. Data tidak disimpan di server. Simpan pekerjaanmu lewat **Arsip pekerjaan** di bawah.")

st.sidebar.markdown('<div class="sb-title">📁 Arsip pekerjaan</div>', unsafe_allow_html=True)
if st.session_state.get("flash"):
    st.sidebar.success(st.session_state.pop("flash"))
up = st.sidebar.file_uploader("Buka arsip (.json)", type=["json"], key=f"up_{ver}",
                              help="Pilih file arsip yang sebelumnya kamu unduh dari aplikasi ini.")
if up is not None:
    try:
        data = json.loads(up.getvalue().decode("utf-8"))
        if not isinstance(data, dict) or data.get("app") != "GradAspal":
            raise ValueError
        st.session_state["arsip"] = data
        st.session_state["ver"] = ver + 1
        st.session_state["flash"] = f"Arsip dimuat: {data.get('judul') or up.name}"
        st.rerun()
    except (ValueError, UnicodeDecodeError):
        st.sidebar.error("File bukan arsip GradAspal yang valid.")
if st.sidebar.button("Mulai baru (kosongkan semua isian)"):
    st.session_state["arsip"] = {}
    st.session_state["ver"] = ver + 1
    st.rerun()
judul = st.sidebar.text_input("Judul arsip (opsional)", value=str(sv.get("judul", "")),
                              placeholder="Contoh: Kelompok 3 - Laston AC-WC", key=f"judul_{tag}")
st.sidebar.markdown('<div class="sb-title">⚙️ Pengaturan campuran</div>', unsafe_allow_html=True)
mix_list = list(SPEC)
mix0 = sv.get("mix") if sv.get("mix") in SPEC else "Laston AC-WC"
mix = st.sidebar.selectbox("Jenis campuran", mix_list, index=mix_list.index(mix0), key=f"mix_{tag}")
W_saved = sv.get("W")
W_in = st.sidebar.number_input("Berat campuran (gr)", min_value=100.0, max_value=10000.0,
                               value=float(W_saved) if W_saved else None, step=50.0,
                               placeholder="Isi berat campuran, mis. 1200", key=f"W_{tag}")
W = float(W_in) if W_in else 0.0
fam = mix.split()[0]
k0 = float(sv["K"]) if sv.get("fam") == fam and "K" in sv else K_DEF[fam]
K = st.sidebar.number_input("Konstanta K", 0.0, 5.0, k0, 0.05, key=f"K_{tag}_{fam}")
lo_k, hi_k = K_RNG[fam]
st.sidebar.caption(f"Rentang K {fam}: {lo_k:g} - {hi_k:g}")
if not lo_k <= K <= hi_k:
    st.sidebar.warning("K di luar rentang yang lazim.")
sp = SPEC[mix]

dash = st.container()
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["1. Analisis Saringan", "2. Gradasi Rencana",
                                              "3. Kadar Aspal, Berat & Koreksi", "4. Uji Marshall & KAO",
                                              "5. Pusat Bantuan & Panduan", "6. Riwayat Akses (Admin)"])

# ------------------------------------------------------------------ tab 1
with tab1:
    st.subheader("Input massa tertahan tiap saringan")
    st.caption("Kolom massa dikosongkan di awal. Klik sel, ketik angka (gram), lalu tekan Enter. Sel kosong dianggap 0.")
    saved_m = sv.get("massa", [])
    if len(saved_m) == len(LAB) + 1 and any(v > 0 for v in saved_m):
        init_m = [float(v) if v > 0 else np.nan for v in saved_m]
    else:
        init_m = [np.nan] * (len(LAB) + 1)
    base = pd.DataFrame({"Saringan (mm)": [f"{s:g}" for s in LAB] + ["Pan"],
                         "Massa Tertahan (gr)": init_m})
    ed = st.data_editor(base, hide_index=True, disabled=["Saringan (mm)"], width="stretch",
                        column_config={"Massa Tertahan (gr)": st.column_config.NumberColumn(
                            "Massa Tertahan (gr)", min_value=0.0, format="%.1f")},
                        key=f"massa_{tag}")
    m = ed["Massa Tertahan (gr)"].fillna(0).astype(float).values
    tot = m.sum()
    if tot <= 0:
        st.warning("Isi massa tertahan terlebih dahulu. Butuh petunjuk? Buka tab 5 (Pusat Bantuan & Panduan).")
        df1 = None
    else:
        cum = np.cumsum(m)
        pk = cum / tot * 100
        lw = 100 - pk
        mhb = pk[:-1].sum() / 100
        df1 = pd.DataFrame({"Saringan (mm)": list(base["Saringan (mm)"]),
                            "Massa Tertahan (gr)": m, "Jumlah Tertahan (gr)": cum,
                            "% Kumulatif Tertahan": pk, "% Lewat": lw})
        df1.loc[len(df1)] = ["Jumlah", tot, cum.sum(), pk[:-1].sum(), np.nan]
        st.dataframe(df1.round(2), hide_index=True, width="stretch")

        c1, c2 = st.columns(2)
        c1.metric("MHB", f"{mhb:.2f}")
        c2.info("MHB campuran gabungan, ditampilkan sebagai informasi. Syarat MHB pasir (1,5 - 3,8) dan "
                "kerikil (5,0 - 8,0) berlaku untuk agregat yang diuji sendiri, bukan untuk campuran.")

        rows, tips = [], []
        for i, sz in enumerate(LAB):
            if sz in sp:
                lo, hi = sp[sz]
                v = lw[i]
                ok = lo - 0.005 <= v <= hi + 0.005
                sel = 0.0 if ok else (v - hi if v > hi else v - lo)
                rows.append([f"{sz:g}", lo, hi, round(v, 2), round(sel, 2), "Masuk batas" if ok else "Di luar batas"])
                if not ok:
                    tips.append(saran(sz, v, lo, hi))
        if rows:
            st.subheader(f"Pengecekan terhadap batas {mix}")
            st.dataframe(pd.DataFrame(rows, columns=["Saringan (mm)", "BB", "BA", "% Lolos Sampel",
                                                     "Selisih ke batas (%)", "Status"]),
                         hide_index=True, width="stretch")
            st.subheader("Saran perbaikan gradasi")
            if tips:
                st.warning(f"{len(tips)} saringan di luar batas. Berikut saran penyesuaian proporsi fraksi agregat:")
                for t in tips:
                    st.markdown(f"- {t}")
                st.caption("Saran bersifat panduan umum. Setelah proporsi diubah, timbang ulang, ayak, lalu isi kembali massa tertahan untuk memeriksa hasilnya.")
            else:
                st.success("Semua saringan yang diatur spesifikasi sudah masuk batas. Tidak ada penyesuaian yang diperlukan.")
        show_fig(plot(mix, LAB, lw[:-1]), "grafik_saringan")

# ------------------------------------------------------------------ tab 2
mids, prev = [], 100.0
for s in S:
    if s in sp:
        prev = sum(sp[s]) / 2
    mids.append(prev)
mids = np.array(mids)
ret = np.r_[100.0, mids[:-1]] - mids  # saringan tanpa batas -> tertahan 0
with tab2:
    st.subheader(f"Gradasi rencana {mix} (nilai tengah)")
    df2 = pd.DataFrame({"Saringan (mm)": [f"{s:g}" for s in S],
                        "BB": [sp[s][0] if s in sp else np.nan for s in S],
                        "BA": [sp[s][1] if s in sp else np.nan for s in S],
                        "% Lolos (Nilai Tengah)": mids, "% Tertahan": ret})
    df2.loc[len(df2)] = ["Pan", np.nan, np.nan, 0.0, mids[-1]]
    df2.loc[len(df2)] = ["Jumlah", np.nan, np.nan, np.nan, ret.sum() + mids[-1]]
    st.dataframe(df2.round(2), hide_index=True, width="stretch")
    st.caption("Saringan tanpa batas di tabel spesifikasi mengikuti nilai tengah sebelumnya, sehingga % tertahannya 0.")
    show_fig(plot(mix, mids=True), "grafik_gradasi_rencana")

# ------------------------------------------------------------------ tab 3
with tab3:
    p236, p075 = mids[S.index(2.36)], mids[S.index(0.075)]
    AK, AH, F = 100 - p236, p236 - p075, p075
    Pb = 0.035 * AK + 0.045 * AH + 0.18 * F + K
    Pbr = float(np.floor(Pb * 2 + 0.5) / 2)

    st.subheader("Kadar aspal estimasi")
    st.info("Ini **kadar aspal estimasi (Pb)**, yaitu perkiraan awal dari rumus, **bukan kadar aspal optimum**. "
            "Kadar optimum (KAO) ditentukan lewat uji Marshall di **Tab 4**, memakai variasi kadar aspal pada tabel di bawah.")
    a, b, c, d = st.columns(4)
    a.metric("AK (tertahan 2,36)", f"{AK:.2f}%")
    b.metric("AH (lolos 2,36 - lolos 0,075)", f"{AH:.2f}%")
    c.metric("F (lolos 0,075)", f"{F:.2f}%")
    d.metric("Pb estimasi", f"{Pb:.3f}%", f"dibulatkan {Pbr:g}%", delta_color="off")
    st.code(f"Pb = 0,035({AK:.2f}) + 0,045({AH:.2f}) + 0,18({F:.2f}) + {K:g} = {Pb:.3f}%  ->  {Pbr:g}%")

    kadar0 = float(sv["kadar"]) if sv.get("kadar_key") == f"{mix}|{K}" else Pbr
    kadar = st.number_input("Kadar aspal yang dipakai (%)", 0.0, 20.0, kadar0, 0.1, key=f"kadar_{tag}_{mix}_{K}")
    if W <= 0:
        st.info("Isi **berat campuran (gr)** di sidebar agar berat aspal dan agregat bisa dihitung.")
    Wa = kadar / 100 * W
    Wg = W - Wa
    e, f_, g = st.columns(3)
    e.metric("Berat campuran", f"{W:.1f} gr")
    f_.metric("Berat aspal", f"{Wa:.2f} gr")
    g.metric("Berat agregat", f"{Wg:.2f} gr")

    st.subheader("Berat tertahan tiap saringan")
    rp = np.r_[ret, mids[-1]]
    df3 = pd.DataFrame({"Saringan (mm)": [f"{s:g}" for s in S] + ["Pan"],
                        "% Tertahan": rp, "Berat Tertahan (gr)": rp / 100 * Wg})
    df3.loc[len(df3)] = ["Jumlah", rp.sum(), rp.sum() / 100 * Wg]
    st.dataframe(df3.round(2), hide_index=True, width="stretch")

    st.subheader("Koreksi terhadap kadar aspal")
    fk = 1 - kadar / 100
    sebelum = [AK, AH, F, np.nan, AK + AH + F]
    koreksi = [AK * fk, AH * fk, F * fk, kadar]
    koreksi.append(sum(koreksi))
    df4 = pd.DataFrame({"Komponen": ["Agregat kasar (AK)", "Agregat halus (AH)", "Filler (F)", "Aspal", "Total"],
                        "Sebelum Koreksi (%)": sebelum, "Setelah Koreksi (%)": koreksi,
                        "Berat (gr)": [k / 100 * W for k in koreksi]})
    st.dataframe(df4.round(2), hide_index=True, width="stretch")
    st.caption("Koreksi: persentase agregat dikalikan (1 - kadar aspal), lalu ditambah aspal sehingga total 100%.")

    st.subheader("Variasi kadar aspal untuk benda uji Marshall")
    st.caption("Lima kadar aspal di sekitar kadar yang dipakai (selisih 0,5%). Timbang bahan sesuai tabel untuk membuat benda uji, "
               "lalu tentukan kadar optimum dari hasil uji Marshall. Berat campuran per benda uji = berat campuran di sidebar.")
    kadar_list = [round(kadar + d, 2) for d in (-1.0, -0.5, 0.0, 0.5, 1.0)]
    kadar_list = [k for k in kadar_list if 0 < k < 100]
    labels = ["Pb -1", "Pb -0,5", "Pb", "Pb +0,5", "Pb +1"][-len(kadar_list):] if len(kadar_list) == 5 else [f"{k:g}%" for k in kadar_list]
    dv1 = pd.DataFrame({"Variasi": labels, "Kadar Aspal (%)": kadar_list,
                        "Berat Aspal (gr)": [k / 100 * W for k in kadar_list],
                        "Berat Agregat (gr)": [W - k / 100 * W for k in kadar_list]})
    st.dataframe(dv1.round(2), hide_index=True, width="stretch")
    dv2 = pd.DataFrame({"Saringan (mm)": [f"{x:g}" for x in S] + ["Pan"]})
    for k in kadar_list:
        dv2[f"{k:g}% (gr)"] = rp / 100 * (W - k / 100 * W)
    dv2.loc[len(dv2)] = ["Jumlah"] + [dv2[c].sum() for c in dv2.columns[1:]]
    st.markdown("**Berat agregat tertahan tiap saringan untuk tiap kadar aspal**")
    st.dataframe(dv2.round(2), hide_index=True, width="stretch")

    xbuf = xlsx_gradasi(mix, W, K, kadar, Pbr, m, kadar_list, labels)
    st.download_button("Unduh hasil (Excel)", xbuf, "hasil_gradaspal.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("File Excel berisi rumus yang saling terhubung dan grafik. Ubah sel kuning (massa tertahan, berat campuran, K, kadar aspal) dan semua hasil ikut dihitung ulang.")

# ------------------------------------------------------------------ tab 4: Marshall
with tab4:
    st.subheader("Perhitungan Uji Marshall dan Kadar Aspal Optimum (KAO)")
    st.caption("Semua isian kosong di awal. Isi data hasil pengujian benda uji, lalu aplikasi menghitung stabilitas, flow, MQ, "
               "VMA, VIM, VFA, menampilkan grafik, dan menentukan KAO.")
    msv = sv.get("marshall") or {}
    mp = msv.get("params") or {}

    def pnum(label, key, help_=None):
        v = mp.get(key)
        return st.number_input(label, min_value=0.0, value=float(v) if v else None, step=0.001, format="%.4f",
                               placeholder="Isi", help=help_, key=f"m_{key}_{tag}")

    st.markdown("**1. Berat jenis bahan**")
    raw = dict(mp.get("raw") or {})
    _has_direct = any(mp.get(k_) for k_ in ("kasar_bulk", "halus_bulk", "filler_bulk", "kasar_semu", "halus_semu", "filler_semu", "aspal"))
    MODE_H, MODE_L = "Hitung otomatis dari berat hasil pengujian", "Isi langsung nilai berat jenis"
    _def_h = mp.get("bj_mode") == "hitung" or (mp.get("bj_mode") is None and not _has_direct)
    hitung = st.radio("Cara mengisi berat jenis", [MODE_H, MODE_L], horizontal=True, index=0 if _def_h else 1,
                      key=f"m_bjmode_{tag}") == MODE_H

    def rnum(label, key, help_=None, default=None):
        v = raw.get(key, default)
        return st.number_input(label, min_value=0.0, value=float(v) if v else None, step=0.01, format="%.2f",
                               placeholder="Isi", help=help_, key=f"r_{key}_{tag}")

    if hitung:
        st.caption("Isi berat hasil pengujian di laboratorium (satuan gram). Berat jenis bulk, SSD, semu, dan penyerapan "
                   "terhitung otomatis, lalu dipakai untuk VMA, VIM, dan VFA.")
        cK, cH = st.columns(2)
        with cK:
            st.markdown("**Agregat kasar** (SNI 1969)")
            k_bk = rnum("Berat benda uji kering oven, Bk (gr)", "k_bk")
            k_bj = rnum("Berat benda uji SSD di udara, Bj (gr)", "k_bj")
            k_ba = rnum("Berat benda uji SSD dalam air, Ba (gr)", "k_ba")
        with cH:
            st.markdown("**Agregat halus** (SNI 1970)")
            h_bk = rnum("Berat benda uji kering oven, Bk (gr)", "h_bk")
            h_b = rnum("Berat piknometer + air, B (gr)", "h_b")
            h_bt = rnum("Berat piknometer + benda uji + air, Bt (gr)", "h_bt")
            h_s = rnum("Berat benda uji SSD, S (gr)", "h_s", "Biasanya 500 gr.", default=500)
        cF, cA = st.columns(2)
        with cF:
            st.markdown("**Filler** (piknometer)")
            f_bk = rnum("Berat filler kering, Bk (gr)", "f_bk")
            f_b = rnum("Berat piknometer + air, B (gr)", "f_b")
            f_bt = rnum("Berat piknometer + filler + air, Bt (gr)", "f_bt")
        with cA:
            st.markdown("**Aspal** (SNI 2441)")
            a_a = rnum("Berat piknometer kosong, A (gr)", "a_a")
            a_b = rnum("Berat piknometer + air, B (gr)", "a_b")
            a_c = rnum("Berat piknometer + aspal, C (gr)", "a_c")
            a_d = rnum("Berat piknometer + aspal + air, D (gr)", "a_d")
        raw = {"k_bk": k_bk, "k_bj": k_bj, "k_ba": k_ba, "h_bk": h_bk, "h_b": h_b, "h_bt": h_bt, "h_s": h_s,
               "f_bk": f_bk, "f_b": f_b, "f_bt": f_bt, "a_a": a_a, "a_b": a_b, "a_c": a_c, "a_d": a_d}
        rk, rh = bj_kasar(k_bk, k_bj, k_ba), bj_halus(h_bk, h_b, h_bt, h_s)
        rf, ra_ = bj_filler(f_bk, f_b, f_bt), bj_aspal(a_a, a_b, a_c, a_d)
        for nm, vals, res, tip in [
            ("Agregat kasar", (k_bk, k_bj, k_ba), rk, "Berat SSD harus lebih besar dari berat dalam air, dan berat kering harus lebih besar dari berat dalam air."),
            ("Agregat halus", (h_bk, h_b, h_bt, h_s), rh, "Periksa berat piknometer. Berat B + S harus lebih besar dari Bt."),
            ("Filler", (f_bk, f_b, f_bt), rf, "Periksa berat piknometer. Berat B + Bk harus lebih besar dari Bt."),
            ("Aspal", (a_a, a_b, a_c, a_d), ra_, "Periksa urutan berat: A < C, dan D harus sesuai (piknometer + aspal + air)."),
        ]:
            if any(vals) and res is None and all(vals):
                st.warning(f"{nm}: angka belum masuk akal. {tip}")
        bulk_k, semu_k = (rk["bulk"], rk["semu"]) if rk else (None, None)
        bulk_h, semu_h = (rh["bulk"], rh["semu"]) if rh else (None, None)
        bulk_f = semu_f = rf
        bj_asp = ra_
        _f = lambda v_, nd=4: "-" if v_ is None else f"{v_:.{nd}f}"
        _t = pd.DataFrame([
            ["Agregat kasar", _f(rk and rk["bulk"]), _f(rk and rk["ssd"]), _f(rk and rk["semu"]), _f(rk and rk["serap"], 2)],
            ["Agregat halus", _f(rh and rh["bulk"]), _f(rh and rh["ssd"]), _f(rh and rh["semu"]), _f(rh and rh["serap"], 2)],
            ["Filler", _f(rf), "-", _f(rf), "-"],
            ["Aspal", "-", "-", _f(ra_), "-"],
        ], columns=["Bahan", "BJ bulk", "BJ SSD", "BJ semu (aspal: BJ)", "Penyerapan (%)"])
        st.markdown("**Hasil perhitungan berat jenis (otomatis)**")
        st.dataframe(_t, hide_index=True, width="stretch")
        st.caption("Filler dan aspal hanya punya satu nilai berat jenis. Untuk filler, nilai itu dipakai sebagai BJ bulk dan BJ semu.")
        if rk and rk["bulk"] > rk["semu"] or rh and rh["bulk"] > rh["semu"]:
            st.warning("BJ bulk tidak mungkin lebih besar dari BJ semu. Periksa kembali berat yang kamu isi.")
        r3 = st.columns(3)
        with r3[0]:
            kal_in = pnum("Faktor kalibrasi alat (per skala pembacaan)", "kal",
                          "Faktor proving ring dari alat Marshall di laboratoriummu.")
        with r3[1]:
            unit = st.radio("Satuan faktor kalibrasi", ["kg", "lbs"], horizontal=True,
                            index=1 if mp.get("unit") == "lbs" else 0, key=f"m_unit_{tag}",
                            help="Jika faktor di modulmu dalam lbs, pilih lbs. Aplikasi mengubahnya ke kg (x 0,45359).")
    else:
        r1 = st.columns(3)
        with r1[0]:
            bulk_k = pnum("BJ bulk agregat kasar", "kasar_bulk")
        with r1[1]:
            bulk_h = pnum("BJ bulk agregat halus", "halus_bulk")
        with r1[2]:
            bulk_f = pnum("BJ bulk filler", "filler_bulk")
        r2 = st.columns(3)
        with r2[0]:
            semu_k = pnum("BJ semu agregat kasar", "kasar_semu")
        with r2[1]:
            semu_h = pnum("BJ semu agregat halus", "halus_semu")
        with r2[2]:
            semu_f = pnum("BJ semu filler", "filler_semu")
        r3 = st.columns(3)
        with r3[0]:
            bj_asp = pnum("BJ aspal", "aspal")
        with r3[1]:
            kal_in = pnum("Faktor kalibrasi alat (per skala pembacaan)", "kal",
                          "Faktor proving ring dari alat Marshall di laboratoriummu.")
        with r3[2]:
            unit = st.radio("Satuan faktor kalibrasi", ["kg", "lbs"], horizontal=True,
                            index=1 if mp.get("unit") == "lbs" else 0, key=f"m_unit_{tag}",
                            help="Jika faktor di modulmu dalam lbs, pilih lbs. Aplikasi mengubahnya ke kg (x 0,45359).")
    kal_kg = (kal_in * 0.45359237 if unit == "lbs" else kal_in) if kal_in else None
    if kal_kg:
        st.caption(f"Faktor kalibrasi yang dipakai: {kal_kg:.4f} kg per skala.")

    p475 = mids[S.index(4.75)]
    pCA, pFA, pFF = 100 - p475, p475 - mids[S.index(0.075)], mids[S.index(0.075)]
    st.caption(f"Proporsi agregat dari gradasi rencana {mix} (terhadap total agregat): "
               f"kasar (tertahan 4,75 mm) {pCA:.2f}%, halus {pFA:.2f}%, filler {pFF:.2f}%.")
    bj_ok = all(v for v in (bulk_k, bulk_h, bulk_f, semu_k, semu_h, semu_f, bj_asp))
    Gsb = Gse = Gsa = None
    if bj_ok:
        Gsb = 100 / (pCA / bulk_k + pFA / bulk_h + pFF / bulk_f)
        Gsa = 100 / (pCA / semu_k + pFA / semu_h + pFF / semu_f)
        Gse = 100 / (pCA / ((bulk_k + semu_k) / 2) + pFA / ((bulk_h + semu_h) / 2) + pFF / ((bulk_f + semu_f) / 2))
        g1, g2, g3 = st.columns(3)
        g1.metric("BJ bulk agregat gabungan", f"{Gsb:.4f}")
        g2.metric("BJ semu agregat gabungan", f"{Gsa:.4f}")
        g3.metric("BJ efektif agregat gabungan", f"{Gse:.4f}")
    else:
        st.info("Lengkapi berat hasil pengujian berat jenis (agregat kasar, halus, filler, dan aspal) agar VMA, VIM, dan VFA bisa dihitung." if hitung else "Isi semua berat jenis (bulk, semu, aspal) agar VMA, VIM, dan VFA bisa dihitung.")

    st.markdown("**2. Syarat spesifikasi**")
    base_sp = MSPEC.get(mix, {})
    sdf = pd.DataFrame({"Parameter": MPARAMS,
                        "Min": [(base_sp.get(p) or (None, None))[0] if base_sp.get(p) and base_sp[p][0] is not None else np.nan for p in MPARAMS],
                        "Maks": [(base_sp.get(p) or (None, None))[1] if base_sp.get(p) and base_sp[p][1] is not None else np.nan for p in MPARAMS]})
    if base_sp:
        st.caption(f"Syarat {mix} sesuai Spesifikasi Umum 2018 (Tabel 12.2 laporan). Boleh diubah jika modulmu berbeda. "
                   "MQ tidak ada di Tabel 12.2, jadi tidak dibatasi (isi min 250 jika modulmu mensyaratkan). "
                   f"Jumlah tumbukan per bidang: {'112' if mix == 'Laston AC-Base' else '75'}. "
                   "Stabilitas Marshall sisa (min 90%) dan rongga refusal (min 2%) tidak dihitung aplikasi ini.")
    else:
        st.caption(f"Syarat untuk {mix} belum terisi otomatis. Isi sendiri sesuai spesifikasi yang kamu pakai.")
    sed = st.data_editor(sdf, hide_index=True, disabled=["Parameter"], width="stretch", key=f"ms_{tag}_{mix}",
                         column_config={"Min": st.column_config.NumberColumn("Min", format="%.2f"),
                                        "Maks": st.column_config.NumberColumn("Maks", format="%.2f")})
    specs = {r_["Parameter"]: (r_["Min"], r_["Maks"]) for _, r_ in sed.iterrows()}

    st.markdown("**3. Data benda uji**")
    kor_opts = ["Dari volume benda uji", "Dari tebal benda uji (mm)"]
    kor_mode = st.radio("Penentuan angka korelasi stabilitas", kor_opts, horizontal=True,
                        index=1 if mp.get("kor_mode") == "tebal" else 0, key=f"m_kor_{tag}",
                        help="Volume = berat SSD - berat dalam air (otomatis). Tebal = hasil pengukuran tebal benda uji dengan jangka sorong. "
                             "Keduanya memakai tabel SNI 06-2489-1991 dengan interpolasi.")
    kor_tebal = kor_mode == kor_opts[1]
    st.caption("Satu baris per benda uji. Boleh lebih dari satu benda uji per kadar aspal (hasilnya dirata-rata). "
               "Tambah baris lewat tombol + di bawah tabel. Kadar aspal terhadap campuran. "
               + ("Isi kolom Tebal Benda Uji (mm) untuk tiap benda uji. " if kor_tebal else "Kolom Tebal Benda Uji boleh dikosongkan. ")
               + "Kolom Angka Korelasi (opsional) dikosongkan agar dihitung otomatis; jika diisi, angka itu yang dipakai.")
    C_TB, C_KOR = "Tebal Benda Uji (mm)", "Angka Korelasi (opsional)"
    cols_m = ["Kadar Aspal (%)", "Berat Kering (gr)", "Berat SSD (gr)", "Berat dalam Air (gr)",
              "Pembacaan Stabilitas", "Flow (mm)", C_TB, C_KOR]
    saved_rows = msv.get("rows") or []
    if saved_rows:
        base_m = pd.DataFrame(saved_rows, columns=cols_m).astype(float)
    else:
        base_m = pd.DataFrame({c_: [np.nan] * 5 for c_ in cols_m})
    ed_m = st.data_editor(base_m, num_rows="dynamic", hide_index=True, width="stretch", key=f"mr_{tag}",
                          column_config={c_: st.column_config.NumberColumn(c_, min_value=0.0, format="%.3f") for c_ in cols_m})

    recs, warn_vol, warn_tb, kosong_tb, beda_kor, kor_eff = [], [], [], [], [], {}
    for i_, rw in ed_m.reset_index(drop=True).iterrows():
        pbv, bk_, bs_, ba_ = rw[cols_m[0]], rw[cols_m[1]], rw[cols_m[2]], rw[cols_m[3]]
        if any(pd.isna(x) for x in (pbv, bk_, bs_, ba_)) or bs_ - ba_ <= 0 or not 0 < pbv < 100:
            continue
        vol = bs_ - ba_
        gmb = bk_ / vol
        rec = {"No": i_ + 1, "Kadar Aspal (%)": float(pbv), "Berat Kering (gr)": float(bk_), "Berat dalam Air (gr)": float(ba_),
               "Berat SSD (gr)": float(bs_), "Volume (cm3)": vol, "BJ Bulk Campuran": gmb}
        if bj_ok:
            ps = 100 - pbv
            gmm = 100 / (ps / Gse + pbv / bj_asp)
            vma = 100 - gmb * ps / Gsb
            vim = 100 * (gmm - gmb) / gmm
            rec.update({"BJ Teoritis Maks": gmm, "VMA (%)": vma, "VIM (%)": vim, "VFA (%)": 100 * (vma - vim) / vma})
        if not 200 <= vol <= 625:
            warn_vol.append(i_ + 1)
        tb_ = rw[C_TB]
        has_tb = not pd.isna(tb_) and tb_ > 0
        f_vol = f_korelasi(vol)
        f_tb = f_korelasi_tebal(tb_) if has_tb else None
        if has_tb:
            rec["Tebal (mm)"] = float(tb_)
            if not 25.4 - 0.05 <= tb_ <= 76.2 + 0.05:
                warn_tb.append(i_ + 1)
            if abs(f_tb - f_vol) > 0.1:
                beda_kor.append(i_ + 1)
        if not pd.isna(rw[C_KOR]):
            fk, dasar = float(rw[C_KOR]), "Manual"
        elif kor_tebal and has_tb:
            fk, dasar = f_tb, "Tebal"
            kor_eff[i_] = f_tb
        else:
            fk, dasar = f_vol, "Volume"
            if kor_tebal:
                kosong_tb.append(i_ + 1)
        rec["Angka Korelasi"] = fk
        rec["Dasar Korelasi"] = dasar
        rd_, fl_ = rw[cols_m[4]], rw[cols_m[5]]
        if not pd.isna(rd_):
            rec["Pembacaan"] = float(rd_)
        if kal_kg and not pd.isna(rd_):
            rec["Stabilitas (kg)"] = rd_ * kal_kg * fk
            if not pd.isna(fl_) and fl_ > 0:
                rec["Flow (mm)"] = fl_
                rec["MQ (kg/mm)"] = rec["Stabilitas (kg)"] / fl_
        elif not pd.isna(fl_):
            rec["Flow (mm)"] = fl_
        recs.append(rec)

    # data untuk Excel: kolom angka korelasi diisi dari tebal bila memilih tebal (7 kolom, tanpa kolom tebal)
    x_rows = ed_m.drop(columns=[C_TB]).reset_index(drop=True)
    for i_, v_ in kor_eff.items():
        if pd.isna(x_rows.at[i_, C_KOR]):
            x_rows.at[i_, C_KOR] = v_

    ra = None
    kao = None
    rm = pd.DataFrame(recs)
    if rm.empty:
        st.info("Isi data benda uji (kadar aspal, berat kering, berat SSD, berat dalam air) untuk melihat hasil.")
    else:
        for pcol in MPARAMS:
            if pcol not in rm.columns:
                rm[pcol] = np.nan
        if warn_vol:
            st.warning(f"Volume benda uji nomor {', '.join(map(str, warn_vol))} di luar rentang tabel korelasi (200-625 cm3). "
                       "Cek kembali berat SSD dan berat dalam air.")
        if warn_tb:
            st.warning(f"Tebal benda uji nomor {', '.join(map(str, warn_tb))} di luar rentang tabel korelasi (25,4-76,2 mm). "
                       "Cek kembali pengukuran tebal (satuan mm).")
        if kosong_tb:
            st.info(f"Tebal benda uji nomor {', '.join(map(str, kosong_tb))} belum diisi, jadi angka korelasinya memakai volume.")
        if beda_kor:
            st.warning(f"Pada benda uji nomor {', '.join(map(str, beda_kor))}, angka korelasi dari tebal dan dari volume berbeda lebih dari 0,1. "
                       "Cek kembali pengukuran tebal atau penimbangan (berat SSD dan berat dalam air).")
        if not kal_kg:
            st.info("Isi faktor kalibrasi alat agar stabilitas dan MQ dihitung.")
        st.markdown("**Hasil per benda uji**")
        tampil = st.radio("Tampilan hasil", ["Ringkas", "Format lembar kerja"], horizontal=True, key=f"m_view_{tag}",
                          help="Format lembar kerja = kolom F sampai Q seperti tabel data Marshall di praktikum, lengkap dengan baris rata-rata per kadar aspal.")
        if tampil == "Ringkas":
            show_cols = [c_ for c_ in ["No", "Kadar Aspal (%)", "Volume (cm3)", "BJ Bulk Campuran", "BJ Teoritis Maks", "VMA (%)", "VIM (%)",
                                       "VFA (%)", "Tebal (mm)", "Angka Korelasi", "Dasar Korelasi", "Stabilitas (kg)", "Flow (mm)",
                                       "MQ (kg/mm)"] if c_ in rm.columns]
            st.dataframe(rm[show_cols].round(3), hide_index=True, width="stretch")
        else:
            lk = tabel_lembar_kerja(rm)
            st.caption("Kolom F, G, H diisi hasil timbangan. I = H - G, J = F / I. K, L, M = VIM, VMA, VFB. "
                       "N = bacaan arloji, O = N x kalibrasi x angka korelasi, P = flow, Q = O / P.")
            try:
                sty = lk.style.format(precision=3, na_rep="-").apply(
                    lambda r_: ["font-weight: 700; background-color: #FFF1CC" if r_["Benda uji"] == "rata-rata" else "" for _ in r_], axis=1)
                st.dataframe(sty, hide_index=True, width="stretch")
            except Exception:
                st.dataframe(lk.round(3), hide_index=True, width="stretch")
            lbuf = io.BytesIO()
            with pd.ExcelWriter(lbuf, engine="openpyxl") as xw_:
                lk.round(4).to_excel(xw_, sheet_name="Format Lembar Kerja", index=False)
            st.download_button("Unduh tabel lembar kerja (Excel)", lbuf.getvalue(), "tabel_lembar_kerja_marshall.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_lk")

        ra = rm.groupby("Kadar Aspal (%)", as_index=False).agg(
            **{"Jumlah Benda Uji": ("No", "count")}, **{p_: (p_, "mean") for p_ in MPARAMS})
        st.markdown("**Rata-rata per kadar aspal**")
        st.dataframe(ra.round(3), hide_index=True, width="stretch")

        def lolos(p_, v_):
            lo_, hi_ = specs.get(p_, (np.nan, np.nan))
            if pd.isna(lo_) and pd.isna(hi_):
                return None
            if pd.isna(v_):
                return False
            if not pd.isna(lo_) and v_ < lo_ - 1e-9:
                return False
            if not pd.isna(hi_) and v_ > hi_ + 1e-9:
                return False
            return True

        stat_rows, pass_flags = [], []
        for _, rw in ra.iterrows():
            oks = [lolos(p_, rw[p_]) for p_ in MPARAMS]
            pass_flags.append(any(o is not None for o in oks) and all(o in (True, None) for o in oks))
            stat_rows.append([rw["Kadar Aspal (%)"]] + ["-" if o is None else ("Memenuhi" if o else "Tidak") for o in oks]
                             + ["Ya" if pass_flags[-1] else "Tidak"])
        st.markdown("**Pengecekan terhadap syarat**")
        st.dataframe(pd.DataFrame(stat_rows, columns=["Kadar Aspal (%)"] + MPARAMS + ["Semua memenuhi"]),
                     hide_index=True, width="stretch")

        cand = ra[pd.Series(pass_flags, index=ra.index)]
        st.subheader("Kadar aspal optimum (KAO)")
        if all(pd.isna(lo_) and pd.isna(hi_) for lo_, hi_ in specs.values()):
            st.info("Isi syarat spesifikasi (langkah 2) agar KAO bisa ditentukan.")
        elif cand.empty:
            st.warning("KAO belum dapat ditentukan: belum ada kadar aspal yang memenuhi semua syarat spesifikasi. "
                       "Lihat tabel pengecekan untuk mengetahui parameter mana yang belum terpenuhi.")
        else:
            if cand["Stabilitas (kg)"].notna().any():
                kao = float(cand.loc[cand["Stabilitas (kg)"].idxmax(), "Kadar Aspal (%)"])
            else:
                kao = float(cand["Kadar Aspal (%)"].median())
            st.success(f"KAO = {kao:g}%  (kadar yang memenuhi semua syarat dengan stabilitas tertinggi). "
                       f"Kadar yang memenuhi semua syarat: {', '.join(f'{x:g}%' for x in cand['Kadar Aspal (%)'])}.")
            st.caption("Cara menentukan KAO bisa berbeda antar modul atau dosen. Di sini mengikuti metode laporan: "
                       "kadar yang memenuhi semua syarat dengan stabilitas tertinggi.")
        show_fig(plot_marshall(ra, specs, kao), "grafik_marshall")

        mbuf = xlsx_marshall(mix, specs, {"kasar_bulk": bulk_k, "halus_bulk": bulk_h, "filler_bulk": bulk_f,
                                          "kasar_semu": semu_k, "halus_semu": semu_h, "filler_semu": semu_f,
                                          "aspal": bj_asp}, kal_in, unit, (pCA, pFA, pFF), x_rows,
                             list(ra["Kadar Aspal (%)"]), raw if hitung else None)
        st.download_button("Unduh hasil Marshall (Excel)", mbuf, "hasil_marshall.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_marshall")

mparams = {"kasar_bulk": bulk_k, "halus_bulk": bulk_h, "filler_bulk": bulk_f, "kasar_semu": semu_k,
           "halus_semu": semu_h, "filler_semu": semu_f, "aspal": bj_asp, "kal": kal_in, "unit": unit,
           "bj_mode": "hitung" if hitung else "langsung", "kor_mode": "tebal" if kor_tebal else "volume", "raw": {k_: (None if v_ is None else float(v_)) for k_, v_ in raw.items()}}
mrows = [{k_: (None if pd.isna(v_) else float(v_)) for k_, v_ in r_.items()} for r_ in ed_m.to_dict("records")]

# ------------------------------------------- arsip pekerjaan (unduh)
arsip_out = {"app": "GradAspal", "versi": 1, "judul": judul.strip(), "mix": mix, "fam": fam,
             "W": float(W), "K": float(K), "massa": [float(v) for v in m], "kadar": float(kadar),
             "kadar_key": f"{mix}|{K}", "diperbarui": now_wib(),
             "marshall": {"params": mparams, "rows": mrows}}
slug = re.sub(r"[^a-z0-9]+", "_", (judul.strip() or mix).lower()).strip("_") or "arsip"
st.sidebar.download_button("💾 Simpan arsip pekerjaan", json.dumps(arsip_out, ensure_ascii=False, indent=1).encode("utf-8"),
                           f"gradaspal_{slug}_{datetime.now(timezone(timedelta(hours=7))).strftime('%Y%m%d')}.json",
                           "application/json", key="dl_arsip",
                           help="Unduh file arsip. Buka lagi kapan saja lewat 'Buka arsip' untuk melanjutkan pekerjaan.")
if sv.get("diperbarui"):
    st.sidebar.caption(f"Arsip dibuka, terakhir disimpan: {sv['diperbarui']}")

# ------------------------------------------------------------------ tab 5: bantuan
with tab5:
    st.subheader("Pusat Bantuan & Panduan Penggunaan")
    p1, p2, p3 = st.tabs(["Panduan Langkah demi Langkah", "Tanya Jawab (FAQ)", "Istilah & Rumus"])

    with p1:
        st.markdown("""
**Langkah 1 - Buka aplikasi**
- Tidak perlu daftar atau login. Langsung pakai.
- Pada HP, buka sidebar lewat ikon panah di kiri atas untuk mengisi jenis campuran dan berat campuran.

**Langkah 2 - Atur campuran di sidebar**
- Pilih **Jenis campuran** (Latasir, Lataston, atau Laston beserta kelasnya).
- Isi **Berat campuran (gr)**, misalnya 1200.
- **Konstanta K** sudah terisi nilai umum. Ubah hanya jika diminta asisten/dosen. Rentang lazim ditampilkan di bawah kolom K.

**Langkah 3 - Tab 1: Analisis Saringan**
- Isi **massa tertahan** (gram) tiap saringan dari hasil praktikum, termasuk baris **Pan**.
- Sel yang dikosongkan dianggap 0.
- Hasilnya: % kumulatif tertahan, % lewat, nilai **MHB**, pengecekan terhadap batas spesifikasi (BB/BA), dan grafik.

**Langkah 4 - Tab 2: Gradasi Rencana**
- Terisi otomatis dari nilai tengah spesifikasi campuran yang kamu pilih. Tidak perlu input.

**Langkah 5 - Tab 3: Kadar Aspal, Berat & Koreksi**
- Aplikasi menghitung **Pb estimasi** dan membulatkannya ke 0,5% terdekat.
- Kamu boleh mengganti **kadar aspal yang dipakai**.
- Berat aspal, berat agregat, berat tertahan tiap saringan, dan koreksi dihitung otomatis.
- Tabel **Variasi kadar aspal** memberi berat bahan untuk lima kadar (Pb -1 sampai Pb +1) untuk membuat benda uji Marshall.

**Langkah 5b - Tab 4: Uji Marshall & KAO**
- Isi berat hasil pengujian berat jenis (aplikasi menghitung BJ bulk, SSD, semu, dan penyerapan otomatis), atau isi nilai BJ langsung, lalu isi faktor kalibrasi alat.
- Isi data tiap benda uji: kadar aspal, berat kering, berat SSD, berat dalam air, pembacaan stabilitas, flow, dan (jika diukur) tebal benda uji.
- Pilih penentuan **angka korelasi**: dari volume (otomatis) atau dari tebal benda uji (mm).
- Hasil bisa dilihat dalam tampilan **Ringkas** atau **Format lembar kerja** (kolom F sampai Q dengan baris rata-rata).
- Aplikasi menghitung VMA, VIM, VFA, stabilitas, flow, dan MQ, lalu mengecek syarat dan menentukan **KAO**.

**Langkah 6 - Simpan hasil**
- Klik **ikon kamera** di pojok kanan atas grafik untuk mengunduh gambar (PNG), dan **Unduh hasil (Excel)** untuk semua tabel dan grafik beserta rumusnya.
- Untuk melanjutkan pekerjaan nanti, klik **Simpan arsip pekerjaan** di sidebar. Ini mengunduh file `.json`.
- Kapan saja, buka aplikasi lagi, lalu pilih file itu di **Buka arsip** (sidebar). Semua isian kembali seperti semula.
""")

    with p2:
        with st.expander("Bagaimana cara menyimpan dan melanjutkan pekerjaan?"):
            st.write("Klik **Simpan arsip pekerjaan** di sidebar untuk mengunduh file .json. Nanti, pilih file itu di "
                     "**Buka arsip** untuk melanjutkan. Pekerjaan tidak tersimpan otomatis di server, jadi simpan arsipnya "
                     "sebelum menutup halaman.")
        with st.expander("Arsipnya hilang atau gagal dibuka"):
            st.write("Pastikan memakai file .json yang diunduh dari GradAspal dan belum diubah isinya. "
                     "File arsip ada di folder Unduhan perangkatmu.")
        with st.expander("Kenapa tabel dan grafik di Tab 1 belum muncul?"):
            st.write("Massa tertahan masih kosong. Isi minimal satu sel dengan angka lebih dari 0.")
        with st.expander("Apa arti 'Di luar batas' dan bagaimana memperbaikinya?"):
            st.write("Persen lolos sampelmu pada saringan itu berada di luar rentang BB-BA campuran yang dipilih. "
                     "Lihat bagian **Saran perbaikan gradasi** di bawah tabel pengecekan di Tab 1. Di sana ada penjelasan "
                     "fraksi mana yang perlu ditambah atau dikurangi. Kolom **Selisih ke batas** menunjukkan seberapa jauh "
                     "nilainya dari batas (positif = melebihi BA, negatif = kurang dari BB).")
        with st.expander("Kenapa MHB tidak diberi status memenuhi atau tidak?"):
            st.write("Tab 1 menganalisis campuran agregat gabungan, sehingga MHB hanya ditampilkan sebagai informasi. "
                     "Syarat MHB pasir 1,5-3,8 dan kerikil 5,0-8,0 berlaku untuk agregat yang diuji sendiri. "
                     "Kalau hasilmu terasa janggal, periksa urutan saringan dan angka massa (satuan gram).")
        with st.expander("Apakah hasil aplikasi ini kadar aspal optimum?"):
            st.write("Bukan. Aplikasi menghitung kadar aspal estimasi (Pb) dari rumus sebagai titik awal. "
                     "Kadar optimum ditentukan dari uji Marshall pada benda uji dengan lima variasi kadar aspal "
                     "(Pb -1, -0,5, Pb, +0,5, +1). Tabel berat bahan untuk kelima variasi itu ada di Tab 3, dan perhitungan Marshall serta penentuan KAO ada di Tab 4.")
        with st.expander("Faktor kalibrasi alat Marshall: kg atau lbs?"):
            st.write("Faktor kalibrasi (proving ring) bisa tertulis dalam lbs. Pilih satuannya di Tab 4 dan aplikasi akan "
                     "mengubahnya ke kg (1 lbs = 0,45359 kg). Stabilitas dihitung = pembacaan x faktor kalibrasi (kg) x angka korelasi.")
        with st.expander("Dari mana angka korelasi stabilitas? Volume atau tebal?"):
            st.write("Ada dua pilihan di Tab 4. (1) Dari volume: volume = berat SSD - berat dalam air, lalu dicocokkan ke tabel SNI 06-2489-1991. "
                     "(2) Dari tebal: isi kolom Tebal Benda Uji (mm) hasil pengukuran jangka sorong, lalu dicocokkan ke tabel yang sama. "
                     "Keduanya memakai interpolasi. Jika modulmu memberi angka lain, atau tidak mengalikan angka korelasi, isi kolom "
                     "Angka Korelasi (opsional); isi 1 bila tidak dikoreksi.")
        with st.expander("Kenapa KAO belum muncul?"):
            st.write("KAO hanya muncul jika ada kadar aspal yang memenuhi semua syarat (VIM, VMA, VFA, stabilitas, flow, MQ). "
                     "Lihat tabel Pengecekan terhadap syarat di Tab 4 untuk tahu parameter mana yang belum terpenuhi.")
        with st.expander("Apa itu konstanta K dan berapa nilainya?"):
            st.write("K adalah konstanta yang bergantung jenis campuran dan dipakai di rumus Pb. "
                     "Latasir dan Lataston: 2,0-3,0. Laston: 0,5-1,0. Nilai awal aplikasi: 2,5 / 2,5 / 0,75.")
        with st.expander("Apakah datanya bisa dilihat orang lain?"):
            st.write("Tidak. Isianmu tidak disimpan di server. Arsip hanya ada di file milikmu sendiri. "
                     "Admin hanya melihat jumlah dan waktu kunjungan.")
        with st.expander("Aplikasi tidak tampil rapi di HP"):
            st.write("Buka sidebar lewat ikon panah di kiri atas. Putar HP ke mode lanskap agar tabel lebih lebar. "
                     "Tabel bisa digeser ke samping.")
        with st.expander("Hasil hitunganku beda dengan hitungan manual"):
            st.write("Cek pembulatan: aplikasi menghitung dengan presisi penuh dan baru membulatkan saat ditampilkan. "
                     "Pb estimasi dibulatkan ke 0,5% terdekat.")

    with p3:
        st.markdown("""
| Istilah | Arti |
|---|---|
| **BB / BA** | Batas bawah / batas atas persen lolos menurut spesifikasi |
| **% Lewat (lolos)** | 100 dikurangi persen kumulatif tertahan |
| **MHB** | Modulus Halus Butir: jumlah % kumulatif tertahan dibagi 100 |
| **AK** | Agregat kasar: persen tertahan saringan 2,36 mm |
| **AH** | Agregat halus: persen lolos 2,36 mm dikurangi persen lolos 0,075 mm |
| **F** | Filler: persen lolos saringan 0,075 mm |
| **Pb** | Kadar aspal estimasi (perkiraan awal dari rumus, bukan kadar optimum) |
| **Kadar aspal optimum** | Kadar aspal yang memenuhi semua syarat spesifikasi berdasarkan uji Marshall |
| **Uji Marshall** | Pengujian benda uji aspal dengan beberapa variasi kadar aspal (stabilitas, flow, VIM, VMA, VFA) untuk menentukan kadar optimum |
| **Gmb / Gsb / Gse / Gmm** | BJ bulk campuran / BJ bulk agregat gabungan / BJ efektif agregat gabungan / BJ teoritis maksimum campuran |
| **VIM / VMA / VFA** | Rongga dalam campuran / rongga dalam agregat / rongga terisi aspal |
| **MQ** | Marshall Quotient: stabilitas dibagi flow |
| **KAO** | Kadar aspal optimum |
| **Pan** | Wadah penampung butir yang lolos saringan terkecil |

**Rumus kadar aspal estimasi**
""")
        st.code("Pb = 0,035(AK) + 0,045(AH) + 0,18(F) + K")
        st.markdown("""
**Rumus koreksi:** persentase tiap agregat dikalikan (1 - kadar aspal), lalu aspal ditambahkan sehingga total 100%.

**Berat aspal** = kadar aspal x berat campuran. **Berat agregat** = berat campuran - berat aspal.
""")

    st.divider()
    st.markdown("**Masih butuh bantuan?** Hubungi pengelola aplikasi (isi kontak di sini, misalnya WhatsApp atau email dosen/admin).")

# ------------------------------------------------------------------ tab 6: admin
with tab6:
    st.subheader("Riwayat akses pengguna")
    kode = st.text_input("Kode admin", type="password", key="admin_code")
    if kode and kode == ADMIN_CODE:
        if ADMIN_CODE == "gradaspal-admin":
            st.warning("Kode admin masih bawaan. Ganti lewat variabel GRADASPAL_ADMIN (atau Secrets) sebelum dipublikasikan.")
        if os.path.exists(VISITS):
            dv = pd.read_csv(VISITS)
            a1, a2 = st.columns(2)
            a1.metric("Total kunjungan", len(dv))
            a2.metric("IP unik (perkiraan pengunjung)", dv["ip"].fillna("").replace("", np.nan).nunique())
            st.dataframe(dv.iloc[::-1], hide_index=True, width="stretch")
            st.download_button("Unduh riwayat (CSV)", dv.to_csv(index=False).encode("utf-8"),
                               "riwayat_akses.csv", "text/csv")
        else:
            st.info("Belum ada riwayat akses.")
    elif kode:
        st.error("Kode admin salah.")
    else:
        st.caption("Tab ini hanya untuk pemilik aplikasi. Masukkan kode admin untuk melihat riwayat.")

# ------------------------------------------------------------------ dasbor ringkasan (diisi setelah semua perhitungan)
ok_n = ok_c = 0
if tot > 0:
    for _i, _sz in enumerate(LAB):
        if _sz in sp:
            _lo, _hi = sp[_sz]
            ok_c += 1
            ok_n += int(_lo - 0.005 <= lw[_i] <= _hi + 0.005)
with dash:
    st.markdown(_dash_html(mix, K, W, tot, (mhb if tot > 0 else 0.0), ok_n, ok_c, Pb, Pbr, kadar, len(rm), kao),
                unsafe_allow_html=True)
with side_sum:
    st.markdown(_side_html(mix, W, Pb, kadar, tot, ok_n, ok_c, len(rm), kao), unsafe_allow_html=True)
st.sidebar.markdown(
    '<div class="sbtips"><b>💡 Tips cepat</b><ul><li>Isi berat campuran dulu.</li>'
    '<li>Isi massa saringan di Tab 1, lalu data Marshall di Tab 4.</li>'
    '<li>Simpan arsip sebelum menutup halaman.</li><li>Bingung? Buka Tab 5.</li></ul></div>'
    '<div class="sbfoot">GradAspal · Praktikum Perkerasan Lentur</div>', unsafe_allow_html=True)
