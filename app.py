"""GradAspal - Perencanaan Gradasi Agregat & Kadar Aspal Estimasi
Latasir, Lataston, Laston.  Jalankan: streamlit run app.py
"""
import csv
import hashlib
import io
import json
import os
import re
from datetime import datetime, timedelta, timezone

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

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

# Syarat Marshall Laston (Spesifikasi Umum 2018): (min, maks); None = tidak dibatasi
_MQ = (250.0, None)
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


def plot(mix, xs=None, ys=None, mids=False):
    sp = SPEC[mix]
    x = [s for s in S if s in sp]
    fig, ax = plt.subplots(figsize=(9, 5))
    fig.patch.set_facecolor("#FBF5EC")
    ax.set_facecolor("#FFFDF9")
    lo_v, hi_v = [sp[s][0] for s in x], [sp[s][1] for s in x]
    ax.fill_between(x, lo_v, hi_v, color="#E6CCB2", alpha=0.55, label="Daerah spesifikasi")
    ax.plot(x, lo_v, "o-", color="#5C3317", lw=2, label="BB")
    ax.plot(x, hi_v, "o-", color="#B08968", lw=2, label="BA")
    if mids:
        ax.plot(x, [sum(sp[s]) / 2 for s in x], "d--", color="#8B5A2B", lw=2, label="Nilai Tengah")
    if xs is not None:
        ax.plot(xs, ys, "o-", color="#E07B39", lw=2.5, label="Presentase Butir Lolos (%)")
    ax.set_xscale("log")
    ax.set_xlim(100, 0.01)
    ax.set_ylim(0, 120)
    ax.set_xticks([100, 10, 1, 0.1, 0.01])
    ax.set_xticklabels(["100", "10", "1", "0,1", "0,01"])
    ax.grid(True, which="major", alpha=0.4)
    ax.set_xlabel("Ukuran Lubang Saringan (mm)")
    ax.set_ylabel("Presentase Lolos Kumulatif (%)")
    ax.set_title("Hasil Perbandingan Presentase Lolos Kumulatif Terhadap Persen "
                 f"Berat Butir yang Lolos pada {mix}", fontsize=10, wrap=True)
    ax.legend(loc="lower center", ncol=4, bbox_to_anchor=(0.5, -0.28), frameon=False)
    fig.tight_layout()
    return fig


def plot_marshall(ra, specs, kao=None):
    fig, axes = plt.subplots(2, 3, figsize=(12, 7))
    fig.patch.set_facecolor("#FBF5EC")
    for ax, p in zip(axes.ravel(), ["Stabilitas (kg)", "Flow (mm)", "MQ (kg/mm)", "VMA (%)", "VIM (%)", "VFA (%)"]):
        ax.set_facecolor("#FFFDF9")
        d = ra[["Kadar Aspal (%)", p]].dropna()
        if not d.empty:
            ax.plot(d["Kadar Aspal (%)"], d[p], "o-", color="#E07B39", lw=2, label="Hasil uji")
        lo, hi = specs.get(p, (np.nan, np.nan))
        if not pd.isna(lo):
            ax.axhline(lo, color="#5C3317", ls="--", lw=1.5, label="Syarat min")
        if not pd.isna(hi):
            ax.axhline(hi, color="#B08968", ls="--", lw=1.5, label="Syarat maks")
        if kao is not None:
            ax.axvline(kao, color="#2E7D32", ls=":", lw=1.8, label="KAO")
        ax.set_title(p, fontsize=10)
        ax.set_xlabel("Kadar Aspal (%)")
        ax.grid(True, alpha=0.4)
        if ax.get_legend_handles_labels()[0]:
            ax.legend(fontsize=7)
    fig.tight_layout()
    return fig


def show_fig(fig, name):
    st.pyplot(fig)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    st.download_button("Unduh grafik (PNG)", buf.getvalue(), f"{name}.png", "image/png", key=name)
    plt.close(fig)


st.set_page_config(page_title="GradAspal", page_icon="🛣️", layout="wide")
st.markdown("""
<style>
.stApp { background: linear-gradient(180deg, #FBF5EC 0%, #F1E3D0 100%); }
[data-testid="stSidebar"] { background: linear-gradient(180deg, #EAD7C0 0%, #D2B48C 100%); }
.hero { background: linear-gradient(135deg, #3E2415 0%, #6F4E37 55%, #A9744F 100%);
        padding: 1.6rem 2rem; border-radius: 16px; margin-bottom: 1.2rem;
        box-shadow: 0 6px 18px rgba(62,36,21,.28); }
.hero h1 { color: #FFF3E3; margin: 0; font-size: 2.3rem; letter-spacing: .5px; }
.hero p { color: #EBD5BA; margin: .3rem 0 0 0; font-size: 1rem; }
[data-testid="stMetric"] { background: rgba(255,255,255,.65); border-left: 6px solid #8B5A2B;
        border-radius: 12px; padding: .7rem 1rem; box-shadow: 0 2px 8px rgba(62,36,21,.12); }
.stTabs [data-baseweb="tab"] { font-weight: 600; }
.stDownloadButton button { background: linear-gradient(135deg, #6F4E37, #A9744F); color: #fff;
        border: none; border-radius: 10px; }
.stDownloadButton button:hover { background: linear-gradient(135deg, #5C3317, #8B5A2B); color: #fff; }
@media (max-width: 640px) { .hero { padding: 1rem 1.2rem; } .hero h1 { font-size: 1.7rem; } }
</style>
<div class="hero">
  <h1>🛣️ GradAspal</h1>
  <p>Perencanaan gradasi agregat dan kadar aspal estimasi untuk Latasir, Lataston, dan Laston</p>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------- sidebar (tanpa login)
ver = st.session_state.get("ver", 0)
sv = st.session_state.get("arsip") or {}
tag = f"app_{ver}"
if not st.session_state.get("visited"):
    log_visit("anonim", "buka aplikasi")
    st.session_state["visited"] = True
st.sidebar.caption("Langsung pakai, tanpa login. Data tidak disimpan di server. Simpan pekerjaanmu lewat **Arsip pekerjaan** di bawah.")

st.sidebar.markdown("### 📁 Arsip pekerjaan")
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
st.sidebar.divider()
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
        if 1.5 <= mhb <= 3.8:
            c2.success("MHB memenuhi syarat agregat halus/pasir (1,5 - 3,8)")
        elif 5.0 <= mhb <= 8.0:
            c2.success("MHB memenuhi syarat agregat kasar/kerikil (5,0 - 8,0)")
        else:
            c2.error("MHB di luar syarat pasir (1,5-3,8) dan kerikil (5,0-8,0)")

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
            "Kadar optimum ditentukan lewat uji Marshall dengan beberapa variasi kadar aspal (lihat tabel variasi di bawah).")
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

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        if df1 is not None:
            df1.round(2).to_excel(xw, sheet_name="Analisis Saringan", index=False)
        df2.round(2).to_excel(xw, sheet_name="Gradasi Rencana", index=False)
        pd.DataFrame({"Parameter": ["Jenis campuran", "K", "AK", "AH", "F", "Pb estimasi", "Kadar dipakai",
                                    "Berat campuran", "Berat aspal", "Berat agregat"],
                      "Nilai": [mix, K, round(AK, 2), round(AH, 2), round(F, 2), round(Pb, 3), kadar,
                                W, round(Wa, 2), round(Wg, 2)]}
                     ).to_excel(xw, sheet_name="Kadar Aspal", index=False)
        df3.round(2).to_excel(xw, sheet_name="Berat per Saringan", index=False)
        df4.round(2).to_excel(xw, sheet_name="Koreksi", index=False)
        dv1.round(2).to_excel(xw, sheet_name="Variasi Kadar Aspal", index=False)
        dv2.round(2).to_excel(xw, sheet_name="Variasi Berat Saringan", index=False)
    st.download_button("Unduh hasil (Excel)", buf.getvalue(), "hasil_gradaspal.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
        st.info("Isi semua berat jenis (bulk, semu, aspal) agar VMA, VIM, dan VFA bisa dihitung.")

    st.markdown("**2. Syarat spesifikasi**")
    base_sp = MSPEC.get(mix, {})
    sdf = pd.DataFrame({"Parameter": MPARAMS,
                        "Min": [(base_sp.get(p) or (None, None))[0] if base_sp.get(p) and base_sp[p][0] is not None else np.nan for p in MPARAMS],
                        "Maks": [(base_sp.get(p) or (None, None))[1] if base_sp.get(p) and base_sp[p][1] is not None else np.nan for p in MPARAMS]})
    if base_sp:
        st.caption(f"Syarat {mix} sesuai Spesifikasi Umum 2018 (Tabel 12.1 laporan). Boleh diubah jika modulmu berbeda.")
    else:
        st.caption(f"Syarat untuk {mix} belum terisi otomatis. Isi sendiri sesuai spesifikasi yang kamu pakai.")
    sed = st.data_editor(sdf, hide_index=True, disabled=["Parameter"], width="stretch", key=f"ms_{tag}_{mix}",
                         column_config={"Min": st.column_config.NumberColumn("Min", format="%.2f"),
                                        "Maks": st.column_config.NumberColumn("Maks", format="%.2f")})
    specs = {r_["Parameter"]: (r_["Min"], r_["Maks"]) for _, r_ in sed.iterrows()}

    st.markdown("**3. Data benda uji**")
    st.caption("Satu baris per benda uji. Boleh lebih dari satu benda uji per kadar aspal (hasilnya dirata-rata). "
               "Tambah baris lewat tombol + di bawah tabel. Kadar aspal terhadap campuran. Kosongkan angka korelasi agar dihitung otomatis dari volume.")
    cols_m = ["Kadar Aspal (%)", "Berat Kering (gr)", "Berat SSD (gr)", "Berat dalam Air (gr)",
              "Pembacaan Stabilitas", "Flow (mm)", "Angka Korelasi (opsional)"]
    saved_rows = msv.get("rows") or []
    if saved_rows:
        base_m = pd.DataFrame(saved_rows, columns=cols_m).astype(float)
    else:
        base_m = pd.DataFrame({c_: [np.nan] * 5 for c_ in cols_m})
    ed_m = st.data_editor(base_m, num_rows="dynamic", hide_index=True, width="stretch", key=f"mr_{tag}",
                          column_config={c_: st.column_config.NumberColumn(c_, min_value=0.0, format="%.3f") for c_ in cols_m})

    recs, warn_vol = [], []
    for i_, rw in ed_m.reset_index(drop=True).iterrows():
        pbv, bk_, bs_, ba_ = rw[cols_m[0]], rw[cols_m[1]], rw[cols_m[2]], rw[cols_m[3]]
        if any(pd.isna(x) for x in (pbv, bk_, bs_, ba_)) or bs_ - ba_ <= 0 or not 0 < pbv < 100:
            continue
        vol = bs_ - ba_
        gmb = bk_ / vol
        rec = {"No": i_ + 1, "Kadar Aspal (%)": float(pbv), "Volume (cm3)": vol, "BJ Bulk Campuran": gmb}
        if bj_ok:
            ps = 100 - pbv
            gmm = 100 / (ps / Gse + pbv / bj_asp)
            vma = 100 - gmb * ps / Gsb
            vim = 100 * (gmm - gmb) / gmm
            rec.update({"BJ Teoritis Maks": gmm, "VMA (%)": vma, "VIM (%)": vim, "VFA (%)": 100 * (vma - vim) / vma})
        if not 200 <= vol <= 625:
            warn_vol.append(i_ + 1)
        fk = rw[cols_m[6]] if not pd.isna(rw[cols_m[6]]) else f_korelasi(vol)
        rec["Angka Korelasi"] = fk
        rd_, fl_ = rw[cols_m[4]], rw[cols_m[5]]
        if kal_kg and not pd.isna(rd_):
            rec["Stabilitas (kg)"] = rd_ * kal_kg * fk
            if not pd.isna(fl_) and fl_ > 0:
                rec["Flow (mm)"] = fl_
                rec["MQ (kg/mm)"] = rec["Stabilitas (kg)"] / fl_
        elif not pd.isna(fl_):
            rec["Flow (mm)"] = fl_
        recs.append(rec)

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
        if not kal_kg:
            st.info("Isi faktor kalibrasi alat agar stabilitas dan MQ dihitung.")
        st.markdown("**Hasil per benda uji**")
        show_cols = [c_ for c_ in ["No", "Kadar Aspal (%)", "Volume (cm3)", "BJ Bulk Campuran", "BJ Teoritis Maks", "VMA (%)", "VIM (%)",
                                   "VFA (%)", "Angka Korelasi", "Stabilitas (kg)", "Flow (mm)", "MQ (kg/mm)"] if c_ in rm.columns]
        st.dataframe(rm[show_cols].round(3), hide_index=True, width="stretch")

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

        mbuf = io.BytesIO()
        with pd.ExcelWriter(mbuf, engine="openpyxl") as xw:
            rm[show_cols].round(4).to_excel(xw, sheet_name="Per Benda Uji", index=False)
            ra.round(4).to_excel(xw, sheet_name="Rata-rata", index=False)
            pd.DataFrame(stat_rows, columns=["Kadar Aspal (%)"] + MPARAMS + ["Semua memenuhi"]).to_excel(
                xw, sheet_name="Pengecekan Syarat", index=False)
        st.download_button("Unduh hasil Marshall (Excel)", mbuf.getvalue(), "hasil_marshall.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_marshall")

mparams = {"kasar_bulk": bulk_k, "halus_bulk": bulk_h, "filler_bulk": bulk_f, "kasar_semu": semu_k,
           "halus_semu": semu_h, "filler_semu": semu_f, "aspal": bj_asp, "kal": kal_in, "unit": unit}
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
- Isi berat jenis bahan (bulk, semu, aspal) dan faktor kalibrasi alat.
- Isi data tiap benda uji: kadar aspal, berat kering, berat SSD, berat dalam air, pembacaan stabilitas, dan flow.
- Aplikasi menghitung VMA, VIM, VFA, stabilitas, flow, dan MQ, lalu mengecek syarat dan menentukan **KAO**.

**Langkah 6 - Simpan hasil**
- Klik **Unduh grafik (PNG)** untuk gambar dan **Unduh hasil (Excel)** untuk semua tabel.
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
        with st.expander("MHB-ku di luar syarat. Salah di mana?"):
            st.write("Periksa urutan saringan dan angka massa (satuan gram). Syarat MHB: pasir 1,5-3,8 dan kerikil 5,0-8,0. "
                     "Jika sampelmu campuran, nilainya bisa di antara keduanya, sehingga ditandai di luar syarat.")
        with st.expander("Apakah hasil aplikasi ini kadar aspal optimum?"):
            st.write("Bukan. Aplikasi menghitung kadar aspal estimasi (Pb) dari rumus sebagai titik awal. "
                     "Kadar optimum ditentukan dari uji Marshall pada benda uji dengan lima variasi kadar aspal "
                     "(Pb -1, -0,5, Pb, +0,5, +1). Tabel berat bahan untuk kelima variasi itu ada di Tab 3, dan perhitungan Marshall serta penentuan KAO ada di Tab 4.")
        with st.expander("Faktor kalibrasi alat Marshall: kg atau lbs?"):
            st.write("Faktor kalibrasi (proving ring) bisa tertulis dalam lbs. Pilih satuannya di Tab 4 dan aplikasi akan "
                     "mengubahnya ke kg (1 lbs = 0,45359 kg). Stabilitas dihitung = pembacaan x faktor kalibrasi (kg) x angka korelasi.")
        with st.expander("Dari mana angka korelasi stabilitas?"):
            st.write("Otomatis dari volume benda uji (berat SSD - berat dalam air) memakai tabel SNI 06-2489-1991 dengan "
                     "interpolasi. Jika modulmu memberi angka lain, isi kolom Angka Korelasi (opsional) di tabel benda uji.")
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
