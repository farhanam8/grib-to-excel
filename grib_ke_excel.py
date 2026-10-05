"""
GRIB ke Excel
=============
Aplikasi kecil untuk mengubah file GRIB (ERA5, ERA5-Land, dan sejenisnya)
menjadi file Excel pada satu titik lokasi.

Cara jalan:
    pip install -r requirements.txt
    streamlit run grib_ke_excel.py

Alur:
    1. Pengguna memasukkan koordinat (lintang, bujur).
    2. Aplikasi mencari titik grid terdekat di setiap dataset dalam file GRIB.
    3. Variabel mentah diterjemahkan ke besaran yang mudah dibaca,
       misalnya u10 + v10 menjadi kecepatan dan arah angin.
    4. Hasilnya diunduh sebagai Excel (sheet Data, Ringkasan, Info).
"""

from __future__ import annotations

import io
import os
import tempfile
import warnings
from dataclasses import dataclass, field
from datetime import datetime

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=FutureWarning)

R_BUMI_KM = 6371.0088
NAMA_ARAH = ["Utara", "Timur Laut", "Timur", "Tenggara",
             "Selatan", "Barat Daya", "Barat", "Barat Laut"]
ZONA_WAKTU = {"UTC": 0, "WIB (UTC+7)": 7, "WITA (UTC+8)": 8, "WIT (UTC+9)": 9}
DIM_WAKTU = {"time", "step", "valid_time"}
DIM_RUANG = {"latitude", "longitude"}

# --------------------------------------------------------------------------
# Kamus terjemahan variabel
# --------------------------------------------------------------------------
# Pasangan komponen angin: (u, v, keterangan ketinggian)
PASANGAN_ANGIN = [
    ("u10", "v10", "10 m"),
    ("u100", "v100", "100 m"),
    ("u10n", "v10n", "10 m netral"),
    ("u", "v", ""),          # level tekanan, label level ditambahkan otomatis
]


def _k2c(x):
    return x - 273.15


def _kali(f):
    return lambda x: x * f


# nama_pendek: (nama tampil, satuan, fungsi konversi, keterangan, jenis)
# jenis: "biasa", "jumlah" (boleh dijumlah, akumulasi), "fluks" (akumulasi
# yang dirata-ratakan), "arah" (derajat).
KONVERSI = {
    "t2m":  ("Suhu udara 2 m", "°C", _k2c, "K - 273,15", "biasa"),
    "d2m":  ("Suhu titik embun 2 m", "°C", _k2c, "K - 273,15", "biasa"),
    "skt":  ("Suhu kulit permukaan", "°C", _k2c, "K - 273,15", "biasa"),
    "sst":  ("Suhu permukaan laut", "°C", _k2c, "K - 273,15", "biasa"),
    "t":    ("Suhu udara", "°C", _k2c, "K - 273,15", "biasa"),
    "mx2t": ("Suhu maksimum 2 m", "°C", _k2c, "K - 273,15", "biasa"),
    "mn2t": ("Suhu minimum 2 m", "°C", _k2c, "K - 273,15", "biasa"),
    "stl1": ("Suhu tanah lapisan 1 (0-7 cm)", "°C", _k2c, "K - 273,15", "biasa"),
    "stl2": ("Suhu tanah lapisan 2 (7-28 cm)", "°C", _k2c, "K - 273,15", "biasa"),
    "sp":   ("Tekanan permukaan", "hPa", _kali(0.01), "Pa / 100", "biasa"),
    "msl":  ("Tekanan permukaan laut", "hPa", _kali(0.01), "Pa / 100", "biasa"),
    "tp":   ("Curah hujan total", "mm", _kali(1000), "m x 1000", "jumlah"),
    "cp":   ("Hujan konvektif", "mm", _kali(1000), "m x 1000", "jumlah"),
    "lsp":  ("Hujan skala besar", "mm", _kali(1000), "m x 1000", "jumlah"),
    "sf":   ("Salju (setara air)", "mm", _kali(1000), "m x 1000", "jumlah"),
    "ro":   ("Limpasan total", "mm", _kali(1000), "m x 1000", "jumlah"),
    "e":    ("Evaporasi (positif = menguap)", "mm", _kali(-1000), "m x -1000", "jumlah"),
    "pev":  ("Evaporasi potensial (positif = menguap)", "mm", _kali(-1000), "m x -1000", "jumlah"),
    "ssrd": ("Radiasi matahari ke permukaan", "W/m²", _kali(1 / 3600), "J/m² per jam / 3600", "fluks"),
    "strd": ("Radiasi gelombang panjang ke permukaan", "W/m²", _kali(1 / 3600), "J/m² per jam / 3600", "fluks"),
    "ssr":  ("Radiasi matahari bersih", "W/m²", _kali(1 / 3600), "J/m² per jam / 3600", "fluks"),
    "str":  ("Radiasi gelombang panjang bersih", "W/m²", _kali(1 / 3600), "J/m² per jam / 3600", "fluks"),
    "tisr": ("Radiasi matahari puncak atmosfer", "W/m²", _kali(1 / 3600), "J/m² per jam / 3600", "fluks"),
    "tcc":  ("Tutupan awan total", "%", _kali(100), "fraksi x 100", "biasa"),
    "lcc":  ("Tutupan awan rendah", "%", _kali(100), "fraksi x 100", "biasa"),
    "mcc":  ("Tutupan awan menengah", "%", _kali(100), "fraksi x 100", "biasa"),
    "hcc":  ("Tutupan awan tinggi", "%", _kali(100), "fraksi x 100", "biasa"),
    "r":    ("Kelembapan relatif", "%", None, "tanpa konversi", "biasa"),
    "q":    ("Kelembapan spesifik", "g/kg", _kali(1000), "kg/kg x 1000", "biasa"),
    "z":    ("Tinggi geopotensial", "m", _kali(1 / 9.80665), "m²/s² / 9,80665", "biasa"),
    "w":    ("Kecepatan vertikal", "Pa/s", None, "tanpa konversi", "biasa"),
    "swh":  ("Tinggi gelombang signifikan", "m", None, "tanpa konversi", "biasa"),
    "shww": ("Tinggi gelombang angin", "m", None, "tanpa konversi", "biasa"),
    "shts": ("Tinggi swell total", "m", None, "tanpa konversi", "biasa"),
    "mwp":  ("Periode gelombang rata-rata", "s", None, "tanpa konversi", "biasa"),
    "pp1d": ("Periode puncak gelombang", "s", None, "tanpa konversi", "biasa"),
    "mwd":  ("Arah gelombang rata-rata", "°", None, "derajat dari utara, arah datang", "arah"),
    "cape": ("CAPE", "J/kg", None, "tanpa konversi", "biasa"),
    "blh":  ("Tinggi lapisan batas atmosfer", "m", None, "tanpa konversi", "biasa"),
    "i10fg": ("Hembusan angin maksimum 10 m", "m/s", None, "tanpa konversi", "biasa"),
    "fg10": ("Hembusan angin maksimum 10 m", "m/s", None, "tanpa konversi", "biasa"),
    "tcwv": ("Total uap air kolom", "mm", None, "kg/m² setara mm", "biasa"),
    "swvl1": ("Kelembapan tanah lapisan 1 (0-7 cm)", "m³/m³", None, "tanpa konversi", "biasa"),
    "swvl2": ("Kelembapan tanah lapisan 2 (7-28 cm)", "m³/m³", None, "tanpa konversi", "biasa"),
}

# Variabel akumulasi yang nilainya tidak boleh negatif (sisa pembulatan GRIB)
TIDAK_NEGATIF = {"tp", "cp", "lsp", "sf", "ro", "ssrd", "strd", "tisr"}


# --------------------------------------------------------------------------
# Struktur hasil
# --------------------------------------------------------------------------
@dataclass
class InfoGrid:
    sumber: str
    variabel: str
    lat_grid: float
    lon_grid: float
    jarak_km: float
    resolusi: str
    catatan: str = ""


@dataclass
class Hasil:
    data: pd.DataFrame                      # hasil terjemahan, index waktu lokal
    mentah: pd.DataFrame                    # nilai asli, index waktu UTC
    meta: dict                              # kolom mentah -> keterangan
    satuan: dict                            # kolom hasil -> satuan
    jenis: dict                             # kolom hasil -> jenis
    konversi: list                          # (kolom, asal, keterangan)
    grid: list                              # list[InfoGrid]
    peringatan: list = field(default_factory=list)
    zona: str = "UTC"
    lat_input: float = 0.0
    lon_input: float = 0.0
    sumber: list = field(default_factory=list)


# --------------------------------------------------------------------------
# Membaca GRIB dan memilih grid
# --------------------------------------------------------------------------
def buka_grib(path: str):
    """Buka semua 'hypercube' di file GRIB. ERA5 sering mencampur variabel
    instan dan akumulasi, jadi open_datasets lebih aman dari open_dataset."""
    import cfgrib
    return cfgrib.open_datasets(path, backend_kwargs={"indexpath": ""})


def jarak_km(lat1, lon1, lat2, lon2):
    p1, p2 = np.radians(lat1), np.radians(lat2)
    dphi = p2 - p1
    dlam = np.radians(lon2 - lon1)
    a = np.sin(dphi / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(dlam / 2) ** 2
    return 2 * R_BUMI_KM * np.arcsin(np.sqrt(np.clip(a, 0, 1)))


def _mask_berisi(ds):
    """Peta True/False grid yang punya nilai. Dipakai supaya titik daratan
    tidak terpilih untuk data laut (gelombang, SST) dan sebaliknya."""
    var = next(iter(ds.data_vars.values()))
    lain = [d for d in var.dims if d not in DIM_RUANG]
    if not lain:
        arr = var.transpose("latitude", "longitude").values
        return np.isfinite(arr)
    bentuk = [var.sizes[d] for d in lain]
    for k in range(min(int(np.prod(bentuk)), 60)):
        idx = dict(zip(lain, np.unravel_index(k, bentuk)))
        arr = var.isel(idx).transpose("latitude", "longitude").values
        if np.isfinite(arr).any():
            return np.isfinite(arr)
    return None


def cari_grid_terdekat(ds, lat, lon, lewati_kosong=True):
    """Kembalikan (j, i, lat_grid, lon_grid, jarak_km, catatan, peringatan)."""
    if "latitude" not in ds.dims or "longitude" not in ds.dims:
        raise ValueError("Dataset tidak memakai grid lintang-bujur reguler.")
    lats = ds["latitude"].values
    lons = ds["longitude"].values
    LA, LO = np.meshgrid(lats, lons, indexing="ij")
    d = jarak_km(lat, lon, LA, LO)

    j0, i0 = np.unravel_index(np.argmin(d), d.shape)
    catatan, peringatan = "", ""

    dlat = float(np.abs(np.diff(lats)).min()) if lats.size > 1 else 0.25
    dlon = float(np.abs(np.diff(lons)).min()) if lons.size > 1 else 0.25
    batas = 1.5 * max(dlat, dlon) * 111.2
    if d[j0, i0] > batas:
        peringatan = (f"Koordinat input berada di luar cakupan data "
                      f"({lats.min():g} s.d. {lats.max():g} LU/LS, "
                      f"{lons.min():g} s.d. {lons.max():g} BT/BB). "
                      f"Grid terdekat berjarak {d[j0, i0]:.1f} km.")

    j, i = j0, i0
    if lewati_kosong:
        mask = _mask_berisi(ds)
        if mask is not None and mask.any() and not mask[j0, i0]:
            d2 = np.where(mask, d, np.inf)
            j, i = np.unravel_index(np.argmin(d2), d2.shape)
            catatan = (f"Grid terdekat ({lats[j0]:g}, {lons[i0]:g}) tidak berisi data, "
                       f"dipakai grid berisi terdekat.")

    lon_g = float(lons[i])
    if lon_g > 180:
        lon_g -= 360
    return int(j), int(i), float(lats[j]), lon_g, float(d[j, i]), catatan, peringatan, f"{dlat:g}° x {dlon:g}°"


def _label_level(dim, nilai):
    if dim == "isobaricInhPa":
        return f"{float(nilai):g} hPa"
    if dim == "depthBelowLandLayer":
        return f"kedalaman {float(nilai):g} m"
    if dim == "number":
        return f"anggota {int(nilai)}"
    return f"{dim} {nilai}"


def ekstrak_titik(ds, j, i):
    """Ambil deret waktu di satu titik grid. Hasil: dict kunci -> Series
    (index waktu UTC) dan dict meta per kunci."""
    titik = ds.isel(latitude=j, longitude=i).load()
    seri, meta = {}, {}
    for nama, da in titik.data_vars.items():
        # Koordinat waktu skalar (file berisi satu waktu) dijadikan dimensi
        for c in ("time", "step"):
            if c in da.coords and c not in da.dims and da[c].ndim == 0:
                da = da.expand_dims(c)

        label_tetap = []
        for c in ("isobaricInhPa", "depthBelowLandLayer"):
            if c in da.coords and c not in da.dims and da[c].ndim == 0:
                label_tetap.append(_label_level(c, da[c].values))
        dim_level = [d for d in da.dims if d not in DIM_WAKTU]

        df = da.to_dataframe(name="__nilai").reset_index()
        if "valid_time" in df.columns:
            waktu = pd.to_datetime(df["valid_time"])
        elif "time" in df.columns:
            waktu = pd.to_datetime(df["time"])
            if "step" in df.columns:
                waktu = waktu + pd.to_timedelta(df["step"])
        else:
            continue

        kerja = pd.DataFrame({"waktu": waktu.values, "nilai": df["__nilai"].values})
        if dim_level:
            kerja["label"] = df[dim_level].apply(
                lambda r: ", ".join(_label_level(dd, r[dd]) for dd in dim_level), axis=1)
        else:
            kerja["label"] = ""
        if label_tetap:
            awal = ", ".join(label_tetap)
            kerja["label"] = kerja["label"].map(lambda s: f"{awal}, {s}" if s else awal)
        kerja = kerja.dropna(subset=["waktu", "nilai"])

        langkah_maks = None
        if "step" in da.coords:
            langkah = pd.to_timedelta(np.atleast_1d(da["step"].values))
            if len(langkah):
                langkah_maks = float(langkah.max() / pd.Timedelta(hours=1))

        for label, g in kerja.groupby("label", sort=False):
            kunci = f"{nama}@{label}" if label else nama
            s = g.groupby("waktu")["nilai"].first().sort_index()
            s.index.name = "waktu"
            seri[kunci] = s
            meta[kunci] = {
                "nama": nama,
                "level": label,
                "satuan": da.attrs.get("units", da.attrs.get("GRIB_units", "")),
                "nama_panjang": da.attrs.get("long_name", da.attrs.get("GRIB_name", nama)),
                "jenis_langkah": da.attrs.get("GRIB_stepType", ""),
                "langkah_maks_jam": langkah_maks,
            }
    return seri, meta


def proses_file(path, nama_tampil, lat, lon, lewati_kosong=True):
    seri, meta, grid, peringatan = {}, {}, [], []
    daftar = buka_grib(path)
    if not daftar:
        raise ValueError(f"{nama_tampil}: tidak ada data yang bisa dibaca.")
    for ds in daftar:
        j, i, lat_g, lon_g, jarak, catatan, warn, res = cari_grid_terdekat(ds, lat, lon, lewati_kosong)
        if warn:
            peringatan.append(f"{nama_tampil}: {warn}")
        s, m = ekstrak_titik(ds, j, i)
        seri.update(s)
        meta.update(m)
        grid.append(InfoGrid(nama_tampil, ", ".join(ds.data_vars), lat_g, lon_g, jarak, res, catatan))
    return seri, meta, grid, peringatan


# --------------------------------------------------------------------------
# Terjemahan variabel
# --------------------------------------------------------------------------
def ke_mata_angin(derajat: pd.Series) -> pd.Series:
    hasil = pd.Series(pd.NA, index=derajat.index, dtype="object")
    ada = derajat.notna()
    idx = (np.floor(((derajat[ada] % 360) + 22.5) / 45).astype(int)) % 8
    hasil[ada] = [NAMA_ARAH[k] for k in idx]
    return hasil


def kumulatif_ke_per_jam(s: pd.Series) -> pd.Series:
    """ERA5-Land: akumulasi dimulai dari 00 UTC. Nilai per jam = selisih
    dengan jam sebelumnya, kecuali pukul 01 UTC yang sudah nilai 1 jam."""
    s = s.dropna().sort_index()
    if s.empty:
        return s
    per_jam = s.diff()
    jam1 = s.index.hour == 1
    per_jam[jam1] = s[jam1]
    loncat = s.index.to_series().diff() != pd.Timedelta(hours=1)
    per_jam[loncat.values & ~jam1] = np.nan
    return per_jam


def terjemahkan(mentah: pd.DataFrame, meta: dict, mode_akumulasi="otomatis"):
    keluar, satuan, jenis, konversi = {}, {}, {}, []

    grup = {}
    for kunci in mentah.columns:
        m = meta[kunci]
        grup.setdefault(m["level"], {})[m["nama"]] = kunci

    def tambah(kolom, nilai, sat, jns, asal, ket):
        keluar[kolom] = nilai
        satuan[kolom] = sat
        jenis[kolom] = jns
        konversi.append((kolom, asal, ket))

    for level, var in grup.items():
        akhiran = f" {level}" if level else ""
        terpakai = set()

        for u, v, tinggi in PASANGAN_ANGIN:
            if u in var and v in var:
                U, V = mentah[var[u]], mentah[var[v]]
                dasar = "angin" + (f" {tinggi}" if tinggi else "") + akhiran
                kec = np.sqrt(U ** 2 + V ** 2)
                arah = (270 - np.degrees(np.arctan2(V, U))) % 360
                asal = f"{var[u]}, {var[v]}"
                tambah(f"Kecepatan {dasar} (m/s)", kec, "m/s", "biasa", asal, "akar(u² + v²)")
                tambah(f"Arah {dasar} (°)", arah, "°", "arah", asal,
                       "derajat dari utara searah jarum jam, arah datangnya angin")
                tambah(f"Arah {dasar} (mata angin)", ke_mata_angin(arah), "", "teks", asal,
                       "8 arah mata angin dari kolom derajat")
                terpakai |= {u, v}

        for nama, kunci in var.items():
            if nama in terpakai:
                continue
            s = mentah[kunci]
            m = meta[kunci]
            if nama in KONVERSI:
                tampil, sat, fungsi, ket, jns = KONVERSI[nama]
                if jns in ("jumlah", "fluks"):
                    mode = mode_akumulasi
                    if mode == "otomatis":
                        mode = "sejak_00" if (m["langkah_maks_jam"] or 0) > 12 else "per_jam"
                    if mode == "sejak_00":
                        s = kumulatif_ke_per_jam(s).reindex(mentah.index)
                        ket += "; diubah dari akumulasi sejak 00 UTC ke nilai per jam"
                    else:
                        ket += "; nilai sudah akumulasi 1 jam"
                nilai = fungsi(s) if fungsi else s
                if nama in TIDAK_NEGATIF:
                    nilai = nilai.clip(lower=0)
                kolom = f"{tampil}{akhiran} ({sat})"
                tambah(kolom, nilai, sat, jns, kunci, ket)
                if jns == "arah":
                    tambah(f"{tampil}{akhiran} (mata angin)", ke_mata_angin(nilai), "", "teks",
                           kunci, "8 arah mata angin dari kolom derajat")
            else:
                sat = m["satuan"]
                kolom = f"{m['nama_panjang']}{akhiran}" + (f" ({sat})" if sat else "")
                tambah(kolom, s, sat, "biasa", kunci, "tanpa konversi, nama dan satuan dari file")

        if "t2m" in var and "d2m" in var:
            T = mentah[var["t2m"]] - 273.15
            Td = mentah[var["d2m"]] - 273.15
            rh = 100 * np.exp(17.625 * Td / (243.04 + Td)) / np.exp(17.625 * T / (243.04 + T))
            tambah(f"Kelembapan relatif 2 m{akhiran} (%)", rh.clip(upper=100), "%", "biasa",
                   f"{var['t2m']}, {var['d2m']}", "rumus Magnus dari suhu udara dan titik embun")

    data = pd.DataFrame(keluar, index=mentah.index)
    return data, satuan, jenis, konversi


def proses(daftar_file, lat, lon, zona="UTC", mode_akumulasi="otomatis", lewati_kosong=True):
    """daftar_file: list (path, nama_tampil)."""
    semua_seri, meta, grid, peringatan = {}, {}, [], []
    for path, nama in daftar_file:
        s, m, g, p = proses_file(path, nama, lat, lon, lewati_kosong)
        for k, v in s.items():
            semua_seri[k] = v if k not in semua_seri else semua_seri[k].combine_first(v)
        meta.update(m)
        grid.extend(g)
        peringatan.extend(p)
    if not semua_seri:
        raise ValueError("Tidak ada variabel yang berhasil diambil dari file.")

    mentah = pd.concat(semua_seri, axis=1).sort_index()
    mentah.index.name = "waktu"
    data, satuan, jenis, konversi = terjemahkan(mentah, meta, mode_akumulasi)

    geser = pd.Timedelta(hours=ZONA_WAKTU[zona])
    data.index = data.index + geser
    data = data.dropna(how="all")

    return Hasil(data=data, mentah=mentah, meta=meta, satuan=satuan, jenis=jenis,
                 konversi=konversi, grid=grid, peringatan=peringatan, zona=zona,
                 lat_input=lat, lon_input=lon, sumber=[n for _, n in daftar_file])


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------
def buat_excel(h: Hasil, sertakan_mentah=False) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
    from openpyxl.utils import get_column_letter

    FONT = "Century Gothic"
    garis = Side(style="thin", color="FFD9D9D9")
    KOTAK = Border(left=garis, right=garis, top=garis, bottom=garis)
    KUNING = PatternFill("solid", fgColor="FFFFFF00")

    wb = Workbook()
    gaya = {
        "th": NamedStyle("th", font=Font(name=FONT, size=10, bold=True),
                         alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)),
        "waktu": NamedStyle("waktu", font=Font(name=FONT, size=10), number_format="yyyy-mm-dd hh:mm",
                            alignment=Alignment(horizontal="left", vertical="center")),
        "angka": NamedStyle("angka", font=Font(name=FONT, size=10), number_format="0.000",
                            alignment=Alignment(horizontal="right", vertical="center")),
        "teks": NamedStyle("teks", font=Font(name=FONT, size=10),
                           alignment=Alignment(horizontal="left", vertical="center", indent=1)),
        "angka_mentah": NamedStyle("angka_mentah", font=Font(name=FONT, size=10), number_format="0.0######",
                                   alignment=Alignment(horizontal="right", vertical="center")),
        "th_k": NamedStyle("th_k", font=Font(name=FONT, size=10, bold=True), border=KOTAK,
                           alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)),
        "teks_k": NamedStyle("teks_k", font=Font(name=FONT, size=10), border=KOTAK,
                             alignment=Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)),
        "angka_k": NamedStyle("angka_k", font=Font(name=FONT, size=10), border=KOTAK, number_format="0.00",
                              alignment=Alignment(horizontal="right", vertical="center")),
        "persen_k": NamedStyle("persen_k", font=Font(name=FONT, size=10), border=KOTAK, number_format="0.0%",
                               alignment=Alignment(horizontal="right", vertical="center")),
        "strip_k": NamedStyle("strip_k", font=Font(name=FONT, size=10), border=KOTAK,
                              alignment=Alignment(horizontal="center", vertical="center")),
        "label": NamedStyle("label", font=Font(name=FONT, size=10, bold=True),
                            alignment=Alignment(horizontal="left", vertical="center")),
    }
    for g in gaya.values():
        wb.add_named_style(g)

    def lebar(teks, minimum=12, maksimum=30):
        return max(minimum, min(maksimum, round(len(str(teks)) * 1.15 / 1.6 + 3)))

    def tulis_tabel_waktu(ws, df, judul_waktu, kolom_teks=(), gaya_angka="angka"):
        ws.append([judul_waktu] + list(df.columns))
        for c in range(1, df.shape[1] + 2):
            ws.cell(1, c).style = "th"
        ws.row_dimensions[1].height = 36
        ws.column_dimensions["A"].width = 20
        for k, nama in enumerate(df.columns, start=2):
            ws.column_dimensions[get_column_letter(k)].width = lebar(nama, 14, 24)
        gaya_kolom = ["waktu"] + ["teks" if c in kolom_teks else gaya_angka for c in df.columns]
        nilai = df.astype(object).where(df.notna(), None)
        for r, (t, baris) in enumerate(zip(df.index, nilai.itertuples(index=False)), start=2):
            ws.append([t.to_pydatetime()] + [None if v is pd.NA else v for v in baris])
            for c, gs in enumerate(gaya_kolom, start=1):
                ws.cell(r, c).style = gs
        ws.sheet_format.defaultRowHeight = 18
        ws.sheet_format.customHeight = True
        ws.auto_filter.ref = f"A1:{get_column_letter(df.shape[1] + 1)}{df.shape[0] + 1}"

    # ---- Sheet Data
    ws = wb.active
    ws.title = "Data"
    kolom_teks = [c for c, j in h.jenis.items() if j == "teks"]
    tulis_tabel_waktu(ws, h.data, f"Waktu ({h.zona.split(' ')[0]})", kolom_teks)
    n_akhir = h.data.shape[0] + 1
    huruf = {nama: get_column_letter(k) for k, nama in enumerate(h.data.columns, start=2)}

    # ---- Sheet Ringkasan (formula, ikut berubah kalau Data diedit)
    wr = wb.create_sheet("Ringkasan")
    judul = ["Variabel", "Satuan", "Jumlah data", "Rata-rata", "Minimum", "Maksimum", "Total"]
    wr.append(judul)
    for c in range(1, len(judul) + 1):
        wr.cell(1, c).style = "th_k"
    wr.row_dimensions[1].height = 22
    r = 2
    for nama in h.data.columns:
        if h.jenis[nama] in ("arah", "teks"):
            continue
        rg = f"Data!{huruf[nama]}2:{huruf[nama]}{n_akhir}"
        wr.cell(r, 1, nama).style = "teks_k"
        wr.cell(r, 2, h.satuan[nama]).style = "teks_k"
        wr.cell(r, 3, f"=COUNT({rg})").style = "angka_k"
        wr.cell(r, 3).number_format = "0"
        wr.cell(r, 4, f'=IF(COUNT({rg})=0,"-",AVERAGE({rg}))').style = "angka_k"
        wr.cell(r, 5, f'=IF(COUNT({rg})=0,"-",MIN({rg}))').style = "angka_k"
        wr.cell(r, 6, f'=IF(COUNT({rg})=0,"-",MAX({rg}))').style = "angka_k"
        if h.jenis[nama] == "jumlah":
            wr.cell(r, 7, f"=SUM({rg})").style = "angka_k"
        else:
            wr.cell(r, 7, "-").style = "strip_k"
        wr.row_dimensions[r].height = 18
        r += 1

    for nama in [c for c in h.data.columns if h.jenis[c] == "teks"]:
        r += 1
        wr.cell(r, 1, "Frekuensi " + nama.replace(" (mata angin)", "")[:1].lower() + nama.replace(" (mata angin)", "")[1:]).style = "label"
        wr.row_dimensions[r].height = 22
        r += 1
        for c, t in enumerate(["Arah", "Jumlah data", "Persentase"], start=1):
            wr.cell(r, c, t).style = "th_k"
        wr.row_dimensions[r].height = 22
        awal = r + 1
        rg = f"Data!{huruf[nama]}2:{huruf[nama]}{n_akhir}"
        for k, arah in enumerate(NAMA_ARAH):
            rr = awal + k
            wr.cell(rr, 1, arah).style = "teks_k"
            wr.cell(rr, 2, f'=COUNTIF({rg},"{arah}")').style = "angka_k"
            wr.cell(rr, 2).number_format = "0"
            wr.cell(rr, 3, f"=IF(SUM($B${awal}:$B${awal + 7})=0,0,B{rr}/SUM($B${awal}:$B${awal + 7}))").style = "persen_k"
            wr.row_dimensions[rr].height = 18
        r = awal + 8
    for kol, w in zip("ABCDEFG", [46, 10, 13, 13, 13, 13, 13]):
        wr.column_dimensions[kol].width = w

    # ---- Sheet Info
    wi = wb.create_sheet("Info")
    baris_info = [
        ("Dibuat", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("File sumber", ", ".join(h.sumber)),
        ("Koordinat input", f"{h.lat_input:.4f}, {h.lon_input:.4f}"),
        ("Zona waktu", h.zona),
        ("Periode", f"{h.data.index.min():%Y-%m-%d %H:%M} s.d. {h.data.index.max():%Y-%m-%d %H:%M}"
         if len(h.data) else "-"),
        ("Jumlah baris", len(h.data)),
    ]
    for k, (a, b) in enumerate(baris_info, start=1):
        wi.cell(k, 1, a).style = "label"
        wi.cell(k, 2, b).style = "teks"
        wi.row_dimensions[k].height = 18
    r = len(baris_info) + 2
    for p in h.peringatan:
        sel = wi.cell(r, 1, f"Perhatian: {p}")
        sel.style = "teks"
        sel.fill = KUNING
        r += 1
    r += 0 if not h.peringatan else 1

    wi.cell(r, 1, "Grid terpilih").style = "label"
    r += 1
    judul_grid = ["File", "Variabel", "Catatan", "Lintang grid", "Bujur grid", "Jarak (km)", "Resolusi"]
    for c, t in enumerate(judul_grid, start=1):
        wi.cell(r, c, t).style = "th_k"
    wi.row_dimensions[r].height = 22
    for g in h.grid:
        r += 1
        isi = [g.sumber, g.variabel, g.catatan or "-", g.lat_grid, g.lon_grid, g.jarak_km, g.resolusi]
        for c, v in enumerate(isi, start=1):
            wi.cell(r, c, v).style = "angka_k" if isinstance(v, float) else "teks_k"
        wi.cell(r, 4).number_format = wi.cell(r, 5).number_format = "0.00"
        if g.catatan:
            wi.cell(r, 3).fill = KUNING
        wi.row_dimensions[r].height = 30 if g.catatan else 18

    r += 2
    wi.cell(r, 1, "Asal dan konversi kolom").style = "label"
    r += 1
    for c, t in enumerate(["Kolom di sheet Data", "Variabel asal", "Konversi"], start=1):
        wi.cell(r, c, t).style = "th_k"
    wi.row_dimensions[r].height = 22
    for kol, asal, ket in h.konversi:
        r += 1
        for c, v in enumerate([kol, asal, ket], start=1):
            wi.cell(r, c, v).style = "teks_k"
        wi.row_dimensions[r].height = 30 if (len(kol) > 30 or len(ket) > 42) else 18

    r += 2
    catatan = [
        "Arah angin dan gelombang adalah arah datang, diukur dari utara searah jarum jam.",
        "Nilai akumulasi (hujan, radiasi) dicatat pada akhir jam akumulasinya.",
        "Kolom angka di sheet Data disimpan tanpa pembulatan; tampilan 3 desimal hanya format sel.",
    ]
    for t in catatan:
        wi.cell(r, 1, t).font = Font(name=FONT, size=8, color="FF7F7F7F")
        wi.row_dimensions[r].height = 18
        r += 1
    for kol, w in zip("ABCDEFG", [36, 28, 48, 13, 13, 12, 15]):
        wi.column_dimensions[kol].width = w

    for lembar in (wr, wi):
        lembar.page_setup.orientation = "landscape"
        lembar.page_setup.fitToWidth = 1
        lembar.page_setup.fitToHeight = 0
        lembar.sheet_properties.pageSetUpPr.fitToPage = True

    # ---- Sheet Data mentah (opsional)
    if sertakan_mentah:
        wm = wb.create_sheet("Data mentah")
        mentah = h.mentah.copy()
        mentah.columns = [
            f"{k} ({h.meta[k]['satuan']})" if h.meta[k]["satuan"] else k for k in mentah.columns
        ]
        tulis_tabel_waktu(wm, mentah, "Waktu (UTC)", gaya_angka="angka_mentah")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------
# Antarmuka Streamlit
# --------------------------------------------------------------------------
def main():
    import streamlit as st

    st.set_page_config(page_title="GRIB ke Excel", layout="wide")
    st.title("GRIB ke Excel")
    st.caption("Ambil deret waktu ERA5 dari file GRIB pada titik grid terdekat, lalu simpan sebagai Excel.")

    with st.sidebar:
        st.header("1. File GRIB")
        cara = st.radio("Sumber file", ["Unggah", "Path di komputer"], horizontal=True,
                        help="File besar lebih cepat dibaca langsung dari path tanpa diunggah.")
        unggahan, path_teks = [], ""
        if cara == "Unggah":
            unggahan = st.file_uploader("Pilih file .grib", accept_multiple_files=True,
                                        type=["grib", "grb", "grib1", "grib2", "grb2"])
        else:
            path_teks = st.text_area("Path file, satu per baris",
                                     placeholder="D:/data/era5_angin_2024.grib")

        st.header("2. Lokasi")
        c1, c2 = st.columns(2)
        lat = c1.number_input("Lintang", min_value=-90.0, max_value=90.0, value=None,
                              format="%.4f", placeholder="-7.2500", help="Negatif untuk lintang selatan")
        lon = c2.number_input("Bujur", min_value=-180.0, max_value=360.0, value=None,
                              format="%.4f", placeholder="112.7500", help="Negatif untuk bujur barat")

        st.header("3. Pengaturan")
        zona = st.selectbox("Zona waktu di Excel", list(ZONA_WAKTU), index=1)
        mode_label = {
            "Otomatis": "otomatis",
            "Per jam (ERA5)": "per_jam",
            "Akumulasi sejak 00 UTC (ERA5-Land)": "sejak_00",
        }
        mode = st.selectbox("Variabel akumulasi (hujan, radiasi)", list(mode_label),
                            help="ERA5 menyimpan hujan per jam. ERA5-Land menyimpan total sejak 00 UTC, "
                                 "jadi perlu diselisihkan. Mode otomatis menebak dari panjang step di file.")
        lewati = st.checkbox("Lewati grid kosong", value=True,
                             help="Contoh: data gelombang tidak ada di daratan. Kalau grid terdekat kosong, "
                                  "dipakai grid berisi yang paling dekat.")
        mentah = st.checkbox("Sertakan sheet data mentah", value=False)
        jalan = st.button("Proses", type="primary", width="stretch")

    if jalan:
        if lat is None or lon is None:
            st.warning("Isi lintang dan bujur dulu.")
            st.stop()
        sementara, daftar = [], []
        try:
            if cara == "Unggah":
                if not unggahan:
                    st.warning("Belum ada file yang diunggah.")
                    st.stop()
                for f in unggahan:
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".grib")
                    tmp.write(f.getbuffer())
                    tmp.close()
                    sementara.append(tmp.name)
                    daftar.append((tmp.name, f.name))
            else:
                for p in [x.strip().strip('"') for x in path_teks.splitlines() if x.strip()]:
                    if not os.path.isfile(p):
                        st.error(f"File tidak ditemukan: {p}")
                        st.stop()
                    daftar.append((p, os.path.basename(p)))
                if not daftar:
                    st.warning("Isi path file dulu.")
                    st.stop()

            with st.spinner("Membaca GRIB dan menerjemahkan variabel..."):
                hasil = proses(daftar, float(lat), float(lon), zona, mode_label[mode], lewati)
                xlsx = buat_excel(hasil, mentah)
            st.session_state["hasil"] = hasil
            st.session_state["xlsx"] = xlsx
        except Exception as e:  # noqa: BLE001
            pesan = str(e)
            st.error(f"Gagal memproses file: {pesan}")
            if "eccodes" in pesan.lower() or "ecCodes" in pesan:
                st.info("Library ecCodes belum terpasang. Coba: conda install -c conda-forge cfgrib")
            st.stop()
        finally:
            for p in sementara:
                try:
                    os.remove(p)
                except OSError:
                    pass

    hasil: Hasil | None = st.session_state.get("hasil")
    if hasil is None:
        st.info("Pilih file GRIB, isi koordinat, lalu tekan Proses.")
        return

    for p in hasil.peringatan:
        st.warning(p)

    g0 = hasil.grid[0]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Grid terpilih", f"{g0.lat_grid:.2f}, {g0.lon_grid:.2f}")
    m2.metric("Jarak dari input", f"{g0.jarak_km:.1f} km")
    m3.metric("Jumlah baris", f"{len(hasil.data):,}".replace(",", "."))
    m4.metric("Kolom hasil", hasil.data.shape[1])

    st.subheader("Grid per dataset")
    st.dataframe(pd.DataFrame([g.__dict__ for g in hasil.grid]).rename(columns={
        "sumber": "File", "variabel": "Variabel", "lat_grid": "Lintang grid", "lon_grid": "Bujur grid",
        "jarak_km": "Jarak (km)", "resolusi": "Resolusi", "catatan": "Catatan"}),
        hide_index=True, width="stretch")

    peta = pd.DataFrame(
        [{"lat": hasil.lat_input, "lon": hasil.lon_input if hasil.lon_input <= 180 else hasil.lon_input - 360,
          "warna": "#d62728"}]
        + [{"lat": g.lat_grid, "lon": g.lon_grid, "warna": "#1f77b4"} for g in hasil.grid])
    st.map(peta, latitude="lat", longitude="lon", color="warna", size=600, zoom=8)
    st.caption("Merah: koordinat input. Biru: grid yang dipakai.")

    st.subheader("Pratinjau data")
    st.dataframe(hasil.data.head(500), width="stretch")

    angka = [c for c, j in hasil.jenis.items() if j not in ("teks", "arah")]
    if angka:
        pilih = st.multiselect("Tampilkan grafik", angka, default=angka[:1])
        if pilih:
            st.line_chart(hasil.data[pilih])

    nama_file = f"era5_{hasil.lat_input:.3f}_{hasil.lon_input:.3f}.xlsx".replace("-", "m")
    st.download_button("Unduh Excel", st.session_state["xlsx"], file_name=nama_file,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")


if __name__ == "__main__":
    main()
