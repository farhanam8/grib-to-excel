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
    2. Aplikasi mencari titik grid terdekat untuk setiap bentuk grid di file GRIB,
       lalu membaca nilai di titik itu saja dari setiap message (via ecCodes).
    3. Variabel mentah diterjemahkan ke besaran yang mudah dibaca,
       misalnya u10 + v10 menjadi kecepatan dan arah angin.
    4. Grafik deret waktu (resolusi otomatis) dan windrose untuk data angin.
    5. Hasilnya diunduh sebagai Excel (sheet Data, Ringkasan, Windrose, Info).
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
    angin: list = field(default_factory=list)   # [{"nama", "kec", "arah"}] untuk windrose


# --------------------------------------------------------------------------
# Membaca GRIB dan memilih grid
# --------------------------------------------------------------------------
# GRIB disimpan sebagai deretan "message", satu message = satu variabel pada
# satu waktu dan satu level. Aplikasi ini membaca message satu per satu lewat
# ecCodes dan hanya mengambil SATU nilai (titik grid terpilih) dari tiap message,
# tanpa membongkar seluruh peta. Ini jauh lebih cepat dan hemat memori
# dibanding membuka seluruh file dengan xarray/cfgrib.

def jarak_km(lat1, lon1, lat2, lon2):
    p1, p2 = np.radians(lat1), np.radians(lat2)
    dphi = p2 - p1
    dlam = np.radians(lon2 - lon1)
    a = np.sin(dphi / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(dlam / 2) ** 2
    return 2 * R_BUMI_KM * np.arcsin(np.sqrt(np.clip(a, 0, 1)))


def _ambil(ec, h, kunci, bawaan=None):
    try:
        return ec.codes_get(h, kunci) if ec.codes_is_defined(h, kunci) else bawaan
    except Exception:  # noqa: BLE001
        return bawaan


def _geometri(ec, h, lat, lon):
    """Hitung jarak input ke semua titik grid. Dipanggil sekali per bentuk grid."""
    lats = ec.codes_get_array(h, "latitudes")
    lons = ec.codes_get_array(h, "longitudes")
    d = jarak_km(lat, lon, lats, lons)
    idx0 = int(np.argmin(d))

    if _ambil(ec, h, "gridType") == "regular_ll":
        di = float(_ambil(ec, h, "iDirectionIncrementInDegrees", 0.25))
        dj = float(_ambil(ec, h, "jDirectionIncrementInDegrees", 0.25))
        resolusi, langkah = f"{di:g}° x {dj:g}°", max(di, dj)
    else:
        resolusi, langkah = str(_ambil(ec, h, "gridType", "-")), 0.5

    peringatan = ""
    if d[idx0] > 1.5 * langkah * 111.2:
        peringatan = (f"Koordinat input berada di luar cakupan data "
                      f"(lintang {lats.min():g} s.d. {lats.max():g}, bujur {lons.min():g} s.d. {lons.max():g}). "
                      f"Grid terdekat berjarak {d[idx0]:.1f} km.")
    return {"lats": lats, "lons": lons, "d": d, "idx0": idx0,
            "resolusi": resolusi, "peringatan": peringatan}


def _label_level(ec, h):
    jenis = _ambil(ec, h, "typeOfLevel", "")
    level = _ambil(ec, h, "level", 0)
    label = ""
    if jenis == "isobaricInhPa":
        label = f"{level:g} hPa"
    elif jenis in ("hybrid", "modelLevel"):
        label = f"model level {level}"
    elif jenis == "potentialVorticity":
        label = f"PV {level}"
    elif jenis == "theta":
        label = f"theta {level} K"
    anggota = _ambil(ec, h, "numberOfForecastsInEnsemble", 0) or 0
    if anggota > 1:
        no = _ambil(ec, h, "number", 0)
        label = f"{label}, anggota {no}" if label else f"anggota {no}"
    return label


def _pesan_grib(path):
    """Potong file menjadi message GRIB mentah (bytes) tanpa ecCodes.
    Menghasilkan (posisi_byte, bytes). Message GRIB1 berukuran sangat besar
    (format 'large GRIB1') dikembalikan None supaya dibaca lewat ecCodes."""
    import mmap
    with open(path, "rb") as f:
        if os.path.getsize(path) == 0:
            return
        mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
        try:
            pos, n = 0, len(mm)
            while True:
                pos = mm.find(b"GRIB", pos)
                if pos < 0 or pos + 16 > n:
                    break
                edisi = mm[pos + 7]
                if edisi == 1:
                    panjang = int.from_bytes(mm[pos + 4:pos + 7], "big")
                    if panjang & 0x800000:          # large GRIB1, serahkan ke ecCodes
                        yield pos, None
                        return
                elif edisi == 2:
                    panjang = int.from_bytes(mm[pos + 8:pos + 16], "big")
                else:
                    pos += 4
                    continue
                if panjang < 16 or pos + panjang > n:
                    break
                yield pos, mm[pos:pos + panjang]
                pos += panjang
        finally:
            mm.close()


def _tanda_besaran(b):
    """Bilangan bulat bertanda gaya GRIB1 (bit pertama = tanda)."""
    v = int.from_bytes(b, "big")
    bit = len(b) * 8 - 1
    return -(v & ((1 << bit) - 1)) if v >> bit else v


def _ibm_float(b):
    s = -1.0 if b[0] & 0x80 else 1.0
    eksp = (b[0] & 0x7F) - 64
    mant = int.from_bytes(b[1:4], "big")
    return s * mant * 2.0 ** -24 * 16.0 ** eksp


_SATUAN_WAKTU_GRIB1 = {0: 1 / 60, 1: 1, 2: 24, 10: 3, 11: 6, 12: 12, 13: 0.25, 14: 0.5, 254: 1 / 3600}


def _grib1_cepat(m):
    """Urai header GRIB1 dengan packing sederhana (format umum ERA5 dari CDS).
    Mengembalikan dict, atau None kalau format di luar jalur cepat."""
    pl = int.from_bytes(m[8:11], "big")
    pds = m[8:8 + pl]
    if pl < 28:
        return None
    bendera = pds[7]
    if not bendera & 0x80:                       # tanpa GDS, tidak bisa
        return None
    satuan = _SATUAN_WAKTU_GRIB1.get(pds[17])
    p1, p2, tri = pds[18], pds[19], pds[20]
    if satuan is None or tri not in (0, 1, 2, 3, 4, 5, 10):
        return None
    if tri == 10:
        geser = (p1 * 256 + p2) * satuan
    elif tri in (2, 3, 4, 5):
        geser = p2 * satuan
    elif tri == 0:
        geser = p1 * satuan
    else:
        geser = 0
    tahun = (pds[24] - 1) * 100 + pds[12]
    try:
        waktu = pd.Timestamp(tahun, pds[13], pds[14], pds[15], pds[16]) + pd.Timedelta(hours=geser)
    except ValueError:
        return None

    off = 8 + pl
    gl = int.from_bytes(m[off:off + 3], "big")
    gds = m[off:off + gl]
    off += gl
    bitmap = None
    if bendera & 0x40:
        bl = int.from_bytes(m[off:off + 3], "big")
        if int.from_bytes(m[off + 4:off + 6], "big") != 0:   # bitmap bawaan, jarang
            return None
        bitmap = m[off + 6:off + bl]
        off += bl
    if m[off + 3] & 0xF0:                        # bukan grid point simple packing
        return None
    return {
        "header": (pds[3], pds[8], pds[9], pds[10], pds[11], gds),
        "waktu": waktu,
        "langkah": float(geser),
        "D": _tanda_besaran(pds[26:28]),
        "E": _tanda_besaran(m[off + 4:off + 6]),
        "R": _ibm_float(m[off + 6:off + 10]),
        "nb": m[off + 10],
        "data": off + 11,
        "bitmap": bitmap,
    }


def _nilai_grib1(m, c, idx, cache_bitmap):
    """Ambil satu nilai dari data simple packing tanpa membongkar seluruh peta."""
    p = idx
    bm = c["bitmap"]
    if bm is not None:
        if not (bm[idx >> 3] >> (7 - (idx & 7))) & 1:
            return np.nan
        kunci = (bytes(bm), idx)
        p = cache_bitmap.get(kunci)
        if p is None:
            bit = np.unpackbits(np.frombuffer(bm, dtype=np.uint8))
            p = int(bit[:idx].sum())
            cache_bitmap[kunci] = p
    nb = c["nb"]
    x = 0
    if nb:
        posbit = p * nb
        awal = c["data"] + (posbit >> 3)
        nbyte = ((posbit & 7) + nb + 7) >> 3
        x = (int.from_bytes(m[awal:awal + nbyte], "big") >> (nbyte * 8 - (posbit & 7) - nb)) & ((1 << nb) - 1)
    return (c["R"] + x * 2.0 ** c["E"]) / 10.0 ** c["D"]


def baca_titik_grib(path, nama_tampil, lat, lon, lewati_kosong=True, kabar=None):
    """Baca deret waktu di grid terdekat dari satu file GRIB.

    Jalur cepat: message GRIB1 simple packing (format ERA5 dari CDS) diurai
    langsung dengan Python, hanya mengambil satu nilai per message. ecCodes
    cukup dipakai sekali untuk setiap kombinasi variabel/level/grid baru.
    Jalur biasa: format lain (GRIB2, packing lain) dibaca lewat ecCodes.
    kabar: fungsi opsional kabar(fraksi_0_sampai_1) untuk progress bar."""
    import eccodes as ec

    ukuran = max(os.path.getsize(path), 1)
    geo = {}            # md5 grid -> info geometri
    pilihan = {}        # (md5, variabel) -> (indeks, catatan)
    daftar_header = {}  # header GRIB1 mentah -> info variabel
    cache_bitmap = {}
    nilai, meta = {}, {}
    peringatan = set()
    jumlah_ganda = 0
    n = 0

    def info_dari_handle(h):
        """Metadata satu message lewat ecCodes, plus pemilihan grid."""
        var = _ambil(ec, h, "cfVarName", "unknown")
        if var in (None, "unknown", "~"):
            var = _ambil(ec, h, "shortName", "unknown")
        if var in (None, "unknown", "~"):
            var = f"param{_ambil(ec, h, 'paramId', n)}"
        md5 = _ambil(ec, h, "md5GridSection", "grid")
        if md5 not in geo:
            geo[md5] = _geometri(ec, h, lat, lon)
            if geo[md5]["peringatan"]:
                peringatan.add(f"{nama_tampil}: {geo[md5]['peringatan']}")
        g = geo[md5]
        ada_bitmap = bool(_ambil(ec, h, "bitmapPresent", 0))
        hilang = float(_ambil(ec, h, "missingValue", 9999))
        if (md5, var) not in pilihan:
            idx, catatan = g["idx0"], ""
            if lewati_kosong and ada_bitmap:
                isi = ec.codes_get_values(h) != hilang
                if isi.any() and not isi[idx]:
                    idx = int(np.argmin(np.where(isi, g["d"], np.inf)))
                    catatan = (f"Grid terdekat ({g['lats'][g['idx0']]:g}, {g['lons'][g['idx0']]:g}) "
                               f"tidak berisi data, dipakai grid berisi terdekat.")
            pilihan[(md5, var)] = (idx, catatan)
        label = _label_level(ec, h)
        return {
            "var": var, "md5": md5, "idx": pilihan[(md5, var)][0], "label": label,
            "kunci": f"{var}@{label}" if label else var,
            "ada_bitmap": ada_bitmap, "hilang": hilang,
            "satuan": _ambil(ec, h, "units", "") or "",
            "nama_panjang": _ambil(ec, h, "name", var) or var,
            "jenis_langkah": _ambil(ec, h, "stepType", "") or "",
        }

    def simpan(info, waktu, v, langkah):
        nonlocal jumlah_ganda
        kunci = info["kunci"]
        deret = nilai.setdefault(kunci, {})
        if waktu in deret:
            jumlah_ganda += 1
        else:
            deret[waktu] = v
        if kunci not in meta:
            meta[kunci] = {
                "nama": info["var"], "level": info["label"], "satuan": info["satuan"],
                "nama_panjang": info["nama_panjang"], "jenis_langkah": info["jenis_langkah"],
                "langkah_maks_jam": langkah, "md5": info["md5"],
            }
        elif langkah > meta[kunci]["langkah_maks_jam"]:
            meta[kunci]["langkah_maks_jam"] = langkah

    def lewat_eccodes(h):
        info = info_dari_handle(h)
        v = ec.codes_get_double_element(h, "values", info["idx"])
        if info["ada_bitmap"] and v == info["hilang"]:
            v = np.nan
        tgl = int(_ambil(ec, h, "validityDate"))
        jam = int(_ambil(ec, h, "validityTime"))
        waktu = pd.Timestamp(tgl // 10000, tgl // 100 % 100, tgl % 100, jam // 100, jam % 100)
        try:
            langkah = float(_ambil(ec, h, "endStep", 0))
        except (TypeError, ValueError):
            langkah = 0.0
        simpan(info, waktu, v, langkah)

    for pos, m in _pesan_grib(path):
        if m is None:
            # Format yang tidak bisa dipotong manual: baca sisa file lewat ecCodes
            with open(path, "rb") as f:
                f.seek(pos)
                while True:
                    h = ec.codes_grib_new_from_file(f)
                    if h is None:
                        break
                    try:
                        lewat_eccodes(h)
                    finally:
                        ec.codes_release(h)
                    n += 1
                    if kabar and n % 250 == 0:
                        kabar(min(f.tell() / ukuran, 1.0))
            break

        c = _grib1_cepat(m) if m[7] == 1 else None
        if c is not None:
            info = daftar_header.get(c["header"])
            if info is None:
                h = ec.codes_new_from_message(bytes(m))
                try:
                    info = info_dari_handle(h)
                finally:
                    ec.codes_release(h)
                daftar_header[c["header"]] = info
            simpan(info, c["waktu"], _nilai_grib1(m, c, info["idx"], cache_bitmap), c["langkah"])
        else:
            h = ec.codes_new_from_message(bytes(m))
            try:
                lewat_eccodes(h)
            finally:
                ec.codes_release(h)

        n += 1
        if kabar and n % 2000 == 0:
            kabar(min(pos / ukuran, 1.0))

    if n == 0:
        raise ValueError(f"{nama_tampil}: tidak ada message GRIB yang terbaca. Pastikan file berformat GRIB.")
    if jumlah_ganda:
        peringatan.add(f"{nama_tampil}: {jumlah_ganda} nilai ganda (variabel, level, dan waktu sama) diabaikan.")

    seri = {}
    for kunci, deret in nilai.items():
        s = pd.Series(deret, dtype=float).sort_index()
        s.index.name = "waktu"
        seri[kunci] = s

    kelompok = {}
    for (md5, var), (idx, catatan) in pilihan.items():
        kelompok.setdefault((md5, idx, catatan), []).append(var)
    grid = []
    for (md5, idx, catatan), daftar_var in kelompok.items():
        g = geo[md5]
        lon_g = float(g["lons"][idx])
        if lon_g > 180:
            lon_g -= 360
        grid.append(InfoGrid(nama_tampil, ", ".join(daftar_var), float(g["lats"][idx]), lon_g,
                             float(g["d"][idx]), g["resolusi"], catatan))
    if kabar:
        kabar(1.0)
    return seri, meta, grid, sorted(peringatan)


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
    keluar, satuan, jenis, konversi, angin = {}, {}, {}, [], []

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
                angin.append({"nama": dasar, "kec": f"Kecepatan {dasar} (m/s)", "arah": f"Arah {dasar} (°)"})
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
    return data, satuan, jenis, konversi, angin


def proses(daftar_file, lat, lon, zona="UTC", mode_akumulasi="otomatis", lewati_kosong=True, kabar=None):
    """daftar_file: list (path, nama_tampil). kabar(fraksi, teks) untuk progress."""
    semua_seri, meta, grid, peringatan = {}, {}, [], []
    total = len(daftar_file)
    for k, (path, nama) in enumerate(daftar_file):
        sub = (lambda fr, k=k, nama=nama: kabar((k + fr) / total, f"Membaca {nama} ({fr:.0%})")) if kabar else None
        s, m, g, p = baca_titik_grib(path, nama, lat, lon, lewati_kosong, sub)
        for k, v in s.items():
            semua_seri[k] = v if k not in semua_seri else semua_seri[k].combine_first(v)
        meta.update(m)
        grid.extend(g)
        peringatan.extend(p)
    if not semua_seri:
        raise ValueError("Tidak ada variabel yang berhasil diambil dari file.")

    mentah = pd.concat(semua_seri, axis=1).sort_index()
    mentah.index.name = "waktu"
    data, satuan, jenis, konversi, angin = terjemahkan(mentah, meta, mode_akumulasi)

    geser = pd.Timedelta(hours=ZONA_WAKTU[zona])
    data.index = data.index + geser
    data = data.dropna(how="all")

    return Hasil(data=data, mentah=mentah, meta=meta, satuan=satuan, jenis=jenis,
                 konversi=konversi, grid=grid, peringatan=peringatan, zona=zona, angin=angin,
                 lat_input=lat, lon_input=lon, sumber=[n for _, n in daftar_file])


# --------------------------------------------------------------------------
# Grafik deret waktu dan windrose
# --------------------------------------------------------------------------
RESOLUSI = {"Per jam": None, "Per hari": "D", "Per bulan": "MS", "Per tahun": "YS"}

# Kelas kecepatan angin (m/s) mengikuti pembagian umum WRPLOT
TENANG = 0.5
KELAS_ANGIN = [(0.5, 2.1), (2.1, 3.6), (3.6, 5.7), (5.7, 8.8), (8.8, 11.1), (11.1, np.inf)]
LABEL_KELAS = ["0,5-2,1", "2,1-3,6", "3,6-5,7", "5,7-8,8", "8,8-11,1", "≥ 11,1"]
WARNA_KELAS = ["#8cc5e3", "#3a8fc7", "#0b4f8a", "#fdae6b", "#e6550d", "#a63603"]
SEKTOR = {
    16: ["U", "UTL", "TL", "TTL", "T", "TTG", "TG", "STG", "S", "SBD", "BD", "BBD", "B", "BBL", "BL", "UBL"],
    8: ["U", "TL", "T", "TG", "S", "BD", "B", "BL"],
}
NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
              "Agustus", "September", "Oktober", "November", "Desember"]
MUSIM = {
    "Des-Jan-Feb (DJF)": [12, 1, 2],
    "Mar-Apr-Mei (MAM)": [3, 4, 5],
    "Jun-Jul-Agu (JJA)": [6, 7, 8],
    "Sep-Okt-Nov (SON)": [9, 10, 11],
}


def resolusi_otomatis(index) -> str:
    """Pilih resolusi grafik dari panjang rentang data."""
    if len(index) < 2:
        return "Per jam"
    rentang = index.max() - index.min()
    if rentang <= pd.Timedelta(days=7):
        return "Per jam"
    if rentang <= pd.Timedelta(days=120):
        return "Per hari"
    if rentang <= pd.Timedelta(days=3 * 366):
        return "Per bulan"
    return "Per tahun"


KOLOM_AGREGASI = ["waktu", "label", "nilai", "min", "maks", "p10", "p90", "n", "kelengkapan"]


def agregasi(s: pd.Series, jenis: str, resolusi: str) -> pd.DataFrame:
    """Ringkas deret waktu per periode. 'nilai' = total untuk variabel akumulasi,
    rata-rata untuk lainnya. p10/p90 = persentil 10 dan 90 data di periode itu."""
    s = s.dropna().sort_index()
    frek = RESOLUSI[resolusi]
    if s.empty:
        return pd.DataFrame(columns=KOLOM_AGREGASI)
    if frek is None:
        return pd.DataFrame({"waktu": s.index, "label": s.index.strftime("%Y-%m-%d %H:%M"),
                             "nilai": s.values, "min": s.values, "maks": s.values,
                             "p10": s.values, "p90": s.values, "n": 1, "kelengkapan": 1.0})

    g = s.resample(frek)
    df = pd.DataFrame({
        "nilai": g.sum(min_count=1) if jenis == "jumlah" else g.mean(),
        "min": g.min(), "maks": g.max(), "p10": g.quantile(0.1), "p90": g.quantile(0.9),
        "n": g.count(),
    }).dropna(subset=["nilai"])

    # Kelengkapan: jumlah data dibanding yang seharusnya ada pada periode itu
    langkah = s.index.to_series().diff().median()
    if pd.isna(langkah) or langkah <= pd.Timedelta(0):
        langkah = pd.Timedelta(hours=1)
    ofs = pd.tseries.frequencies.to_offset(frek)
    harusnya = [(t + ofs - t) / langkah for t in df.index]
    df["kelengkapan"] = np.clip(df["n"].values / np.maximum(harusnya, 1), 0, 1)

    if frek == "MS":
        df["label"] = [f"{NAMA_BULAN_PENDEK[t.month - 1]} {t.year}" for t in df.index]
    else:
        df["label"] = df.index.strftime({"D": "%Y-%m-%d", "YS": "%Y"}[frek])
    df.index.name = "waktu"
    return df.reset_index()[KOLOM_AGREGASI]


def grafik_deret(df: pd.DataFrame, nama: str, satuan: str, jenis: str, resolusi: str):
    """Grafik Altair: batang untuk total akumulasi, garis rata-rata dengan
    pita persentil 10-90 untuk variabel lain."""
    import altair as alt

    alt.data_transformers.disable_max_rows()
    df = df.copy()
    df["lengkap"] = np.where(df["kelengkapan"] >= 0.9, "Lengkap", "Belum lengkap")
    df["kelengkapan_persen"] = (df["kelengkapan"] * 100).round(0)

    per_jam = RESOLUSI[resolusi] is None
    sumbu_x = (alt.X("waktu:T", title=None) if resolusi in ("Per jam", "Per hari")
               else alt.X("label:O", title=None, sort=None, axis=alt.Axis(labelAngle=0 if resolusi == "Per tahun" else -45, labelOverlap="greedy")))
    judul_y = f"{'Total' if jenis == 'jumlah' else 'Rata-rata'} {nama.split(' (')[0].lower()} ({satuan})" if not per_jam \
        else f"{nama.split(' (')[0]} ({satuan})"
    tip = [alt.Tooltip("label:N", title="Periode"),
           alt.Tooltip("nilai:Q", title="Total" if jenis == "jumlah" else "Rata-rata", format=".2f")]
    if not per_jam:
        tip += [alt.Tooltip("p10:Q", title="Persentil 10", format=".2f"),
                alt.Tooltip("p90:Q", title="Persentil 90", format=".2f"),
                alt.Tooltip("min:Q", title="Minimum", format=".2f"),
                alt.Tooltip("maks:Q", title="Maksimum", format=".2f"),
                alt.Tooltip("n:Q", title="Jumlah data"),
                alt.Tooltip("kelengkapan_persen:Q", title="Kelengkapan (%)")]

    dasar = alt.Chart(df).encode(x=sumbu_x)
    if jenis == "jumlah":
        grafik = dasar.mark_bar(opacity=0.9).encode(
            y=alt.Y("nilai:Q", title=judul_y),
            opacity=alt.Opacity("lengkap:N", scale=alt.Scale(domain=["Lengkap", "Belum lengkap"], range=[0.9, 0.4]),
                                legend=alt.Legend(title=None, orient="top")),
            tooltip=tip)
    elif per_jam:
        grafik = dasar.mark_line(strokeWidth=1).encode(y=alt.Y("nilai:Q", title=judul_y), tooltip=tip)
    else:
        pita = dasar.mark_area(opacity=0.25).encode(y=alt.Y("p10:Q", title=judul_y), y2="p90:Q")
        garis = dasar.mark_line(strokeWidth=2).encode(y="nilai:Q")
        titik = dasar.mark_point(filled=True, size=60).encode(
            y="nilai:Q", tooltip=tip,
            shape=alt.Shape("lengkap:N", scale=alt.Scale(domain=["Lengkap", "Belum lengkap"], range=["circle", "triangle"]),
                            legend=alt.Legend(title=None, orient="top")))
        grafik = pita + garis + titik
    return grafik.properties(height=360)


def saring_periode(index, pilihan: str):
    """Mask boolean untuk filter windrose: semua data, musim, bulan, atau tahun."""
    if pilihan == "Semua data":
        return np.ones(len(index), dtype=bool)
    if pilihan in MUSIM:
        return index.month.isin(MUSIM[pilihan])
    if pilihan in NAMA_BULAN:
        return index.month == NAMA_BULAN.index(pilihan) + 1
    if pilihan.startswith("Tahun "):
        return index.year == int(pilihan.split()[1])
    return np.ones(len(index), dtype=bool)


def tabel_windrose(kec: pd.Series, arah: pd.Series, n_sektor=16):
    """Persentase kejadian per sektor arah x kelas kecepatan.
    Mengembalikan (tabel %, persen tenang, jumlah data)."""
    ok = kec.notna() & arah.notna()
    kec, arah = kec[ok].values, arah[ok].values
    total = len(kec)
    label = SEKTOR[n_sektor]
    tabel = pd.DataFrame(0.0, index=label, columns=LABEL_KELAS)
    if total == 0:
        return tabel, 0.0, 0
    lebar = 360 / n_sektor
    sektor = (np.floor(((arah + lebar / 2) % 360) / lebar).astype(int)) % n_sektor
    for (a, b), nama in zip(KELAS_ANGIN, LABEL_KELAS):
        m = (kec >= a) & (kec < b)
        tabel[nama] = np.bincount(sektor[m], minlength=n_sektor) / total * 100
    tenang = float((kec < TENANG).sum() / total * 100)
    return tabel, tenang, total


def gambar_windrose(tabel: pd.DataFrame, judul=""):
    import plotly.graph_objects as go

    n = len(tabel)
    sudut = [k * 360 / n for k in range(n)]
    fig = go.Figure()
    for nama, warna in zip(tabel.columns, WARNA_KELAS):
        if tabel[nama].sum() <= 0:
            continue
        fig.add_trace(go.Barpolar(
            r=tabel[nama].values, theta=sudut, width=[360 / n * 0.92] * n,
            name=f"{nama} m/s", marker_color=warna, marker_line_width=0,
            customdata=tabel.index,
            hovertemplate="%{customdata}: %{r:.2f}%<extra>" + nama + " m/s</extra>"))
    fig.update_layout(
        title=judul, height=520, margin=dict(t=60, b=30, l=30, r=30),
        legend=dict(title="Kecepatan", orientation="v"),
        polar=dict(
            bargap=0,
            angularaxis=dict(direction="clockwise", rotation=90, tickmode="array",
                             tickvals=sudut, ticktext=list(tabel.index)),
            radialaxis=dict(ticksuffix="%", angle=90, tickangle=90),
        ),
    )
    return fig


NAMA_BULAN_PENDEK = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]


def label_bulan(t) -> str:
    return f"{NAMA_BULAN_PENDEK[t.month - 1]} {t.year}"


def label_tanggal(t) -> str:
    return f"{t.day} {NAMA_BULAN_PENDEK[t.month - 1]} {t.year}"


def saring_rentang(index, mulai, akhir):
    """Mask boolean untuk data dalam rentang [mulai, akhir], keduanya inklusif."""
    return (index >= pd.Timestamp(mulai)) & (index <= pd.Timestamp(akhir))


def _gaya_mpl():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 10,
        "axes.spines.top": False, "axes.spines.right": False,
        "axes.edgecolor": "#7f7f7f", "axes.labelcolor": "#333333",
        "xtick.color": "#333333", "ytick.color": "#333333",
    })
    return plt


def _simpan_png(fig, plt) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return buf.getvalue()


def png_deret(df: pd.DataFrame, nama: str, satuan: str, jenis: str, resolusi: str,
              judul: str, catatan: str) -> bytes:
    """Versi gambar (PNG, latar putih) dari grafik deret waktu, siap untuk laporan."""
    plt = _gaya_mpl()
    import matplotlib.dates as mdates

    warna, pucat = "#2171b5", "#9ecae1"
    fig, ax = plt.subplots(figsize=(11, 4.8))
    lengkap = (df["kelengkapan"] >= 0.9).values
    kategori = resolusi in ("Per bulan", "Per tahun")
    x = np.arange(len(df)) if kategori else pd.to_datetime(df["waktu"]).values

    nama_pendek = nama.split(" (")[0]
    if jenis == "jumlah" and (~lengkap).any():
        catatan = catatan + "\nBatang pucat: data pada periode itu belum lengkap, sehingga totalnya lebih kecil dari seharusnya."
    if jenis == "jumlah":
        lebar = 0.8 if kategori else (0.8 if resolusi == "Per hari" else 0.8 / 24)
        ax.bar(x, df["nilai"], width=lebar, color=[warna if k else pucat for k in lengkap])
        ax.set_ylabel(f"Total {nama_pendek.lower()} ({satuan})" if resolusi != "Per jam" else f"{nama_pendek} ({satuan})")
    elif resolusi == "Per jam":
        ax.plot(x, df["nilai"], color=warna, lw=0.8)
        ax.set_ylabel(f"{nama_pendek} ({satuan})")
    else:
        ax.fill_between(x, df["p10"], df["p90"], color=warna, alpha=0.18, lw=0, label="Persentil 10-90")
        ax.plot(x, df["nilai"], color=warna, lw=2, label="Rata-rata")
        if lengkap.any():
            ax.scatter(np.asarray(x)[lengkap], df["nilai"].values[lengkap], color=warna, s=28, zorder=3)
        if (~lengkap).any():
            ax.scatter(np.asarray(x)[~lengkap], df["nilai"].values[~lengkap], marker="^", s=50, zorder=3,
                       facecolor="white", edgecolor=warna, linewidth=1.5, label="Data belum lengkap")
        ax.set_ylabel(f"Rata-rata {nama_pendek.lower()} ({satuan})")
        ax.legend(loc="lower left", bbox_to_anchor=(0.0, 1.0), ncol=3, frameon=False, fontsize=9,
                  borderaxespad=0.2, handlelength=1.8)

    miring = kategori and resolusi == "Per bulan" and len(df) > 12
    if kategori:
        langkah = max(1, int(np.ceil(len(df) / 24)))
        ax.set_xticks(x[::langkah])
        ax.set_xticklabels(df["label"].values[::langkah], rotation=45 if miring else 0,
                           ha="right" if miring else "center")
        ax.set_xlim(-0.6, len(df) - 0.4)
    else:
        from matplotlib.ticker import FuncFormatter
        waktu = pd.to_datetime(df["waktu"])
        hari = (waktu.max() - waktu.min()) / pd.Timedelta(days=1) if len(waktu) else 0

        def fmt(v, _):
            t = mdates.num2date(v)
            bln = NAMA_BULAN_PENDEK[t.month - 1]
            if resolusi == "Per jam" and hari <= 3:
                return f"{t:%H:%M}\n{t.day} {bln}"
            if hari <= 150:
                return f"{t.day} {bln}"
            return f"{bln} {t.year}"
        if resolusi == "Per hari" and hari <= 14:
            ax.xaxis.set_major_locator(mdates.DayLocator())
        else:
            ax.xaxis.set_major_locator(mdates.AutoDateLocator(maxticks=12))
        ax.xaxis.set_major_formatter(FuncFormatter(fmt))
    if jenis == "jumlah" or df["nilai"].min() >= 0:
        ax.set_ylim(bottom=0)
    ax.grid(axis="y", color="#d9d9d9", lw=0.6)
    ax.set_axisbelow(True)
    ax.set_title(judul, loc="left", fontsize=13, fontweight="bold",
                 pad=30 if (jenis != "jumlah" and resolusi != "Per jam") else 12)
    turun = -70 if miring else (-34 if kategori else (-48 if resolusi == "Per jam" and hari <= 3 else -36))
    ax.annotate(catatan, xy=(0, 0), xycoords="axes fraction", xytext=(0, turun), textcoords="offset points",
                fontsize=8, color="#7f7f7f", ha="left", va="top")
    return _simpan_png(fig, plt)


def png_windrose(tabel: pd.DataFrame, judul: str, catatan: str) -> bytes:
    """Versi gambar (PNG, latar putih) dari windrose."""
    plt = _gaya_mpl()
    from matplotlib.ticker import FuncFormatter, MaxNLocator

    n = len(tabel)
    th = np.deg2rad(np.arange(n) * 360 / n)
    fig = plt.figure(figsize=(9, 7.6))
    ax = fig.add_subplot(projection="polar")
    ax.set_theta_zero_location("N")
    ax.set_theta_direction(-1)
    bawah = np.zeros(n)
    for nama, warna in zip(tabel.columns, WARNA_KELAS):
        v = tabel[nama].values
        if v.sum() <= 0:
            continue
        ax.bar(th, v, width=2 * np.pi / n * 0.92, bottom=bawah, color=warna,
               edgecolor="white", linewidth=0.6, label=f"{nama} m/s", zorder=3)
        bawah = bawah + v
    ax.set_xticks(th)
    ax.set_xticklabels(tabel.index, fontsize=10)
    ax.yaxis.set_major_locator(MaxNLocator(4))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}%"))
    ax.set_rlabel_position(360 / n / 2)
    ax.tick_params(axis="y", labelsize=8, colors="#555555")
    ax.grid(color="#cccccc", lw=0.6)
    ax.set_axisbelow(True)
    ax.spines["polar"].set_color("#bbbbbb")
    ax.legend(title="Kecepatan", loc="upper left", bbox_to_anchor=(1.08, 1.0), frameon=False, fontsize=9)
    ax.set_title(judul, loc="center", fontsize=13, fontweight="bold", pad=28)
    fig.text(0.5, 0.01, catatan, fontsize=8, color="#7f7f7f", ha="center", va="bottom", linespacing=1.5)
    return _simpan_png(fig, plt)


# --------------------------------------------------------------------------
# Windrose di atas peta
# --------------------------------------------------------------------------
PETA_DASAR = {
    "Peta jalan": {
        "url": "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}.png",
        "atribusi": "© OpenStreetMap contributors © CARTO",
    },
    "Citra satelit": {
        "url": "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        "atribusi": "Tiles © Esri, Maxar, Earthstar Geographics",
    },
}


def titik_tujuan(lat, lon, arah_deg, jarak_km):
    """Koordinat titik yang berjarak jarak_km dari (lat, lon) ke arah arah_deg (dari utara)."""
    d = np.asarray(jarak_km, dtype=float) / R_BUMI_KM
    b = np.radians(arah_deg)
    p1, l1 = np.radians(lat), np.radians(lon)
    p2 = np.arcsin(np.sin(p1) * np.cos(d) + np.cos(p1) * np.sin(d) * np.cos(b))
    l2 = l1 + np.arctan2(np.sin(b) * np.sin(d) * np.cos(p1), np.cos(d) - np.sin(p1) * np.sin(p2))
    return np.degrees(p2), (np.degrees(l2) + 540) % 360 - 180


def _cincin_rapi(maks_persen):
    """2-4 nilai persen bulat untuk lingkaran acuan."""
    for langkah in (1, 2, 2.5, 5, 10, 20, 25, 50):
        if maks_persen / langkah <= 4:
            break
    return [v for v in np.arange(langkah, maks_persen + 1e-9, langkah)]


def bentuk_kelopak(tabel: pd.DataFrame, lat0, lon0, radius_km, tujuan=False, titik_busur=12):
    """Poligon kelopak windrose dalam lintang/bujur.
    Kelopak terpanjang = radius_km. Kalau tujuan=True, kelopak diputar 180°
    sehingga menunjuk ke arah angin bertiup.
    Mengembalikan (daftar kelas, cincin) dengan kelas berisi lat/lon bersela None."""
    n = len(tabel)
    total = tabel.sum(axis=1).values
    maks = float(total.max()) if total.max() > 0 else 1.0
    skala = radius_km / maks
    lebar = 360 / n * 0.9
    kelas, bawah = [], np.zeros(n)
    for nama, warna in zip(tabel.columns, WARNA_KELAS):
        v = tabel[nama].values
        if v.sum() <= 0:
            continue
        lats, lons = [], []
        for k in range(n):
            if v[k] <= 0:
                continue
            pusat = k * 360 / n + (180 if tujuan else 0)
            sudut = np.linspace(pusat - lebar / 2, pusat + lebar / 2, titik_busur)
            r0, r1 = bawah[k] * skala, (bawah[k] + v[k]) * skala
            la1, lo1 = titik_tujuan(lat0, lon0, sudut, np.full(titik_busur, r1))
            if r0 > 0:
                la0, lo0 = titik_tujuan(lat0, lon0, sudut[::-1], np.full(titik_busur, r0))
            else:
                la0, lo0 = np.array([lat0]), np.array([lon0])
            la = np.concatenate([la1, la0, la1[:1]])
            lo = np.concatenate([lo1, lo0, lo1[:1]])
            lats += list(la) + [None]
            lons += list(lo) + [None]
        kelas.append({"nama": nama, "warna": warna, "lat": lats, "lon": lons})
        bawah = bawah + v
    cincin = []
    for p in _cincin_rapi(maks):
        sudut = np.linspace(0, 360, 73)
        la, lo = titik_tujuan(lat0, lon0, sudut, np.full(73, p * skala))
        cincin.append({"persen": p, "lat": la, "lon": lo, "jarak_km": p * skala})
    return kelas, cincin


def _zoom_untuk(lat, lebar_km, lebar_px, ukuran_tile):
    m_per_px = lebar_km * 1000 / lebar_px
    keliling = 40075016.686 * np.cos(np.radians(lat))
    return float(np.log2(keliling / (ukuran_tile * m_per_px)))


def _rgba(hex_warna, alpha):
    h = hex_warna.lstrip("#")
    return f"rgba({int(h[0:2], 16)},{int(h[2:4], 16)},{int(h[4:6], 16)},{alpha:.2f})"


def gambar_windrose_peta(tabel, lat0, lon0, lat_g, lon_g, radius_km, tujuan, peta_dasar, judul="", opasitas=0.6):
    """Windrose interaktif (Plotly) di atas peta, berpusat di koordinat input (lat0, lon0).
    (lat_g, lon_g) = grid data asal nilai, ditandai titik hitam. opasitas 0-1 untuk kelopak."""
    import plotly.graph_objects as go

    lon0 = (lon0 + 540) % 360 - 180
    kelas, cincin = bentuk_kelopak(tabel, lat0, lon0, radius_km, tujuan)
    fig = go.Figure()
    for c in cincin:
        fig.add_trace(go.Scattermap(lat=c["lat"], lon=c["lon"], mode="lines", hoverinfo="skip",
                                    line=dict(width=1, color="rgba(60,60,60,0.55)"), showlegend=False))
    for kls in kelas:
        fig.add_trace(go.Scattermap(lat=kls["lat"], lon=kls["lon"], mode="lines", fill="toself",
                                    fillcolor=_rgba(kls["warna"], opasitas),
                                    line=dict(width=1, color="rgba(30,30,30,0.75)"),
                                    name=f"{kls['nama']} m/s", hoverinfo="skip"))

    # Info saat disorot: titik di ujung tiap kelopak
    n = len(tabel)
    total = tabel.sum(axis=1)
    maks = float(total.max()) if total.max() > 0 else 1.0
    arah_ujung = np.arange(n) * 360 / n + (180 if tujuan else 0)
    la, lo = titik_tujuan(lat0, lon0, arah_ujung, total.values / maks * radius_km)
    teks = []
    for k, (sektor, baris) in enumerate(tabel.iterrows()):
        rinci = "<br>".join(f"{kls} m/s: {v:.1f}%" for kls, v in baris.items() if v > 0)
        ke = tabel.index[(k + n // 2) % n]
        teks.append(f"<b>Angin dari {sektor}, bertiup ke {ke}</b>: {baris.sum():.1f}%<br>{rinci}")
    fig.add_trace(go.Scattermap(lat=la, lon=lo, mode="markers", marker=dict(size=8, opacity=0),
                                text=teks, hovertemplate="%{text}<extra></extra>", showlegend=False))
    for c in cincin:
        fig.add_trace(go.Scattermap(lat=[c["lat"][0]], lon=[c["lon"][0]], mode="text",
                                    text=[f"{c['persen']:g}%"], textposition="top center",
                                    textfont=dict(size=11, color="#333333"), hoverinfo="skip", showlegend=False))
    fig.add_trace(go.Scattermap(lat=[lat_g], lon=[lon_g], mode="markers", name="Grid data (asal nilai)",
                                marker=dict(size=8, color="#111111"),
                                hovertemplate=f"Grid data ERA5<br>{lat_g:.2f}, {lon_g:.2f}<extra></extra>"))
    fig.add_trace(go.Scattermap(lat=[lat0], lon=[lon0], mode="markers", name="Koordinat input",
                                marker=dict(size=12, color="#d62728"),
                                hovertemplate=f"Koordinat input<br>{lat0:.4f}, {lon0:.4f}<extra></extra>"))

    gaya = PETA_DASAR.get(peta_dasar, PETA_DASAR["Peta jalan"])
    if peta_dasar == "Citra satelit":
        peta = dict(style="white-bg", layers=[dict(below="traces", sourcetype="raster",
                                                   sourceattribution=gaya["atribusi"], source=[gaya["url"]])])
    else:
        peta = dict(style="carto-positron")
    peta.update(center=dict(lat=lat0, lon=lon0), zoom=_zoom_untuk(lat0, radius_km * 2.8, 640, 512))
    fig.update_layout(map=peta, height=640, margin=dict(t=50 if judul else 10, b=10, l=10, r=10),
                      title=judul, legend=dict(title="Kecepatan", bgcolor="rgba(255,255,255,0.85)",
                                                font=dict(color="#222222"), x=0.01, y=0.99))
    return fig


def _tile(url_pola, z, x, y):
    """Ambil satu tile peta (PNG/JPG). Mengembalikan gambar PIL atau None."""
    import urllib.request
    from PIL import Image

    n = 2 ** z
    url = url_pola.format(s="abcd"[(x + y) % 4], z=z, x=x % n, y=y)
    req = urllib.request.Request(url, headers={"User-Agent": "grib-ke-excel/1.0 (aplikasi Streamlit)"})
    try:
        with urllib.request.urlopen(req, timeout=8) as r:
            return Image.open(io.BytesIO(r.read())).convert("RGB")
    except Exception:  # noqa: BLE001
        return None


def ambil_peta_dasar(lat0, lon0, lebar_km, peta_dasar="Peta jalan", target_px=1000, pengambil=None):
    """Susun peta dasar persegi berpusat di (lat0, lon0) selebar lebar_km.
    Mengembalikan (gambar PIL, fungsi lat/lon -> piksel, km per piksel) atau None kalau gagal."""
    from concurrent.futures import ThreadPoolExecutor
    from PIL import Image

    pengambil = pengambil or _tile
    pola = PETA_DASAR.get(peta_dasar, PETA_DASAR["Peta jalan"])["url"]
    z = int(np.clip(np.floor(_zoom_untuk(lat0, lebar_km, target_px, 256)), 2, 17))
    skala = 256 * 2 ** z

    def ke_px(lat, lon):
        lat = np.clip(np.asarray(lat, dtype=float), -85, 85)
        x = (np.asarray(lon, dtype=float) + 180) / 360 * skala
        y = (1 - np.log(np.tan(np.radians(lat)) + 1 / np.cos(np.radians(lat))) / np.pi) / 2 * skala
        return x, y

    cx, cy = ke_px(lat0, lon0)
    km_per_px = 40075.016686 * np.cos(np.radians(lat0)) / skala
    setengah = lebar_km / 2 / km_per_px
    x0, x1, y0, y1 = cx - setengah, cx + setengah, cy - setengah, cy + setengah
    tx = range(int(np.floor(x0 / 256)), int(np.floor(x1 / 256)) + 1)
    ty = range(max(int(np.floor(y0 / 256)), 0), min(int(np.floor(y1 / 256)), 2 ** z - 1) + 1)
    daftar = [(x, y) for y in ty for x in tx]
    if len(daftar) > 64:
        return None
    with ThreadPoolExecutor(max_workers=8) as ex:
        hasil = list(ex.map(lambda t: pengambil(pola, z, t[0], t[1]), daftar))
    if not any(h is not None for h in hasil):
        return None
    kanvas = Image.new("RGB", (len(tx) * 256, len(ty) * 256), (235, 235, 235))
    for (x, y), img in zip(daftar, hasil):
        if img is not None:
            kanvas.paste(img.resize((256, 256)), ((x - tx[0]) * 256, (y - ty[0]) * 256))
    kiri, atas = int(round(x0 - tx[0] * 256)), int(round(y0 - ty[0] * 256))
    sisi = int(round(2 * setengah))
    gambar = kanvas.crop((kiri, atas, kiri + sisi, atas + sisi))

    def lokal(lat, lon):
        x, y = ke_px(lat, lon)
        return x - x0, y - y0
    return gambar, lokal, km_per_px


def png_windrose_peta(tabel, lat0, lon0, lat_g, lon_g, radius_km, tujuan, peta_dasar,
                      judul, catatan, opasitas=0.6, pengambil=None) -> bytes:
    """Versi gambar windrose di atas peta, berpusat di koordinat input (lat0, lon0),
    lengkap dengan arah utara, skala, dan atribusi peta."""
    plt = _gaya_mpl()
    from matplotlib.colors import to_rgba
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch

    lon0 = (lon0 + 540) % 360 - 180
    lebar_km = radius_km * 2.8
    peta = ambil_peta_dasar(lat0, lon0, lebar_km, peta_dasar, pengambil=pengambil)
    fig, ax = plt.subplots(figsize=(9, 9))
    if peta is not None:
        gambar, lokal, km_per_px = peta
        ax.imshow(gambar, extent=(0, gambar.width, gambar.height, 0), zorder=0)
        lebar_px = gambar.width
        atribusi = PETA_DASAR.get(peta_dasar, PETA_DASAR["Peta jalan"])["atribusi"]
    else:
        # Cadangan tanpa peta: proyeksi datar sederhana di sekitar titik
        lebar_px = 1000
        km_per_px = lebar_km / lebar_px

        def lokal(lat, lon):
            dy = (np.asarray(lat, dtype=float) - lat0) * 111.32
            dx = (np.asarray(lon, dtype=float) - lon0) * 111.32 * np.cos(np.radians(lat0))
            return lebar_px / 2 + dx / km_per_px, lebar_px / 2 - dy / km_per_px
        ax.set_facecolor("#f2f2f2")
        atribusi = "Peta dasar tidak dapat dimuat"

    kelas, cincin = bentuk_kelopak(tabel, lat0, lon0, radius_km, tujuan)
    for c in cincin:
        x, y = lokal(c["lat"], c["lon"])
        ax.plot(x, y, color="#333333", lw=0.8, ls=(0, (4, 3)), alpha=0.7, zorder=2)
        ax.text(x[0], y[0] - 4, f"{c['persen']:g}%", ha="center", va="bottom", fontsize=8, color="#222222",
                zorder=5, bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.75))
    for kls in kelas:
        lat_arr = [v for v in kls["lat"]]
        lon_arr = [v for v in kls["lon"]]
        bagian_lat, bagian_lon = [], []
        for la, lo in zip(lat_arr + [None], lon_arr + [None]):
            if la is None:
                if bagian_lat:
                    x, y = lokal(bagian_lat, bagian_lon)
                    ax.fill(x, y, fc=to_rgba(kls["warna"], opasitas), ec=(0.12, 0.12, 0.12, 0.75), lw=0.7,
                            zorder=3)
                bagian_lat, bagian_lon = [], []
            else:
                bagian_lat.append(la)
                bagian_lon.append(lo)

    gx, gy = lokal(lat_g, lon_g)
    ix, iy = lokal(lat0, lon0)
    ax.scatter([gx], [gy], s=24, color="#111111", edgecolor="white", linewidth=0.8, zorder=6)
    ax.scatter([ix], [iy], s=70, color="#d62728", edgecolor="white", linewidth=1.2, zorder=7)

    ax.set_xlim(0, lebar_px)
    ax.set_ylim(lebar_px, 0)
    ax.set_xticks([])
    ax.set_yticks([])
    for sp in ax.spines.values():
        sp.set_visible(True)
        sp.set_color("#999999")

    # Arah utara
    ax.annotate("", xy=(0.06, 0.95), xytext=(0.06, 0.86), xycoords="axes fraction",
                arrowprops=dict(arrowstyle="-|>", color="#111111", lw=2), zorder=7)
    ax.text(0.06, 0.955, "U", transform=ax.transAxes, ha="center", va="bottom", fontsize=12,
            fontweight="bold", zorder=7, bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.8))
    # Skala jarak
    kandidat = [1, 2, 5, 10, 20, 25, 50, 100, 200, 250, 500]
    km_bar = max([k for k in kandidat if k <= lebar_km / 4] or [1])
    px_bar = km_bar / km_per_px
    x_awal, y_bar = lebar_px * 0.05, lebar_px * 0.95
    ax.plot([x_awal, x_awal + px_bar], [y_bar, y_bar], color="#111111", lw=3, solid_capstyle="butt", zorder=7)
    ax.text(x_awal + px_bar / 2, y_bar - lebar_px * 0.012, f"{km_bar} km", ha="center", va="bottom", fontsize=9,
            zorder=7, bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.8))
    ax.text(0.995, 0.005, atribusi, transform=ax.transAxes, ha="right", va="bottom", fontsize=7, color="#333333",
            zorder=7, bbox=dict(boxstyle="square,pad=0.2", fc="white", ec="none", alpha=0.75))

    pegangan = [Patch(fc=to_rgba(k["warna"], max(opasitas, 0.5)), ec=(0.12, 0.12, 0.12, 0.75), lw=0.7,
                      label=f"{k['nama']} m/s") for k in kelas]
    pegangan += [Line2D([], [], marker="o", ls="", color="#111111", label="Grid data (asal nilai)"),
                 Line2D([], [], marker="o", ls="", color="#d62728", markeredgecolor="white", label="Koordinat input")]
    ax.legend(handles=pegangan, title="Kecepatan", loc="upper right", fontsize=8, title_fontsize=9,
              framealpha=0.9)
    ax.set_title(judul, fontsize=13, fontweight="bold", pad=10)
    ax.annotate(catatan, xy=(0.5, 0), xycoords="axes fraction", xytext=(0, -8), textcoords="offset points",
                fontsize=8, color="#555555", ha="center", va="top", linespacing=1.5)
    return _simpan_png(fig, plt)


def _slug(teks: str) -> str:
    import re
    return re.sub(r"[^a-z0-9]+", "_", teks.lower()).strip("_")


def _angka_id(x, desimal=2) -> str:
    return f"{x:.{desimal}f}".replace(".", ",")


def grid_untuk(h, kolom: str):
    """InfoGrid yang dipakai untuk kolom hasil tertentu."""
    asal = next((asl for kol, asl, _ in h.konversi if kol == kolom), "")
    var = asal.split(",")[0].split("@")[0].strip()
    for x in h.grid:
        if var and var in [v.strip() for v in x.variabel.split(",")]:
            return x
    return h.grid[0] if h.grid else None


def catatan_sumber(h, kolom: str) -> str:
    """Baris sumber untuk gambar: dataset, grid yang dipakai, dan zona waktu."""
    g = grid_untuk(h, kolom)
    teks = "Sumber: ERA5, Copernicus Climate Change Service."
    if g is not None:
        lat = f"{_angka_id(abs(g.lat_grid))}° {'LS' if g.lat_grid < 0 else 'LU'}"
        lon = f"{_angka_id(abs(g.lon_grid))}° {'BB' if g.lon_grid < 0 else 'BT'}"
        teks += f" Grid {lat}, {lon} ({_angka_id(g.jarak_km, 1)} km dari titik input)."
    return teks + f" Waktu {h.zona.split(' ')[0]}."


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

    # ---- Sheet Windrose (formula COUNTIFS dari sheet Data)
    ww = None
    if h.angin:
        ww = wb.create_sheet("Windrose")
        n_kls = len(KELAS_ANGIN)
        kol_total = 4 + n_kls                       # kolom "Total"
        r = 1
        for a in h.angin:
            kec = f"Data!${huruf[a['kec']]}$2:${huruf[a['kec']]}${n_akhir}"
            arah = f"Data!${huruf[a['arah']]}$2:${huruf[a['arah']]}${n_akhir}"

            ww.cell(r, 1, f"Windrose {a['nama']}").style = "label"
            ww.row_dimensions[r].height = 22
            r += 1
            ww.cell(r, 1, "Jumlah data").style = "teks"
            sel_n = ww.cell(r, 2, f"=COUNT({kec})")
            sel_n.style = "angka"
            sel_n.number_format = "0"
            ref_n = f"$B${r}"
            ww.row_dimensions[r].height = 18
            r += 2

            # Header: nama kelas, lalu batas bawah dan atas kecepatan (m/s)
            baris_h, baris_lo, baris_hi = r, r + 1, r + 2
            for c, t in enumerate(["Arah", "Dari (°)", "Sampai (°)"], start=1):
                ww.cell(baris_h, c, t).style = "th_k"
                ww.cell(baris_lo, c, "Kecepatan dari (m/s)" if c == 1 else None).style = "teks_k" if c == 1 else "strip_k"
                ww.cell(baris_hi, c, "Kecepatan sampai (m/s)" if c == 1 else None).style = "teks_k" if c == 1 else "strip_k"
            for k, ((lo, hi), lab) in enumerate(zip(KELAS_ANGIN, LABEL_KELAS)):
                c = 4 + k
                ww.cell(baris_h, c, f"{lab} m/s").style = "th_k"
                ww.cell(baris_lo, c, lo).style = "angka_k"
                ww.cell(baris_hi, c, 999 if np.isinf(hi) else hi).style = "angka_k"
                ww.cell(baris_lo, c).number_format = ww.cell(baris_hi, c).number_format = "0.0"
            ww.cell(baris_h, kol_total, "Total").style = "th_k"
            ww.cell(baris_lo, kol_total).style = "strip_k"
            ww.cell(baris_hi, kol_total).style = "strip_k"
            for rr, tinggi in ((baris_h, 30), (baris_lo, 18), (baris_hi, 18)):
                ww.row_dimensions[rr].height = tinggi

            # 16 sektor arah. Sektor utara melewati 360°, jadi dihitung dua bagian.
            awal = baris_hi + 1
            label16 = SEKTOR[16]
            for s_idx, lab in enumerate(label16):
                rr = awal + s_idx
                dari = (s_idx * 22.5 - 11.25) % 360
                sampai = (s_idx * 22.5 + 11.25) % 360
                ww.cell(rr, 1, lab).style = "teks_k"
                ww.cell(rr, 2, dari).style = "angka_k"
                ww.cell(rr, 3, sampai).style = "angka_k"
                for k in range(n_kls):
                    c = 4 + k
                    col = get_column_letter(c)
                    lo, hi = f"{col}${baris_lo}", f"{col}${baris_hi}"
                    syarat_kec = f'{kec},">="&{lo},{kec},"<"&{hi}'
                    f = (f'=IF({ref_n}=0,0,IF($B{rr}<$C{rr},'
                         f'COUNTIFS({arah},">="&$B{rr},{arah},"<"&$C{rr},{syarat_kec}),'
                         f'COUNTIFS({arah},">="&$B{rr},{syarat_kec})+COUNTIFS({arah},"<"&$C{rr},{syarat_kec}))/{ref_n})')
                    ww.cell(rr, c, f).style = "persen_k"
                ww.cell(rr, kol_total,
                        f"=SUM({get_column_letter(4)}{rr}:{get_column_letter(3 + n_kls)}{rr})").style = "persen_k"
                ww.row_dimensions[rr].height = 18

            akhir = awal + len(label16) - 1
            rr = akhir + 1
            ww.cell(rr, 1, "Jumlah per kelas").style = "th_k"
            ww.cell(rr, 2).style = "strip_k"
            ww.cell(rr, 3).style = "strip_k"
            for c in range(4, kol_total + 1):
                col = get_column_letter(c)
                ww.cell(rr, c, f"=SUM({col}{awal}:{col}{akhir})").style = "persen_k"
            ww.row_dimensions[rr].height = 20
            rr += 1
            ww.cell(rr, 1, f"Tenang (< {TENANG:g} m/s)".replace(".", ",")).style = "teks_k"
            ww.cell(rr, 2).style = "strip_k"
            ww.cell(rr, 3).style = "strip_k"
            for c in range(4, kol_total):
                ww.cell(rr, c).style = "strip_k"
            ww.cell(rr, kol_total,
                    f'=IF({ref_n}=0,0,COUNTIFS({kec},"<"&{get_column_letter(4)}${baris_lo})/{ref_n})').style = "persen_k"
            ww.row_dimensions[rr].height = 18
            rr += 1
            ww.cell(rr, 1, "Total").style = "th_k"
            for c in range(2, kol_total):
                ww.cell(rr, c).style = "strip_k"
            ww.cell(rr, kol_total, f"={get_column_letter(kol_total)}{rr - 2}+{get_column_letter(kol_total)}{rr - 1}").style = "persen_k"
            ww.row_dimensions[rr].height = 20
            r = rr + 3

        ww.cell(r - 2, 1, "Persentase dihitung dari seluruh data. Arah = arah datangnya angin, "
                          "batas sektor: dari ≤ arah < sampai.").font = Font(name=FONT, size=8, color="FF7F7F7F")
        ww.column_dimensions["A"].width = 24
        for c in range(2, kol_total + 1):
            ww.column_dimensions[get_column_letter(c)].width = 12

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

    for lembar in [x for x in (wr, wi, ww) if x is not None]:
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

            bar = st.progress(0.0, text="Mulai membaca GRIB...")
            hasil = proses(daftar, float(lat), float(lon), zona, mode_label[mode], lewati,
                           kabar=lambda fr, teks: bar.progress(min(fr, 1.0) * 0.9, text=teks))
            bar.progress(0.92, text="Menulis Excel...")
            xlsx = buat_excel(hasil, mentah)
            bar.empty()
            st.session_state["hasil"] = hasil
            st.session_state["xlsx"] = xlsx
            st.session_state["proses_id"] = st.session_state.get("proses_id", 0) + 1
        except Exception as e:  # noqa: BLE001
            pesan = str(e)
            st.error(f"Gagal memproses file: {pesan}")
            if "eccodes" in pesan.lower() or "ecCodes" in pesan:
                st.info("Library ecCodes belum terpasang. Coba: pip install eccodes")
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

    nama_file = f"era5_{hasil.lat_input:.3f}_{hasil.lon_input:.3f}.xlsx".replace("-", "m")

    # ---------------- Rentang data untuk grafik
    pid = st.session_state.get("proses_id", 0)
    idx_semua = hasil.data.index
    st.subheader("Grafik")
    with st.container(border=True):
        cara = st.radio("Pilih rentang data berdasarkan", ["Bulan", "Tanggal"], horizontal=True, key=f"cara_{pid}")
        c1, c2 = st.columns(2)
        if cara == "Bulan":
            daftar_bulan = pd.period_range(idx_semua.min().to_period("M"), idx_semua.max().to_period("M"), freq="M")
            label = [label_bulan(p.start_time) for p in daftar_bulan]
            dari = c1.selectbox("Dari bulan", label, index=0, key=f"dari_{pid}")
            sampai = c2.selectbox("Sampai bulan", label, index=len(label) - 1, key=f"sampai_{pid}")
            i0, i1 = sorted((label.index(dari), label.index(sampai)))
            mulai, akhir = daftar_bulan[i0].start_time, daftar_bulan[i1].end_time
            teks_rentang = label[i0] if i0 == i1 else f"{label[i0]} s.d. {label[i1]}"
        else:
            awal_d, akhir_d = idx_semua.min().date(), idx_semua.max().date()
            pilih_tgl = c1.date_input("Rentang tanggal", value=(awal_d, akhir_d), min_value=awal_d,
                                      max_value=akhir_d, format="DD/MM/YYYY", key=f"tgl_{pid}")
            if isinstance(pilih_tgl, (tuple, list)):
                d0 = pilih_tgl[0]
                d1 = pilih_tgl[1] if len(pilih_tgl) > 1 else pilih_tgl[0]
            else:
                d0 = d1 = pilih_tgl
            mulai = pd.Timestamp(d0)
            akhir = pd.Timestamp(d1) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            teks_rentang = (label_tanggal(mulai) if d0 == d1
                            else f"{label_tanggal(mulai)} s.d. {label_tanggal(pd.Timestamp(d1))}")
        mask_rentang = saring_rentang(idx_semua, mulai, akhir)
        data = hasil.data[mask_rentang]
        st.caption(f"{len(data):,} data dalam rentang {teks_rentang}. ".replace(",", ".")
                   + "Rentang ini dipakai untuk grafik deret waktu dan windrose; file Excel tetap berisi seluruh data.")

    if data.empty:
        st.info("Tidak ada data pada rentang ini.")
    else:
        png_deret_c = st.cache_data(show_spinner=False)(png_deret)
        png_windrose_c = st.cache_data(show_spinner=False)(png_windrose)
        png_windrose_peta_c = st.cache_data(show_spinner=False, max_entries=20)(png_windrose_peta)
        slug_rentang = _slug(teks_rentang)

        # ---------------- Grafik deret waktu
        angka = [c for c, j in hasil.jenis.items() if j not in ("teks", "arah")]
        if angka:
            st.markdown("#### Deret waktu")
            auto = resolusi_otomatis(data.index)
            c1, c2 = st.columns([2, 1])
            nama = c1.selectbox("Variabel", angka, key=f"var_{pid}")
            pilih_res = c2.selectbox("Resolusi", ["Otomatis"] + list(RESOLUSI), key=f"res_{pid}",
                                     help="Otomatis: per jam sampai 7 hari, per hari sampai 4 bulan, "
                                          "per bulan sampai 3 tahun, per tahun untuk rentang lebih panjang.")
            res = auto if pilih_res == "Otomatis" else pilih_res

            jenis = hasil.jenis[nama]
            df = agregasi(data[nama], jenis, res)
            if df.empty:
                st.info("Variabel ini tidak punya data pada rentang yang dipilih.")
            else:
                if res == "Per jam" and len(df) > 20000:
                    st.caption(f"{len(df):,} titik per jam, grafik mungkin berat di browser.".replace(",", "."))
                st.altair_chart(grafik_deret(df, nama, hasil.satuan[nama], jenis, res), width="stretch")

                ket = [f"Ditampilkan {res.lower()}" + (" (otomatis)." if pilih_res == "Otomatis" else ".")]
                if res != "Per jam":
                    ket.append("Garis: total per periode." if jenis == "jumlah"
                               else "Garis: rata-rata per periode. Pita: rentang persentil 10-90 (80% data berada di dalamnya).")
                    kurang = df.loc[df["kelengkapan"] < 0.9, "label"].tolist()
                    if kurang:
                        daftar = ", ".join(kurang[:6]) + (f", dan {len(kurang) - 6} lainnya" if len(kurang) > 6 else "")
                        ket.append(f"Periode dengan data belum lengkap (ditandai segitiga/batang pucat): {daftar}.")
                if ket:
                    st.caption(" ".join(ket))

                judul = f"{nama.split(' (')[0]}, {teks_rentang} ({res.lower()})"
                catatan_g = catatan_sumber(hasil, nama)
                png = (lambda df=df, nama=nama, jenis=jenis, res=res, judul=judul, catatan_g=catatan_g:
                       png_deret_c(df, nama, hasil.satuan[nama], jenis, res, judul, catatan_g))
                st.download_button("Unduh grafik (PNG)", png, key=f"dl_grafik_{pid}",
                                   file_name=f"grafik_{_slug(nama.split(' (')[0])}_{slug_rentang}_{_slug(res)}.png",
                                   mime="image/png")

        # ---------------- Windrose
        if hasil.angin:
            st.markdown("#### Windrose")
            c1, c2, c3 = st.columns([2, 2, 1])
            if len(hasil.angin) > 1:
                pilih_angin = c1.selectbox("Ketinggian/level", [a["nama"] for a in hasil.angin], key=f"lvl_{pid}")
            else:
                pilih_angin = hasil.angin[0]["nama"]
                c1.text_input("Ketinggian/level", pilih_angin, disabled=True, key=f"lvl_{pid}")
            a = next(x for x in hasil.angin if x["nama"] == pilih_angin)
            semua = "Semua bulan dalam rentang"
            periode = c2.selectbox("Saring lagi (opsional)", [semua] + list(MUSIM) + NAMA_BULAN, key=f"musim_{pid}",
                                   help="Contoh: rentang 2021 s.d. 2025 lalu pilih JJA untuk melihat angin musim kemarau "
                                        "selama lima tahun itu.")
            n_sektor = c3.radio("Sektor", [16, 8], horizontal=True, key=f"sektor_{pid}")

            mask = saring_periode(data.index, periode)
            sub = data.loc[mask]
            tabel, tenang, n = tabel_windrose(sub[a["kec"]], sub[a["arah"]], n_sektor)
            if n == 0:
                st.info("Tidak ada data angin pada pilihan ini. Cek lagi isian \"Saring lagi\", misalnya bulan yang dipilih tidak ada dalam rentang.")
            else:
                keterangan = teks_rentang if periode == semua else f"{teks_rentang}, {periode}"
                per_arah = tabel.sum(axis=1)
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Arah dominan (datang dari)", per_arah.idxmax(), f"{per_arah.max():.1f}% kejadian",
                          delta_color="off")
                m2.metric("Kecepatan rata-rata", f"{sub[a['kec']].mean():.2f} m/s")
                m3.metric("Angin tenang (< 0,5 m/s)", f"{tenang:.1f}%")
                m4.metric("Jumlah data", f"{n:,}".replace(",", "."))

                jumlah = f"{n:,}".replace(",", ".")
                ringkas = f"Angin tenang (< 0,5 m/s): {_angka_id(tenang, 1)}%. Jumlah data: {jumlah}."
                tab_diagram, tab_peta = st.tabs(["Diagram", "Di atas peta"])

                with tab_diagram:
                    st.plotly_chart(gambar_windrose(tabel, f"Windrose {a['nama']}, {keterangan}"), width="stretch")
                    st.caption("Arah menunjukkan dari mana angin datang. Panjang batang = persentase kejadian dari "
                               "seluruh data pada pilihan ini; angin tenang tidak punya arah sehingga tidak masuk "
                               "batang. Tabel windrose seluruh data ada di sheet Windrose pada file Excel.")
                    catatan = catatan_sumber(hasil, a["kec"]) + "\n" + ringkas + " Arah = arah datangnya angin."
                    judul_wr = f"Windrose {a['nama']}\n{keterangan}"
                    png = lambda: png_windrose_c(tabel, judul_wr, catatan)  # noqa: E731
                    st.download_button("Unduh windrose (PNG)", png, key=f"dl_wr_{pid}",
                                       file_name=f"windrose_{_slug(a['nama'])}_{slug_rentang}"
                                                 + ("" if periode == semua else f"_{_slug(periode)}")
                                                 + f"_{n_sektor}arah.png",
                                       mime="image/png")

                with tab_peta:
                    g = grid_untuk(hasil, a["kec"])
                    c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
                    tampil = c1.radio("Kelopak menunjukkan", ["Arah datang angin", "Arah tujuan angin"],
                                      key=f"tujuan_{pid}",
                                      help="Windrose standar menunjukkan dari mana angin datang. Pilih arah tujuan "
                                           "untuk melihat ke mana angin bertiup dari titik ini.")
                    tujuan = tampil == "Arah tujuan angin"
                    peta_dasar = c2.radio("Peta dasar", list(PETA_DASAR), key=f"basemap_{pid}")
                    radius = c3.slider("Panjang kelopak terpanjang (km)", 2, 200, 30, key=f"radius_{pid}",
                                       help="Hanya skala gambar. Panjang kelopak sebanding dengan persentase "
                                            "kejadian, bukan jarak tempuh angin.")
                    opasitas = c4.slider("Kepekatan kelopak (%)", 10, 100, 55, step=5, key=f"opasitas_{pid}",
                                         help="Kecilkan supaya peta di bawahnya lebih terlihat.") / 100
                    st.plotly_chart(
                        gambar_windrose_peta(tabel, hasil.lat_input, hasil.lon_input, g.lat_grid, g.lon_grid,
                                             radius, tujuan, peta_dasar, opasitas=opasitas),
                        width="stretch")
                    dominan = per_arah.idxmax()
                    ke = tabel.index[(list(tabel.index).index(dominan) + len(tabel) // 2) % len(tabel)]
                    st.caption(
                        ("Kelopak menunjuk ke arah angin bertiup. " if tujuan
                         else "Kelopak menunjuk ke arah datangnya angin (standar windrose). ")
                        + f"Angin paling sering datang dari {dominan} dan bertiup ke {ke}. "
                        + "Windrose berpusat di koordinat input (titik merah). Nilainya diambil dari grid ERA5 "
                        + f"terdekat (titik hitam, {_angka_id(g.jarak_km, 1)} km dari input). "
                        + "Data ERA5 mewakili rata-rata satu kotak grid, jadi efek lokal seperti angin darat-laut "
                          "skala kecil atau bangunan tidak tergambar.")
                    catatan = (catatan_sumber(hasil, a["kec"]) + " Windrose berpusat di koordinat input.\n"
                               + ("Kelopak menunjukkan arah tujuan angin (ke mana angin bertiup). " if tujuan
                                  else "Kelopak menunjukkan arah datang angin. ")
                               + ringkas)
                    def png_peta(tabel=tabel, g=g, radius=radius, tujuan=tujuan, peta_dasar=peta_dasar,
                                 judul=f"Windrose {a['nama']}\n{keterangan}", catatan=catatan, opasitas=opasitas):
                        return png_windrose_peta_c(tabel, hasil.lat_input, hasil.lon_input, g.lat_grid,
                                                   g.lon_grid, radius, tujuan, peta_dasar, judul, catatan, opasitas)
                    st.download_button("Unduh windrose di peta (PNG)", png_peta, key=f"dl_wrp_{pid}",
                                       file_name=f"windrose_peta_{_slug(a['nama'])}_{slug_rentang}"
                                                 + ("" if periode == semua else f"_{_slug(periode)}")
                                                 + ("_arah_tujuan" if tujuan else "_arah_datang") + ".png",
                                       mime="image/png")

    st.subheader("Pratinjau data")
    st.dataframe(hasil.data.head(500), width="stretch")

    st.download_button("Unduh Excel", st.session_state["xlsx"], file_name=nama_file,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")


if __name__ == "__main__":
    main()
