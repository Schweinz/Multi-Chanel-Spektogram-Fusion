"""
step3_dip_enhancement.py
=========================
STEP KETIGA — perkuat gambar spektrogram dengan teknik Digital Image Processing.

INI ADALAH INTI DARI KONTRIBUSI PENELITIAN INI.

Hipotesis: spektrogram adalah gambar grayscale. Teknik DIP yang selama ini
dipakai di medical imaging (CT scan, MRI, X-ray) seharusnya bisa meningkatkan
fitur yang relevan di gambar spektrogram juga — tepi frekuensi, transisi
temporal, batas pola harmonik.

Pipeline DIP yang dipakai (urutan penting!):
  1. CLAHE  — perbaiki kontras lokal, angkat detail yang tertutup noise
  2. Unsharp Masking — tajamkan tepi dan detail halus setelah CLAHE

Kenapa CLAHE dulu, bukan Unsharp Masking dulu?
CLAHE membuat histogram lebih merata → detail tersembunyi jadi visible.
Unsharp Masking bekerja lebih efektif setelah detail sudah ter-enhance
oleh CLAHE. Urutan terbalik akan menajamkan noise sebelum CLAHE sempat
meredamnya.

Input : SPEC_DIR  — gambar grayscale sebelum enhancement (dari step2)
Output: ENHANCED_DIR — gambar grayscale setelah enhancement
"""

import os
import sys
import glob
import numpy as np
import cv2
from tqdm import tqdm
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

sys.path.append('/content/drive/MyDrive/SER_Research')
from config import (
    SPEC_DIR, ENHANCED_DIR,
    CLAHE_CLIP_LIMIT, CLAHE_TILE_GRID,
    UM_SIGMA, UM_STRENGTH
)


# ── CLAHE ─────────────────────────────────────────────────────────────────────

def apply_clahe(img):
    """
    CLAHE = Contrast Limited Adaptive Histogram Equalization.

    Berbeda dari histogram equalization biasa yang bekerja pada gambar
    keseluruhan, CLAHE membagi gambar menjadi tile-tile kecil dan
    menyamakan histogram di setiap tile secara independen. Ini mencegah
    area gelap yang detail-nya tidak kelihatan tetap gelap hanya karena
    area terang mendominasi histogram global.

    Parameter yang dipakai (dari Buriboev et al., Sensors 2024):
      clipLimit   = 2.0  → batas amplifikasi kontras per tile.
                          Tanpa batas ini, area yang sudah kontras tinggi
                          akan over-enhanced dan noise ikut diperkuat.
                          Nilai 2.0 adalah tradeoff yang umum dipakai.
      tileGridSize = (8,8) → gambar 128×128 dibagi 64 tile ukuran 16×16 px.
                          Tile lebih kecil → adaptasi lebih lokal tapi
                          lebih rentan noise. 8×8 cukup lokal untuk
                          menangkap pola frekuensi lokal di spektrogram.

    Input : img — array 2D grayscale uint8 [0, 255]
    Output: img yang sudah di-enhance, ukuran dan tipe sama
    """
    clahe = cv2.createCLAHE(
        clipLimit=CLAHE_CLIP_LIMIT,
        tileGridSize=CLAHE_TILE_GRID
    )
    return clahe.apply(img)


# ── UNSHARP MASKING ───────────────────────────────────────────────────────────

def apply_unsharp_masking(img, sigma=None, strength=None):
    """
    Unsharp Masking — teknik penajaman gambar klasik dari dunia fotografi analog.

    Cara kerjanya:
      1. Buat versi blur dari gambar asli (= "mask")
      2. Kurangi blur dari asli → yang tersisa adalah tepi dan detail halus
      3. Tambahkan detail itu kembali ke gambar asli dengan bobot tertentu

    Secara matematis:
      sharpened = (1 + strength) * original - strength * blurred
    yang sama dengan:
      sharpened = original + strength * (original - blurred)

    Kode OpenCV yang setara:
      cv2.addWeighted(img, 1 + strength, blurred, -strength, 0)

    Parameter (dari Tan et al., IJACSA 2022):
      sigma    = 1.0  → radius blur kecil, hanya mempengaruhi detail
                        sangat halus (sub-piksel). Sigma besar akan
                        menajamkan pola berskala lebih besar tapi
                        berisiko membuat artefak halo.
      strength = 1.5  → penajaman moderat. Range 1.0–2.0 wajar untuk
                        gambar spektrogram. Di bawah 1.0 hampir tidak
                        terlihat, di atas 2.0 mulai terlihat artefak.

    Input : img — array 2D grayscale uint8, sebaiknya sudah di-CLAHE
    Output: img yang sudah ditajamkan, uint8 [0, 255]
    """
    if sigma is None:
        sigma = UM_SIGMA
    if strength is None:
        strength = UM_STRENGTH

    # ksize=0 artinya OpenCV hitung otomatis dari sigma: ksize = ceil(2*3*sigma+1) | 1
    blurred = cv2.GaussianBlur(img, (0, 0), sigmaX=sigma, sigmaY=sigma)

    # addWeighted menghitung: dst = alpha * src1 + beta * src2 + gamma
    # kita set alpha = 1+strength, beta = -strength, gamma = 0
    sharpened = cv2.addWeighted(img, 1 + strength, blurred, -strength, 0)

    return sharpened


# ── PIPELINE DIP LENGKAP ──────────────────────────────────────────────────────

def apply_dip_pipeline(img):
    """
    Jalankan pipeline DIP penuh: CLAHE → Unsharp Masking.

    Fungsi ini adalah yang dipanggil untuk setiap gambar spektrogram.
    Urutan CLAHE → UM adalah keputusan desain yang sudah dibahas
    di docstring modul atas.

    Input : img — grayscale uint8, ukuran sembarang
    Output: img setelah CLAHE + UM, ukuran dan tipe sama
    """
    img_clahe = apply_clahe(img)
    img_enhanced = apply_unsharp_masking(img_clahe)
    return img_enhanced


# ── PROSES SATU FILE ──────────────────────────────────────────────────────────

def enhance_one_file(png_path, output_dir):
    """
    Baca satu PNG grayscale, jalankan pipeline DIP, simpan hasilnya.

    Kenapa perlu cek grayscale?
    cv2.imread() secara default membaca sebagai BGR (3 channel).
    Kita perlu memastikan gambar dibaca sebagai grayscale single-channel
    sebelum dimasukkan ke CLAHE yang hanya menerima single-channel.

    Mengembalikan True kalau berhasil, False kalau ada error.
    """
    try:
        basename = os.path.basename(png_path)
        out_path = os.path.join(output_dir, basename)

        # Skip kalau sudah ada
        if os.path.exists(out_path):
            return True

        # Baca sebagai grayscale langsung
        img = cv2.imread(png_path, cv2.IMREAD_GRAYSCALE)
        if img is None:
            print(f"\nTidak bisa membaca: {png_path}")
            return False

        # Jalankan pipeline DIP
        img_enhanced = apply_dip_pipeline(img)

        # Simpan
        cv2.imwrite(out_path, img_enhanced)
        return True

    except Exception as e:
        print(f"\nError memproses {os.path.basename(png_path)}: {e}")
        return False


# ── VISUALISASI PERBANDINGAN ───────────────────────────────────────────────────

def visualize_dip_comparison(basename_prefix, n_channels=4):
    """
    Tampilkan perbandingan before/after CLAHE/after CLAHE+UM untuk semua channel.

    basename_prefix: nama file tanpa suffix tipe dan ekstensi
                     contoh: '03-01-05-01-02-01-12'

    Layout: n_channels baris × 3 kolom
      kolom 1: asli (dari SPEC_DIR)
      kolom 2: setelah CLAHE saja
      kolom 3: setelah CLAHE + Unsharp Masking (= ENHANCED_DIR)

    Ini membantu verifikasi visual bahwa enhancement bekerja dengan benar.
    """
    channel_types = ['mel', 'mfcc', 'cqt', 'chroma']
    fig, axes = plt.subplots(n_channels, 3, figsize=(12, 4 * n_channels))
    fig.suptitle(f"DIP Comparison: {basename_prefix}", fontsize=13)

    col_labels = ['Asli', 'Setelah CLAHE', 'Setelah CLAHE + UM']
    for ax, label in zip(axes[0], col_labels):
        ax.set_title(label, fontsize=11, fontweight='bold')

    for row, ch_type in enumerate(channel_types[:n_channels]):
        fname = f"{basename_prefix}_{ch_type}.png"
        orig_path = os.path.join(SPEC_DIR, fname)

        orig = cv2.imread(orig_path, cv2.IMREAD_GRAYSCALE)
        if orig is None:
            for col in range(3):
                axes[row, col].text(0.5, 0.5, f"File tidak ada:\n{fname}",
                                    ha='center', va='center', transform=axes[row, col].transAxes)
            continue

        after_clahe = apply_clahe(orig)
        after_both = apply_unsharp_masking(after_clahe)

        for col, (img, title) in enumerate(zip(
            [orig, after_clahe, after_both],
            col_labels
        )):
            axes[row, col].imshow(img, cmap='magma', aspect='auto')
            axes[row, col].set_ylabel(ch_type.upper(), fontsize=9)
            axes[row, col].set_xticks([])
            axes[row, col].set_yticks([])

    plt.tight_layout()
    plt.show()


# ── MAIN ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("STEP 3: DIP Enhancement (CLAHE + Unsharp Masking)")
    print("=" * 50)
    print(f"  CLAHE  : clipLimit={CLAHE_CLIP_LIMIT}, tileGrid={CLAHE_TILE_GRID}")
    print(f"  Unsharp: sigma={UM_SIGMA}, strength={UM_STRENGTH}")
    print()

    os.makedirs(ENHANCED_DIR, exist_ok=True)

    # Temukan semua PNG di SPEC_DIR
    png_files = sorted(glob.glob(f"{SPEC_DIR}/*.png"))
    if not png_files:
        print("Tidak ada PNG di SPEC_DIR. Jalankan step2 dulu.")
        sys.exit(1)

    print(f"Ditemukan {len(png_files)} gambar untuk di-enhance")

    # Visualisasi perbandingan untuk satu sampel
    sample_file = png_files[0]
    sample_prefix = '_'.join(os.path.basename(sample_file).split('_')[:-1])
    print(f"\nVisualisasi DIP untuk sampel: {sample_prefix}")
    visualize_dip_comparison(sample_prefix)

    # Proses semua file
    success = 0
    fail = 0
    for png_path in tqdm(png_files, desc="DIP Enhancement", unit="file"):
        ok = enhance_one_file(png_path, ENHANCED_DIR)
        if ok:
            success += 1
        else:
            fail += 1

    print(f"\nSelesai: {success} berhasil, {fail} gagal")

    # Verifikasi
    out_count = len(glob.glob(f"{ENHANCED_DIR}/*.png"))
    print(f"File di ENHANCED_DIR: {out_count} (expected: {len(png_files)})")

    if out_count >= len(png_files) * 0.99:
        print("Step 3 selesai. Lanjut ke step4_build_dataset.py")
    else:
        print("Ada file yang hilang — periksa error di atas.")
