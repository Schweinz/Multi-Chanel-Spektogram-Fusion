# ============================================================
# step3_dip_enhancement.py — DIP Enhancement pada Spektrogram
# ============================================================
# Jalankan setelah step2_extract_spectrograms.py
#
# Pipeline DIP 2 Tahap (novelty utama penelitian):
#   Tahap 1: CLAHE  → Buriboev et al., Sensors 2024
#            (Contrast Limited Adaptive Histogram Equalization)
#            Meningkatkan kontras lokal secara adaptif per tile
#   Tahap 2: Unsharp Masking → Tan et al., IJACSA 2022
#            Mempertegas detail halus spektrogram
#
# Mengapa DIP pada level gambar, bukan audio?
#   - Memungkinkan penggunaan teknik computer vision yang mapan
#   - Meningkatkan discriminability visual antar kelas emosi
#   - Sesuai dengan pipeline CNN-based image classification
# ============================================================

import os
import sys
import glob
import numpy as np
import cv2
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from tqdm import tqdm

sys.path.append('/content/drive/MyDrive/SER_Research')
from config import *


# ── TAHAP 1: CLAHE ───────────────────────────────────────────

def apply_clahe(img):
    """
    CLAHE: Contrast Limited Adaptive Histogram Equalization
    Referensi: Buriboev et al., Sensors, DOI:10.3390/s24237703, 2024

    Cara kerja:
    1. Bagi gambar menjadi tile-tile kecil (tileGridSize=8×8)
    2. Terapkan histogram equalization secara independen per tile
    3. Batasi amplifikasi dengan clipLimit=2.0 untuk cegah over-amplification
    4. Interpolasi bilinear antar tile untuk transisi halus

    Parameter:
        clipLimit (float): Batas amplifikasi kontras. Makin kecil = makin lembut.
        tileGridSize (tuple): Ukuran grid tile. (8,8) = 8×8 = 64 tile.
    """
    clahe = cv2.createCLAHE(
        clipLimit=CLAHE_CLIP_LIMIT,
        tileGridSize=CLAHE_TILE_GRID
    )
    return clahe.apply(img)


# ── TAHAP 2: UNSHARP MASKING ─────────────────────────────────

def apply_unsharp_masking(img):
    """
    Unsharp Masking untuk Image Detail Enhancement
    Referensi: Tan et al., IJACSA Vol.13 No.8, 2022

    Cara kerja:
    1. Buat versi blur dari gambar (low-pass filter)
    2. Hitung mask = original - blur  (detail layer)
    3. Tambahkan mask yang diperkuat ke gambar asli:
       sharpened = original + strength × mask
       = original + strength × (original - blur)
       = (1 + strength) × original - strength × blur

    Implementasi menggunakan cv2.addWeighted:
       result = alpha × original + beta × blur + gamma
       di mana alpha = 1+strength, beta = -strength

    Parameter:
        sigma (float): Standar deviasi Gaussian blur. Lebih besar = lebih blur.
        strength (float): Kekuatan penajaman. 1.5 = penajaman moderat.
    """
    # Buat blur image
    blurred = cv2.GaussianBlur(img, (0, 0), UM_SIGMA)

    # Unsharp masking: sharpened = (1+strength)*original - strength*blurred
    sharpened = cv2.addWeighted(
        img,      1.0 + UM_STRENGTH,   # original dengan bobot tinggi
        blurred,  -UM_STRENGTH,         # blur dengan bobot negatif
        0                               # gamma offset
    )
    return np.clip(sharpened, 0, 255).astype(np.uint8)


# ── PIPELINE DIP LENGKAP ─────────────────────────────────────

def apply_dip_pipeline(img):
    """
    Pipeline DIP 2 tahap penuh: CLAHE → Unsharp Masking
    Input : gambar grayscale uint8 [0,255]
    Output: gambar yang telah di-enhance uint8 [0,255]
    """
    # Tahap 1: CLAHE
    after_clahe = apply_clahe(img)

    # Tahap 2: Unsharp Masking
    after_um = apply_unsharp_masking(after_clahe)

    return after_um


def enhance_one_file(src_path, out_dir, spec_type):
    """
    Enhance satu file gambar spektrogram.
    Simpan hasil ke {out_dir}/{base_name}_{spec_type}.png
    """
    base_name = os.path.splitext(os.path.basename(src_path))[0]
    # base_name sudah berformat: {audio_name}_{spec_type}
    out_path = os.path.join(out_dir, f"{base_name}.png")

    if os.path.exists(out_path):
        return True

    try:
        img = cv2.imread(src_path, cv2.IMREAD_GRAYSCALE)
        if img is None:
            return False
        enhanced = apply_dip_pipeline(img)
        cv2.imwrite(out_path, enhanced)
        return True
    except Exception as e:
        print(f"\n❌ Error pada {base_name}: {e}")
        return False


# ── VISUALISASI PERBANDINGAN ──────────────────────────────────

def visualize_dip_comparison(wav_basename=None):
    """
    Tampilkan perbandingan Before/After DIP untuk semua 4 channel.
    """
    # Cari satu sampel
    if wav_basename is None:
        mel_files = glob.glob(f"{SPEC_DIR}/*_mel.png")
        if not mel_files:
            print("❌ Tidak ada file spektrogram untuk divisualisasi.")
            return
        wav_basename = os.path.basename(mel_files[0]).replace('_mel.png', '')

    print(f"📊 Visualisasi DIP comparison: {wav_basename}")

    fig, axes = plt.subplots(4, 3, figsize=(15, 20))
    fig.suptitle('DIP Enhancement: Before → After CLAHE → After CLAHE+UM',
                 fontsize=13, fontweight='bold')

    for row_idx, spec_type in enumerate(CHANNEL_ORDER):
        orig_path = os.path.join(SPEC_DIR, f"{wav_basename}_{spec_type}.png")
        enh_path  = os.path.join(ENHANCED_DIR, f"{wav_basename}_{spec_type}.png")

        orig = cv2.imread(orig_path, cv2.IMREAD_GRAYSCALE)
        if orig is None:
            continue

        after_clahe = apply_clahe(orig)
        after_um    = apply_unsharp_masking(after_clahe)

        for col_idx, (title, img) in enumerate([
            (f"{spec_type.upper()}\nOriginal", orig),
            (f"{spec_type.upper()}\nAfter CLAHE", after_clahe),
            (f"{spec_type.upper()}\nAfter CLAHE+UM", after_um),
        ]):
            axes[row_idx, col_idx].imshow(img, cmap='viridis',
                                           aspect='auto', origin='lower')
            axes[row_idx, col_idx].set_title(title, fontsize=10)
            axes[row_idx, col_idx].axis('off')

    plt.tight_layout()
    save_path = os.path.join(RESULT_DIR, "dip_comparison.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.show()
    print(f"✅ Gambar perbandingan disimpan ke {save_path}")


# ── MAIN ─────────────────────────────────────────────────────

if __name__ == "__main__":
    print("🚀 STEP 3: DIP Enhancement pada Spektrogram")
    print("="*50)
    print(f"  Metode 1 : CLAHE (clipLimit={CLAHE_CLIP_LIMIT}, "
          f"tileGrid={CLAHE_TILE_GRID})")
    print(f"  Metode 2 : Unsharp Masking (sigma={UM_SIGMA}, "
          f"strength={UM_STRENGTH})")
    print("="*50)

    os.makedirs(ENHANCED_DIR, exist_ok=True)
    os.makedirs(RESULT_DIR, exist_ok=True)

    # Kumpulkan semua gambar spektrogram
    spec_files = glob.glob(f"{SPEC_DIR}/*.png")
    if len(spec_files) == 0:
        print("❌ Tidak ada gambar spektrogram! Jalankan step2 dulu.")
        sys.exit(1)

    print(f"🖼️  Total gambar spektrogram: {len(spec_files)}")
    print(f"📁 Output ke: {ENHANCED_DIR}")

    success = 0
    fail = 0
    for spec_path in tqdm(spec_files, desc="DIP Enhancement"):
        spec_type = None
        for t in CHANNEL_ORDER:
            if spec_path.endswith(f"_{t}.png"):
                spec_type = t
                break
        ok = enhance_one_file(spec_path, ENHANCED_DIR, spec_type)
        if ok:
            success += 1
        else:
            fail += 1

    total_enhanced = len(glob.glob(f"{ENHANCED_DIR}/*.png"))
    print(f"\n{'='*50}")
    print(f"✅ Berhasil     : {success} gambar")
    print(f"❌ Gagal        : {fail} gambar")
    print(f"🖼️  Total tersimpan: {total_enhanced}")

    # Visualisasi perbandingan
    print("\n📊 Membuat visualisasi perbandingan DIP...")
    visualize_dip_comparison()

    print("\n✅ Step 3 selesai! Lanjut ke step4_build_dataset.py")
