"""
step4_build_dataset.py
=======================
STEP KEEMPAT — gabungkan 4 gambar spektrogram jadi satu tensor,
parse label emosi, dan bagi dataset jadi train/val/test.

Setiap sampel RAVDESS menghasilkan 4 gambar grayscale (mel, mfcc, cqt, chroma)
di ENHANCED_DIR. Step ini menggabungkan keempat gambar tersebut menjadi
satu array 128×128×4. Bayangkan seperti gambar RGB (3 channel warna),
tapi kita punya 4 channel — masing-masing berisi representasi spektrum yang
berbeda dari audio yang sama.

Kenapa digabung jadi tensor, bukan 4 gambar terpisah?
Supaya model bisa belajar hubungan antar channel secara bersamaan.
Kalau diberikan satu per satu, model tidak pernah "lihat" keempatnya
sekaligus dan tidak bisa belajar pola yang muncul saat mel + mfcc
bergerak bersamaan.

Output: FUSED_DIR/dataset.npz
  X_train, X_val, X_test  → float32, shape (N, 128, 128, 4), range [0, 1]
  y_train, y_val, y_test  → int, label kelas [0, 7]
"""

import os
import sys
import glob
import numpy as np
import cv2
from collections import defaultdict
from sklearn.model_selection import train_test_split
from tqdm import tqdm

sys.path.append('/content/drive/MyDrive/SER_Research')
from config import (
    ENHANCED_DIR, SPEC_DIR, FUSED_DIR,
    EMOTION_MAP, EMOTION_LABELS, N_CLASSES,
    IMG_SIZE, CHANNEL_ORDER,
    RANDOM_SEED, TEST_SIZE, VAL_SIZE
)


# ── PARSING LABEL ─────────────────────────────────────────────────────────────

def parse_ravdess_label(filename):
    """
    Ekstrak kode emosi dari nama file RAVDESS dan kembalikan indeks kelas.

    Format nama file RAVDESS:
      03-01-[EMOSI]-01-01-01-01_mel.png
              ↑ ini yang kita ambil, posisi index 2 setelah split '-'

    Kode emosi:
      '01' → neutral  → index 0
      '02' → calm     → index 1
      '03' → happy    → index 2
      '04' → sad      → index 3
      '05' → angry    → index 4
      '06' → fearful  → index 5
      '07' → disgust  → index 6
      '08' → surprised→ index 7

    Mengembalikan integer [0, 7] atau None kalau format tidak dikenali.
    """
    # Hapus suffix channel (_mel, _mfcc, dst) dan ekstensi
    base = os.path.basename(filename)
    # nama file mungkin: 03-01-05-01-02-01-12_mel.png
    # kita buang bagian setelah '_' terakhir
    name_no_ext = os.path.splitext(base)[0]   # hapus .png
    ravdess_part = name_no_ext.rsplit('_', 1)[0]  # hapus suffix _mel / _mfcc / dll

    parts = ravdess_part.split('-')
    if len(parts) < 3:
        return None

    emotion_code = parts[2]  # index 2 = kode emosi
    emotion_name = EMOTION_MAP.get(emotion_code)
    if emotion_name is None:
        return None

    return EMOTION_LABELS.index(emotion_name)


# ── FUSION 4 CHANNEL ──────────────────────────────────────────────────────────

def fuse_channels(file_dict):
    """
    Gabungkan 4 gambar grayscale menjadi satu tensor float32 128×128×4.

    file_dict: dict {channel_type: path_to_png}
               contoh: {'mel': '...mel.png', 'mfcc': '...mfcc.png', ...}

    Urutan channel mengikuti CHANNEL_ORDER dari config.py:
      index 0 = mel
      index 1 = mfcc
      index 2 = cqt
      index 3 = chroma

    Kenapa dibagi 255?
    MobileNetV2 pretrained ImageNet mengharapkan input dalam range [0, 1]
    (setelah normalisasi dengan preprocess_input, tapi kita skip itu
    dan langsung normalize ke [0, 1] supaya konsisten di semua channel).
    """
    channels = []
    for ch_type in CHANNEL_ORDER:
        path = file_dict.get(ch_type)
        if path is None or not os.path.exists(path):
            # Channel hilang: isi dengan nol (akan ada warning tapi tidak crash)
            channels.append(np.zeros(IMG_SIZE, dtype=np.float32))
            continue

        img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
        if img is None:
            channels.append(np.zeros(IMG_SIZE, dtype=np.float32))
            continue

        # Pastikan ukuran tepat (seharusnya sudah 128×128 dari step2)
        if img.shape[:2] != IMG_SIZE:
            img = cv2.resize(img, IMG_SIZE, interpolation=cv2.INTER_AREA)

        channels.append(img.astype(np.float32) / 255.0)

    # Stack: list of (128, 128) → (128, 128, 4)
    tensor = np.stack(channels, axis=-1)
    return tensor


# ── BANGUN DATASET PENUH ──────────────────────────────────────────────────────

def build_full_dataset(enhanced_dir=None):
    """
    Scan semua PNG di enhanced_dir, kelompokkan per sampel,
    fusion 4 channel, kembalikan X dan y.

    enhanced_dir: folder yang berisi PNG. Default ke ENHANCED_DIR.
                  Bisa di-override ke SPEC_DIR untuk kondisi B1/B3
                  (tanpa DIP enhancement).

    Cara kerja pengelompokan:
    Setiap sampel RAVDESS menghasilkan 4 file:
      {basename}_mel.png, {basename}_mfcc.png, dll.
    Kita kelompokkan berdasarkan {basename} (tanpa suffix channel),
    lalu untuk setiap kelompok yang punya semua 4 channel, kita fuse.

    Return: X (N, 128, 128, 4) float32, y (N,) int
    """
    if enhanced_dir is None:
        enhanced_dir = ENHANCED_DIR

    all_png = glob.glob(f"{enhanced_dir}/*.png")
    if not all_png:
        raise FileNotFoundError(
            f"Tidak ada PNG di {enhanced_dir}. "
            "Pastikan step2 dan step3 sudah dijalankan."
        )

    # Kelompokkan file per sampel RAVDESS
    sample_files = defaultdict(dict)
    for png_path in all_png:
        name_no_ext = os.path.splitext(os.path.basename(png_path))[0]
        # Pisahkan: '03-01-05-01-02-01-12_mel' → basename='03-01-05-01-02-01-12', ch='mel'
        parts = name_no_ext.rsplit('_', 1)
        if len(parts) != 2:
            continue
        basename, ch_type = parts
        if ch_type in CHANNEL_ORDER:
            sample_files[basename][ch_type] = png_path

    # Buat tensor dan label untuk setiap sampel yang lengkap (semua 4 channel ada)
    X_list = []
    y_list = []
    skipped = 0

    for basename, file_dict in tqdm(sample_files.items(), desc="Fusing channels", unit="sampel"):
        # Cek semua 4 channel ada
        if not all(ch in file_dict for ch in CHANNEL_ORDER):
            skipped += 1
            continue

        # Parse label dari nama file
        label = parse_ravdess_label(f"{basename}_mel.png")
        if label is None:
            skipped += 1
            continue

        tensor = fuse_channels(file_dict)
        X_list.append(tensor)
        y_list.append(label)

    if skipped > 0:
        print(f"Catatan: {skipped} sampel dilewati (channel tidak lengkap atau label tidak dikenali)")

    X = np.array(X_list, dtype=np.float32)   # (N, 128, 128, 4)
    y = np.array(y_list, dtype=np.int32)      # (N,)
    print(f"Dataset: {len(X)} sampel, shape={X.shape}, label unik={np.unique(y)}")
    return X, y


# ── SPLIT DATASET ─────────────────────────────────────────────────────────────

def split_dataset(X, y):
    """
    Bagi X dan y menjadi train / val / test dengan stratified sampling.

    Stratified artinya proporsi kelas dipertahankan di setiap split.
    Penting untuk dataset yang tidak seimbang — walaupun RAVDESS
    relatif seimbang, stratified tetap lebih aman.

    Rasio: 70% train, 15% val, 15% test
    Seed 42 untuk reproducibility.

    Kenapa split dilakukan sekali di sini dan disimpan ke disk?
    Supaya semua 4 kondisi ablasi (B1–B4) dievaluasi pada test set
    yang persis sama. Kalau split dilakukan ulang di step5 dengan
    seed yang sama, hasilnya seharusnya sama — tapi lebih rapi
    kalau di-hardcode di sini.
    """
    # Dua kali split: pertama pisahkan test, lalu sisanya dibagi train+val
    X_temp, X_test, y_temp, y_test = train_test_split(
        X, y,
        test_size=TEST_SIZE,
        random_state=RANDOM_SEED,
        stratify=y
    )

    # Val adalah TEST_SIZE dari sisa (setelah test diambil)
    # Misalnya total=1440, test=216, sisa=1224, val=216 → val_ratio = 216/1224 ≈ 0.1765
    val_ratio = VAL_SIZE / (1.0 - TEST_SIZE)
    X_train, X_val, y_train, y_val = train_test_split(
        X_temp, y_temp,
        test_size=val_ratio,
        random_state=RANDOM_SEED,
        stratify=y_temp
    )

    print(f"\nSplit dataset:")
    print(f"  Train : {len(X_train)} sampel ({len(X_train)/len(X)*100:.1f}%)")
    print(f"  Val   : {len(X_val)} sampel ({len(X_val)/len(X)*100:.1f}%)")
    print(f"  Test  : {len(X_test)} sampel ({len(X_test)/len(X)*100:.1f}%)")

    return X_train, y_train, X_val, y_val, X_test, y_test


# ── MAIN ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("STEP 4: Bangun Dataset (Fusion 4-Channel + Split)")
    print("=" * 50)

    os.makedirs(FUSED_DIR, exist_ok=True)

    # Cek apakah sudah ada dataset.npz
    out_path = os.path.join(FUSED_DIR, "dataset.npz")
    if os.path.exists(out_path):
        print(f"File {out_path} sudah ada.")
        print("Hapus manual kalau mau rebuild dari awal.")
        print("Load existing dataset untuk verifikasi...")
        data = np.load(out_path)
        print(f"  X_train: {data['X_train'].shape}")
        print(f"  X_val  : {data['X_val'].shape}")
        print(f"  X_test : {data['X_test'].shape}")
        sys.exit(0)

    # Bangun dari ENHANCED_DIR (dengan DIP — untuk kondisi B2 dan B4)
    print("Membangun dataset dari gambar yang sudah di-enhance (CLAHE + UM)...")
    X, y = build_full_dataset(enhanced_dir=ENHANCED_DIR)

    # Split
    X_train, y_train, X_val, y_val, X_test, y_test = split_dataset(X, y)

    # Simpan sebagai .npz (compressed numpy)
    # Kenapa .npz bukan .h5?
    # .npz lebih portable, tidak perlu h5py, dan scikit-learn juga pakai format ini.
    print(f"\nMenyimpan dataset ke {out_path}...")
    np.savez_compressed(
        out_path,
        X_train=X_train, y_train=y_train,
        X_val=X_val,     y_val=y_val,
        X_test=X_test,   y_test=y_test
    )

    file_size_mb = os.path.getsize(out_path) / 1024 / 1024
    print(f"Tersimpan ({file_size_mb:.1f} MB)")
    print("\nStep 4 selesai. Lanjut ke step5_train.py")
