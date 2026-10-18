# ============================================================
# step4_build_dataset.py — Fusi 4-Channel & Pembuatan Dataset
# ============================================================
# Jalankan setelah step3_dip_enhancement.py
#
# Script ini akan:
#   1. Gabungkan 4 gambar spektrogram (Mel+MFCC+CQT+Chroma)
#      menjadi satu tensor 128×128×4 per sampel
#   2. Ekstrak label dari nama file RAVDESS
#   3. Split dataset: 70% train / 15% val / 15% test
#      (stratified per kelas emosi)
#   4. Simpan sebagai file .npz di FUSED_DIR
#   5. Visualisasi distribusi kelas
# ============================================================

import os
import sys
import glob
import numpy as np
import cv2
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from collections import Counter

sys.path.append('/content/drive/MyDrive/SER_Research')
from config import *


# ── EKSTRAKSI LABEL DARI NAMA FILE ───────────────────────────

def parse_ravdess_label(filename):
    """
    Parse label emosi dari nama file RAVDESS.
    Format: 03-01-[emotion]-[intensity]-[statement]-[repetition]-[actor].wav
    Index 2 (kolom ke-3) = kode emosi ('01'–'08')

    Contoh: 03-01-05-01-01-01-01 → '05' → 'angry'
    """
    base = os.path.splitext(os.path.basename(filename))[0]
    # Hapus suffix channel jika ada (_mel, _mfcc, _cqt, _chroma)
    for suffix in ['_mel', '_mfcc', '_cqt', '_chroma']:
        if base.endswith(suffix):
            base = base[:-len(suffix)]
            break
    parts = base.split('-')
    if len(parts) < 3:
        return None
    emo_code = parts[2]
    return EMOTION_MAP.get(emo_code, None)


def label_to_idx(label):
    """Konversi nama emosi ke indeks integer (0-7)."""
    return EMOTION_LABELS.index(label)


# ── FUSI 4-CHANNEL ───────────────────────────────────────────

def fuse_channels(base_name, enhanced_dir):
    """
    Gabungkan 4 gambar grayscale (mel, mfcc, cqt, chroma)
    menjadi satu array numpy 128×128×4.

    Input : base_name (tanpa ekstensi channel, misal '03-01-05-01-01-01-01')
    Output: array float32 [0,1] berukuran (128, 128, 4)
            atau None jika ada channel yang hilang
    """
    channels = []
    for ch_type in CHANNEL_ORDER:
        img_path = os.path.join(enhanced_dir, f"{base_name}_{ch_type}.png")
        img = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)
        if img is None:
            return None
        # Resize ulang untuk memastikan konsistensi
        img = cv2.resize(img, IMG_SIZE, interpolation=cv2.INTER_AREA)
        channels.append(img)

    # Stack → (128, 128, 4), normalisasi ke [0,1]
    tensor = np.stack(channels, axis=-1).astype(np.float32) / 255.0
    return tensor


# ── BANGUN SELURUH DATASET ────────────────────────────────────

def build_full_dataset(enhanced_dir=ENHANCED_DIR):
    """
    Iterasi semua sampel di enhanced_dir, fusi per sampel,
    return X (N, 128, 128, 4) dan y (N,) integer.
    """
    # Temukan semua base_name unik dari file _mel.png
    mel_files = sorted(glob.glob(os.path.join(enhanced_dir, "*_mel.png")))
    if len(mel_files) == 0:
        raise FileNotFoundError(
            f"Tidak ada file *_mel.png di {enhanced_dir}. "
            "Pastikan step3 sudah selesai."
        )

    print(f"📂 Total sampel ditemukan: {len(mel_files)}")

    X_list, y_list, skipped = [], [], []
    for mel_path in mel_files:
        base_name = os.path.basename(mel_path).replace('_mel.png', '')
        label = parse_ravdess_label(base_name)
        if label is None:
            skipped.append(base_name)
            continue

        tensor = fuse_channels(base_name, enhanced_dir)
        if tensor is None:
            skipped.append(base_name)
            continue

        X_list.append(tensor)
        y_list.append(label_to_idx(label))

    X = np.array(X_list, dtype=np.float32)
    y = np.array(y_list, dtype=np.int32)

    print(f"✅ Berhasil diproses : {len(X)} sampel")
    if skipped:
        print(f"⚠️  Dilewati         : {len(skipped)} sampel")
    print(f"📐 Shape X           : {X.shape}")  # (N, 128, 128, 4)
    print(f"📐 Shape y           : {y.shape}")  # (N,)
    return X, y


# ── SPLIT DATASET ─────────────────────────────────────────────

def split_dataset(X, y):
    """
    Stratified split: 70% train / 15% val / 15% test
    Menggunakan RANDOM_SEED untuk reproducibility.

    Stratified = menjaga distribusi kelas yang seimbang
    di setiap split, penting untuk dataset yang tidak rata.
    """
    # 1. Split train (70%) vs temp (30%)
    X_train, X_temp, y_train, y_temp = train_test_split(
        X, y,
        test_size=(TEST_SIZE + VAL_SIZE),
        stratify=y,
        random_state=RANDOM_SEED
    )

    # 2. Split temp → val (50% of 30% = 15%) dan test (50% = 15%)
    val_ratio = VAL_SIZE / (TEST_SIZE + VAL_SIZE)
    X_val, X_test, y_val, y_test = train_test_split(
        X_temp, y_temp,
        test_size=(1.0 - val_ratio),
        stratify=y_temp,
        random_state=RANDOM_SEED
    )

    print(f"\n📊 Hasil Split Dataset:")
    print(f"  Train : {len(X_train):>4} sampel ({len(X_train)/len(X)*100:.1f}%)")
    print(f"  Val   : {len(X_val):>4} sampel ({len(X_val)/len(X)*100:.1f}%)")
    print(f"  Test  : {len(X_test):>4} sampel ({len(X_test)/len(X)*100:.1f}%)")

    return X_train, X_val, X_test, y_train, y_val, y_test


# ── SIMPAN DATASET ────────────────────────────────────────────

def save_dataset(X_train, X_val, X_test, y_train, y_val, y_test):
    """
    Simpan semua split ke file .npz di FUSED_DIR.
    .npz = arsip numpy compressed, cepat di-load ulang.
    """
    os.makedirs(FUSED_DIR, exist_ok=True)

    out_path = os.path.join(FUSED_DIR, "dataset.npz")
    np.savez_compressed(
        out_path,
        X_train=X_train, y_train=y_train,
        X_val=X_val,   y_val=y_val,
        X_test=X_test,  y_test=y_test
    )
    size_mb = os.path.getsize(out_path) / 1e6
    print(f"\n💾 Dataset disimpan ke: {out_path}")
    print(f"   Ukuran file: {size_mb:.1f} MB")


def load_dataset():
    """
    Load dataset dari file .npz.
    Gunakan fungsi ini di step5_train.py dan step6_evaluate_gradcam.py
    """
    data_path = os.path.join(FUSED_DIR, "dataset.npz")
    data = np.load(data_path)
    return (
        data['X_train'], data['y_train'],
        data['X_val'],   data['y_val'],
        data['X_test'],  data['y_test']
    )


# ── VISUALISASI DISTRIBUSI ────────────────────────────────────

def visualize_distribution(y_train, y_val, y_test):
    """
    Plot distribusi kelas emosi di setiap split.
    Idealnya distribusi seragam (balanced) karena RAVDESS dirancang seimbang.
    """
    fig, axes = plt.subplots(1, 3, figsize=(18, 5))
    fig.suptitle('Distribusi Kelas Emosi per Split Dataset',
                 fontsize=14, fontweight='bold')

    for ax, (split_name, y_split) in zip(axes, [
        ('Train', y_train), ('Validation', y_val), ('Test', y_test)
    ]):
        counts = Counter(y_split)
        emotions = [EMOTION_LABELS[i] for i in range(N_CLASSES)]
        values   = [counts.get(i, 0) for i in range(N_CLASSES)]

        bars = ax.bar(emotions, values,
                      color=sns.color_palette('husl', N_CLASSES))
        ax.set_title(f'{split_name} (n={sum(values)})')
        ax.set_xlabel('Emosi')
        ax.set_ylabel('Jumlah Sampel')
        ax.tick_params(axis='x', rotation=45)
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                    str(val), ha='center', va='bottom', fontsize=8)

    plt.tight_layout()
    save_path = os.path.join(RESULT_DIR, "dataset_distribution.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.show()
    print(f"✅ Grafik distribusi disimpan ke {save_path}")


def visualize_sample_tensor(X_train, y_train, n_samples=3):
    """
    Tampilkan contoh tensor 4-channel dari beberapa sampel.
    """
    fig, axes = plt.subplots(n_samples, 4, figsize=(16, 4*n_samples))
    fig.suptitle('Contoh Tensor 4-Channel per Sampel',
                 fontsize=13, fontweight='bold')

    for row in range(n_samples):
        idx = row * (len(X_train) // n_samples)
        label = EMOTION_LABELS[y_train[idx]]
        for col, ch_name in enumerate(CHANNEL_ORDER):
            axes[row, col].imshow(X_train[idx, :, :, col],
                                  cmap='viridis', aspect='auto', origin='lower')
            axes[row, col].set_title(
                f'{ch_name.upper()}\n(label: {label})' if col == 0 else ch_name.upper()
            )
            axes[row, col].axis('off')

    plt.tight_layout()
    save_path = os.path.join(RESULT_DIR, "sample_tensor_4ch.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.show()
    print(f"✅ Gambar tensor tersimpan ke {save_path}")


# ── MAIN ─────────────────────────────────────────────────────

if __name__ == "__main__":
    print("🚀 STEP 4: Fusi 4-Channel & Pembuatan Dataset")
    print("="*50)

    os.makedirs(RESULT_DIR, exist_ok=True)

    # 1. Bangun dataset lengkap
    print("\n[1/4] Membangun dataset dari gambar enhanced...")
    X, y = build_full_dataset(ENHANCED_DIR)

    # 2. Verifikasi label
    print(f"\n[2/4] Distribusi kelas (sebelum split):")
    for i, emo in enumerate(EMOTION_LABELS):
        count = np.sum(y == i)
        print(f"  {i}: {emo:<12} → {count} sampel")

    # 3. Split
    print("\n[3/4] Membagi dataset (stratified split)...")
    X_train, X_val, X_test, y_train, y_val, y_test = split_dataset(X, y)

    # 4. Simpan
    print("\n[4/4] Menyimpan dataset...")
    save_dataset(X_train, X_val, X_test, y_train, y_val, y_test)

    # Visualisasi
    print("\n📊 Membuat visualisasi distribusi...")
    visualize_distribution(y_train, y_val, y_test)

    print("\n📊 Menampilkan contoh tensor 4-channel...")
    visualize_sample_tensor(X_train, y_train)

    print("\n" + "="*50)
    print("✅ Step 4 selesai!")
    print(f"   Dataset tersimpan di: {FUSED_DIR}/dataset.npz")
    print("   Lanjut ke step5_train.py")
