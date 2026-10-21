"""
config.py
=========
Satu tempat untuk semua hyperparameter dan path yang dipakai
di seluruh pipeline. Kalau mau ubah sesuatu, cukup di sini.

Kenapa satu file konfigurasi?
Waktu eksperimen, saya sering lupa sudah ganti parameter di mana
setelah coba-coba beberapa kondisi. Dengan dipusatkan di sini,
perubahan satu nilai langsung terapply ke semua step tanpa harus
buka file satu persatu.
"""

import os


# ---------------------------------------------------------------------------
# PATH — sesuaikan BASE_DIR dengan lokasi Google Drive kamu
# ---------------------------------------------------------------------------

BASE_DIR     = "/content/drive/MyDrive/SER_Research"
RAW_DATA_DIR = os.path.join(BASE_DIR, "data/raw")          # file .wav asli
SPEC_DIR     = os.path.join(BASE_DIR, "data/spectrograms") # gambar sebelum DIP
ENHANCED_DIR = os.path.join(BASE_DIR, "data/enhanced")     # gambar setelah DIP
FUSED_DIR    = os.path.join(BASE_DIR, "data/fused")        # tensor 4-channel .npz
MODEL_DIR    = os.path.join(BASE_DIR, "models")            # bobot model .keras
RESULT_DIR   = os.path.join(BASE_DIR, "results")           # grafik & JSON metrik
GRADCAM_DIR  = os.path.join(BASE_DIR, "gradcam")           # gambar Grad-CAM


# ---------------------------------------------------------------------------
# AUDIO
# ---------------------------------------------------------------------------

# 22050 Hz adalah standard untuk audio speech analysis dengan librosa.
# Lebih tinggi tidak perlu karena frekuensi emosi relevan ada di bawah 8 kHz.
SAMPLE_RATE = 22050


# ---------------------------------------------------------------------------
# PARAMETER EKSTRAKSI SPEKTROGRAM
# ---------------------------------------------------------------------------

# Semua gambar disimpan 128x128 pixel.
# Cukup untuk menangkap pola visual, tidak terlalu besar untuk memory.
IMG_SIZE = (128, 128)

# Mel spectrogram: 128 mel bands, resolusi frekuensi yang baik untuk speech
N_MELS = 128

# FFT window 2048 sample (~93ms di 22050Hz), hop 512 sample (~23ms).
# Tradeoff standar antara resolusi waktu vs frekuensi.
N_FFT      = 2048
HOP_LENGTH = 512

# MFCC: ambil 40 koefisien. Lebih dari itu jarang membantu untuk speech.
N_MFCC = 40

# CQT: 84 bins, 12 bins per oktav = 7 oktav. Cukup untuk range vokal manusia.
N_BINS         = 84
BINS_PER_OCTAVE = 12

# Chroma: 12 pitch class (C, C#, D, ..., B) — satu oktav penuh.
N_CHROMA = 12


# ---------------------------------------------------------------------------
# PARAMETER DIP (Digital Image Processing)
# ---------------------------------------------------------------------------

# CLAHE — nilai dari Buriboev et al. (Sensors 2024).
# clipLimit=2.0: amplifikasi kontras dibatasi 2x. Kalau terlalu tinggi
# noise ikut diperkuat, kalau terlalu rendah efeknya tidak terasa.
# tileGridSize=(8,8): gambar 128x128 dibagi jadi 64 tile ukuran 16x16 px.
CLAHE_CLIP_LIMIT = 2.0
CLAHE_TILE_GRID  = (8, 8)

# Unsharp Masking — nilai dari Tan et al. (IJACSA 2022).
# sigma=1.0: blur ringan, hanya mempengaruhi detail sangat halus.
# strength=1.5: penajaman moderat. Coba 1.0–2.0 kalau mau eksperimen.
UM_SIGMA    = 1.0
UM_STRENGTH = 1.5

# Dua parameter ini tidak dipakai di kode utama tapi saya simpan di sini
# kalau nanti mau eksperimen dengan blending sebelum/sesudah enhancement.
BLEND_ORIGINAL = 0.7
BLEND_EDGE     = 0.3


# ---------------------------------------------------------------------------
# MULTI-CHANNEL FUSION
# ---------------------------------------------------------------------------

# Urutan channel penting karena konsisten dipakai di step2, 3, 4, 5, 6.
# Mel di index 0 karena dipakai kondisi B1 & B2 (single-channel baseline).
N_CHANNELS   = 4
CHANNEL_ORDER = ['mel', 'mfcc', 'cqt', 'chroma']


# ---------------------------------------------------------------------------
# SPLIT DATASET
# ---------------------------------------------------------------------------

# Seed 42 supaya hasil split bisa direproduksi.
# Satu split dibuat di step4, lalu dipakai sama persis di semua kondisi
# ablasi — jadi perbedaan performa bukan karena test set yang berbeda.
RANDOM_SEED = 42
TEST_SIZE   = 0.15   # 15% test   → ~216 sampel dari 1440
VAL_SIZE    = 0.15   # 15% val    → ~216 sampel
                     # 70% train  → ~1008 sampel


# ---------------------------------------------------------------------------
# HYPERPARAMETER TRAINING
# ---------------------------------------------------------------------------

BATCH_SIZE      = 32
EPOCHS_FROZEN   = 10   # Phase 1: backbone dibekukan, latih head saja
EPOCHS_UNFROZEN = 100  # Phase 2: fine-tuning semua layer (early stopping aktif)

# LR phase 1 cukup besar karena layer baru belum terlatih sama sekali.
# LR phase 2 = LR/10 = 1e-5, kecil supaya tidak merusak bobot ImageNet.
LEARNING_RATE = 1e-4
DROPOUT_RATE  = 0.3


# ---------------------------------------------------------------------------
# LABEL EMOSI RAVDESS
# ---------------------------------------------------------------------------

# Kode numerik dari nama file RAVDESS dipetakan ke nama emosi.
# Format file: 03-01-[kode]-[intensity]-[statement]-[repetition]-[actor].wav
# Kolom ke-3 (index 2) setelah split '-' adalah kode emosi.
EMOTION_MAP = {
    '01': 'neutral',
    '02': 'calm',
    '03': 'happy',
    '04': 'sad',
    '05': 'angry',
    '06': 'fearful',
    '07': 'disgust',
    '08': 'surprised'
}
EMOTION_LABELS = list(EMOTION_MAP.values())   # list terurut untuk indeks kelas
N_CLASSES      = len(EMOTION_LABELS)           # = 8


# ---------------------------------------------------------------------------
# UTILITY
# ---------------------------------------------------------------------------

def create_dirs():
    """Buat semua folder pipeline jika belum ada."""
    for d in [RAW_DATA_DIR, SPEC_DIR, ENHANCED_DIR,
              FUSED_DIR, MODEL_DIR, RESULT_DIR, GRADCAM_DIR]:
        os.makedirs(d, exist_ok=True)
    print("Semua folder pipeline sudah siap.")


if __name__ == "__main__":
    create_dirs()
