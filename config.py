# ============================================================
# config.py — Konfigurasi Sentral Seluruh Pipeline SER
# Ubah nilai di sini untuk menyesuaikan dengan setup kamu
# ============================================================

import os

# ── PATH ─────────────────────────────────────────────────────
BASE_DIR        = "/content/drive/MyDrive/SER_Research"
RAW_DATA_DIR    = os.path.join(BASE_DIR, "data/raw")
SPEC_DIR        = os.path.join(BASE_DIR, "data/spectrograms")
ENHANCED_DIR    = os.path.join(BASE_DIR, "data/enhanced")
FUSED_DIR       = os.path.join(BASE_DIR, "data/fused")
MODEL_DIR       = os.path.join(BASE_DIR, "models")
RESULT_DIR      = os.path.join(BASE_DIR, "results")
GRADCAM_DIR     = os.path.join(BASE_DIR, "gradcam")

# ── AUDIO ────────────────────────────────────────────────────
SAMPLE_RATE     = 22050

# ── SPECTROGRAM ──────────────────────────────────────────────
IMG_SIZE        = (128, 128)
N_MELS          = 128
N_FFT           = 2048
HOP_LENGTH      = 512
N_MFCC          = 40
N_BINS          = 84
BINS_PER_OCTAVE = 12
N_CHROMA        = 12

# ── DIP ENHANCEMENT ──────────────────────────────────────────
# CLAHE — Buriboev et al., Sensors 2024
CLAHE_CLIP_LIMIT = 2.0
CLAHE_TILE_GRID  = (8, 8)
# Unsharp Masking — Tan et al., IJACSA 2022
UM_SIGMA         = 1.0
UM_STRENGTH      = 1.5
BLEND_ORIGINAL   = 0.7
BLEND_EDGE       = 0.3

# ── MULTI-CHANNEL FUSION ─────────────────────────────────────
N_CHANNELS      = 4
CHANNEL_ORDER   = ['mel', 'mfcc', 'cqt', 'chroma']

# ── DATASET SPLIT ────────────────────────────────────────────
RANDOM_SEED     = 42
TEST_SIZE       = 0.15
VAL_SIZE        = 0.15

# ── TRAINING ─────────────────────────────────────────────────
BATCH_SIZE      = 32
EPOCHS_FROZEN   = 10
EPOCHS_UNFROZEN = 100
LEARNING_RATE   = 0.0001
DROPOUT_RATE    = 0.3

# ── LABEL EMOSI RAVDESS ──────────────────────────────────────
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
EMOTION_LABELS = list(EMOTION_MAP.values())
N_CLASSES      = len(EMOTION_LABELS)

def create_dirs():
    """Buat semua folder yang diperlukan jika belum ada."""
    dirs = [RAW_DATA_DIR, SPEC_DIR, ENHANCED_DIR,
            FUSED_DIR, MODEL_DIR, RESULT_DIR, GRADCAM_DIR]
    for d in dirs:
        os.makedirs(d, exist_ok=True)
    print("✅ Semua folder berhasil dibuat.")

if __name__ == "__main__":
    create_dirs()
