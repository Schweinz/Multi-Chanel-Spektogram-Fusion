# ============================================================
# step2_extract_spectrograms.py — Ekstraksi 4 Spektrogram
# ============================================================
# Jalankan setelah step1_prepare_data.py
# Script ini menghasilkan 4 jenis gambar spektrogram per file:
#   1. Mel Spectrogram   → representasi energi frekuensi mel
#   2. MFCC              → ciri vokal/timbral
#   3. CQT               → ciri harmonik/nada
#   4. Chroma STFT       → ciri tonal/nada
# Setiap gambar disimpan sebagai .png 128×128 grayscale
# ============================================================

import os
import sys
import glob
import numpy as np
import librosa
import librosa.display
import cv2
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend untuk server
import matplotlib.pyplot as plt
from tqdm import tqdm

sys.path.append('/content/drive/MyDrive/SER_Research')
from config import *

# ── FUNGSI EKSTRAKSI SPEKTROGRAM ────────────────────────────

def load_audio(filepath):
    """Load file audio dengan sample rate yang sudah ditentukan."""
    y, sr = librosa.load(filepath, sr=SAMPLE_RATE)
    # Padding/truncating ke panjang 3 detik agar konsisten
    target_len = SAMPLE_RATE * 3
    if len(y) < target_len:
        y = np.pad(y, (0, target_len - len(y)))
    else:
        y = y[:target_len]
    return y, sr


def extract_mel(y, sr):
    """
    Mel Spectrogram — representasi energi dalam skala mel.
    Baik untuk menangkap karakteristik energi emosi secara umum.
    """
    mel = librosa.feature.melspectrogram(
        y=y, sr=sr,
        n_mels=N_MELS,
        n_fft=N_FFT,
        hop_length=HOP_LENGTH
    )
    mel_db = librosa.power_to_db(mel, ref=np.max)
    return mel_db


def extract_mfcc(y, sr):
    """
    MFCC — Mel Frequency Cepstral Coefficients.
    Merepresentasikan ciri vokal dan timbral dari suara emosi.
    """
    mfcc = librosa.feature.mfcc(
        y=y, sr=sr,
        n_mfcc=N_MFCC,
        n_fft=N_FFT,
        hop_length=HOP_LENGTH
    )
    return mfcc


def extract_cqt(y, sr):
    """
    Constant-Q Transform (CQT) — representasi harmonik/nada.
    Resolusi frekuensi logaritmik, cocok untuk variasi pitch emosi.
    """
    cqt = librosa.cqt(
        y=y, sr=sr,
        n_bins=N_BINS,
        bins_per_octave=BINS_PER_OCTAVE,
        hop_length=HOP_LENGTH
    )
    cqt_db = librosa.amplitude_to_db(np.abs(cqt), ref=np.max)
    return cqt_db


def extract_chroma(y, sr):
    """
    Chroma STFT — representasi distribusi energi per pitch class.
    Menangkap informasi tonal (major/minor) yang berkaitan dengan emosi.
    """
    chroma = librosa.feature.chroma_stft(
        y=y, sr=sr,
        n_chroma=N_CHROMA,
        n_fft=N_FFT,
        hop_length=HOP_LENGTH
    )
    return chroma


def spectrogram_to_image(spec_data, img_size=IMG_SIZE):
    """
    Konversi array numpy spektrogram ke gambar grayscale 128×128.
    Langkah:
    1. Normalisasi ke rentang [0, 255]
    2. Resize ke IMG_SIZE menggunakan INTER_AREA (anti-aliasing)
    3. Konversi ke uint8
    """
    # Normalisasi ke [0, 1]
    spec_min = spec_data.min()
    spec_max = spec_data.max()
    if spec_max - spec_min > 1e-8:
        normalized = (spec_data - spec_min) / (spec_max - spec_min)
    else:
        normalized = np.zeros_like(spec_data)

    # Ke [0, 255] uint8
    img = (normalized * 255).astype(np.uint8)

    # Resize ke target size
    img_resized = cv2.resize(img, img_size, interpolation=cv2.INTER_AREA)
    return img_resized


def process_one_file(wav_path, out_dir):
    """
    Proses satu file .wav → simpan 4 gambar spektrogram.
    Nama file output: {base_name}_{type}.png
    """
    base_name = os.path.splitext(os.path.basename(wav_path))[0]

    # Skip jika semua 4 file sudah ada
    all_exist = all(
        os.path.exists(os.path.join(out_dir, f"{base_name}_{t}.png"))
        for t in CHANNEL_ORDER
    )
    if all_exist:
        return True

    try:
        y, sr = load_audio(wav_path)

        extractors = {
            'mel':    extract_mel(y, sr),
            'mfcc':   extract_mfcc(y, sr),
            'cqt':    extract_cqt(y, sr),
            'chroma': extract_chroma(y, sr),
        }

        for spec_type, spec_data in extractors.items():
            img = spectrogram_to_image(spec_data)
            out_path = os.path.join(out_dir, f"{base_name}_{spec_type}.png")
            cv2.imwrite(out_path, img)

        return True
    except Exception as e:
        print(f"\n❌ Error pada {base_name}: {e}")
        return False


# ── VISUALISASI SAMPEL ───────────────────────────────────────

def visualize_sample(wav_path=None):
    """
    Tampilkan 4 spektrogram dari satu file sample untuk verifikasi.
    """
    if wav_path is None:
        wav_files = glob.glob(f"{RAW_DATA_DIR}/**/*.wav", recursive=True)
        if not wav_files:
            wav_files = glob.glob(f"{RAW_DATA_DIR}/*.wav")
        if not wav_files:
            print("❌ Tidak ada file .wav untuk divisualisasi.")
            return
        wav_path = wav_files[0]

    print(f"📊 Visualisasi spektrogram dari: {os.path.basename(wav_path)}")
    y, sr = load_audio(wav_path)

    fig, axes = plt.subplots(2, 2, figsize=(12, 8))
    fig.suptitle('4-Channel Spectrogram Visualization', fontsize=14, fontweight='bold')

    specs = {
        'Mel Spectrogram':  extract_mel(y, sr),
        'MFCC':             extract_mfcc(y, sr),
        'CQT':              extract_cqt(y, sr),
        'Chroma STFT':      extract_chroma(y, sr),
    }

    for ax, (title, spec) in zip(axes.flat, specs.items()):
        img = spectrogram_to_image(spec)
        ax.imshow(img, cmap='viridis', aspect='auto', origin='lower')
        ax.set_title(title)
        ax.set_xlabel('Time Frames')
        ax.set_ylabel('Frequency Bins')
        plt.colorbar(ax.images[0], ax=ax)

    plt.tight_layout()
    save_path = os.path.join(RESULT_DIR, "sample_spectrograms.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.show()
    print(f"✅ Gambar disimpan ke {save_path}")


# ── MAIN ─────────────────────────────────────────────────────

if __name__ == "__main__":
    print("🚀 STEP 2: Ekstraksi Spektrogram")
    print("="*50)

    # Buat output dir
    os.makedirs(SPEC_DIR, exist_ok=True)
    os.makedirs(RESULT_DIR, exist_ok=True)

    # Kumpulkan semua file wav
    wav_files = glob.glob(f"{RAW_DATA_DIR}/**/*.wav", recursive=True)
    if not wav_files:
        wav_files = glob.glob(f"{RAW_DATA_DIR}/*.wav")

    if len(wav_files) == 0:
        print("❌ Tidak ada file .wav! Jalankan step1 dulu.")
        sys.exit(1)

    print(f"📂 Total file .wav: {len(wav_files)}")
    print(f"📁 Output ke: {SPEC_DIR}")
    print(f"🔄 Mengekstrak 4 spektrogram per file...")

    # Proses semua file
    success = 0
    fail = 0
    for wav_path in tqdm(wav_files, desc="Ekstraksi"):
        ok = process_one_file(wav_path, SPEC_DIR)
        if ok:
            success += 1
        else:
            fail += 1

    total_imgs = len(glob.glob(f"{SPEC_DIR}/*.png"))
    print(f"\n{'='*50}")
    print(f"✅ Berhasil : {success} file")
    print(f"❌ Gagal    : {fail} file")
    print(f"🖼️  Total gambar tersimpan: {total_imgs}")
    print(f"   (Expected: {len(wav_files) * 4} = {len(wav_files)} file × 4 jenis)")

    # Visualisasi 1 sampel
    print("\n📊 Membuat visualisasi sampel...")
    visualize_sample()

    print("\n✅ Step 2 selesai! Lanjut ke step3_dip_enhancement.py")
