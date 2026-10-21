"""
step2_extract_spectrograms.py
==============================
STEP KEDUA — ubah file audio .wav jadi gambar spektrogram.

Setiap file .wav menghasilkan 4 gambar 128×128 piksel:
  • Mel Spectrogram  → representasi energi per frekuensi-mel vs waktu
  • MFCC             → koefisien cepstral, ringkasan kompak suara vokal
  • CQT              → Constant-Q Transform, log-frekuensi, mirip skala nada musik
  • Chroma           → distribusi energi per kelas nada (12 nada oktav)

Kenapa 4 jenis?
Masing-masing menangkap aspek berbeda dari sinyal bicara.
Mel bagus untuk tekstur vokal, MFCC untuk identitas pembicara,
CQT untuk kontur nada, Chroma untuk harmoni. Digabung jadi 4 channel
nanti di step4, semoga saling melengkapi.

Output folder: SPEC_DIR (dari config.py)
Nama file    : {basename}_{jenis}.png
              contoh: 03-01-05-01-02-01-12_mel.png
"""

import os
import sys
import glob
import numpy as np
import librosa
import librosa.display
import cv2
from tqdm import tqdm
import matplotlib
matplotlib.use('Agg')   # non-interactive backend — tidak perlu tampilan grafis
import matplotlib.pyplot as plt

sys.path.append('/content/drive/MyDrive/SER_Research')
from config import (
    RAW_DATA_DIR, SPEC_DIR,
    SAMPLE_RATE, IMG_SIZE,
    N_MELS, N_FFT, HOP_LENGTH,
    N_MFCC, N_BINS, BINS_PER_OCTAVE, N_CHROMA
)


# ── LOAD AUDIO ────────────────────────────────────────────────────────────────

def load_audio(filepath, target_duration=3.0):
    """
    Baca file .wav dan normalkan panjangnya ke target_duration detik.

    Kenapa di-pad/potong ke 3 detik?
    RAVDESS punya durasi bervariasi (rata-rata 3-4 detik). Model butuh
    input ukuran tetap. 3 detik = 66150 sampel di 22050 Hz, cukup
    menangkap satu ungkapan emosi penuh tanpa terlalu banyak padding.

    Padding pakai `reflect` bukan nol (zero-pad) supaya spektrogram
    tidak punya "dinding" diam di tepi yang bisa jadi artefak.
    """
    target_samples = int(SAMPLE_RATE * target_duration)

    y, sr = librosa.load(filepath, sr=SAMPLE_RATE)

    if len(y) < target_samples:
        # Audio terlalu pendek: pad dengan mode reflect
        y = np.pad(y, (0, target_samples - len(y)), mode='reflect')
    elif len(y) > target_samples:
        # Audio terlalu panjang: ambil dari tengah supaya tidak kehilangan
        # bagian penting di awal/akhir ucapan
        start = (len(y) - target_samples) // 2
        y = y[start: start + target_samples]

    return y, sr


# ── EKSTRAKSI SPEKTROGRAM ─────────────────────────────────────────────────────

def extract_mel(y, sr):
    """
    Mel Spectrogram dalam skala decibel.

    librosa.feature.melspectrogram() menghasilkan power spectrogram
    lalu memetakannya ke mel filterbank. Kita konversi ke dB supaya
    rentang nilai lebih manageable (log-compression).

    n_mels=128 → 128 baris di spektrogram
    n_fft=2048 → FFT window ~93ms
    hop_length=512 → melangkah ~23ms per kolom
    """
    S = librosa.feature.melspectrogram(
        y=y, sr=sr,
        n_mels=N_MELS, n_fft=N_FFT, hop_length=HOP_LENGTH
    )
    S_db = librosa.power_to_db(S, ref=np.max)
    return S_db


def extract_mfcc(y, sr):
    """
    Mel-Frequency Cepstral Coefficients (MFCC).

    MFCC adalah 'ringkasan' dari spektrum mel. Diambil 40 koefisien
    pertama — koefisien rendah merepresentasikan timbre vokal secara
    kasar, koefisien tinggi merepresentasikan detail halus yang jarang
    relevan.

    Dipakai luas di speech recognition sejak 1980-an. Di sini kita
    representasikan sebagai gambar 2D (koefisien vs waktu).
    """
    mfcc = librosa.feature.mfcc(
        y=y, sr=sr,
        n_mfcc=N_MFCC, n_fft=N_FFT, hop_length=HOP_LENGTH
    )
    return mfcc


def extract_cqt(y, sr):
    """
    Constant-Q Transform (CQT) dalam skala decibel.

    Berbeda dari FFT yang punya resolusi frekuensi seragam, CQT punya
    resolusi yang proporsional terhadap frekuensi (seperti piano).
    Ini cocok untuk sinyal bicara yang kontur nadanya penting.

    n_bins=84, bins_per_octave=12 → 7 oktav dengan 12 nada per oktav,
    persis seperti skala musik standar.
    """
    C = np.abs(librosa.cqt(
        y=y, sr=sr,
        n_bins=N_BINS, bins_per_octave=BINS_PER_OCTAVE,
        hop_length=HOP_LENGTH
    ))
    C_db = librosa.amplitude_to_db(C, ref=np.max)
    return C_db


def extract_chroma(y, sr):
    """
    Chroma Energy Normalized Statistics (CHROMA_STFT).

    Merangkum distribusi energi per kelas nada (C, C#, D, ..., B).
    12 baris = 12 kelas nada dalam satu oktav.

    Berguna untuk mendeteksi pola harmonis dalam ucapan emosional.
    Emosi "happy" cenderung punya kontur nada naik, "sad" turun —
    chroma menangkap pola ini dengan cara berbeda dari mel.
    """
    chroma = librosa.feature.chroma_stft(
        y=y, sr=sr,
        n_fft=N_FFT, hop_length=HOP_LENGTH
    )
    return chroma


# ── KONVERSI KE GAMBAR ────────────────────────────────────────────────────────

def spectrogram_to_image(spec_data):
    """
    Ubah array 2D spektrogram jadi gambar grayscale 128×128 piksel.

    Langkah:
    1. Normalisasi ke [0, 255] — nilai terendah → hitam, tertinggi → putih
    2. Resize ke IMG_SIZE dengan interpolasi area (INTER_AREA),
       bagus untuk downscaling karena tidak aliasing
    3. Simpan sebagai uint8 (8-bit grayscale)

    Kenapa grayscale, bukan colormap (seperti viridis)?
    Karena kita akan stack 4 channel menjadi tensor 4-channel sendiri.
    Colormap hanya akan mengacaukan karena menambah dimensi warna
    yang tidak kita butuhkan.
    """
    # Normalisasi ke [0, 255]
    spec_min = spec_data.min()
    spec_max = spec_data.max()
    if spec_max - spec_min > 1e-8:
        normalized = (spec_data - spec_min) / (spec_max - spec_min) * 255
    else:
        normalized = np.zeros_like(spec_data)

    img = normalized.astype(np.uint8)

    # Resize ke 128×128
    img_resized = cv2.resize(img, IMG_SIZE, interpolation=cv2.INTER_AREA)
    return img_resized


# ── PROSES SATU FILE ──────────────────────────────────────────────────────────

def process_one_file(wav_path, output_dir):
    """
    Proses satu file .wav: load audio → 4 spektrogram → 4 gambar PNG.

    Mengembalikan True kalau berhasil, False kalau ada error.
    Error dilog tapi tidak menghentikan loop — supaya satu file
    yang korup tidak menggagalkan seluruh dataset.
    """
    try:
        basename = os.path.splitext(os.path.basename(wav_path))[0]

        # Cek apakah sudah diproses sebelumnya (resume friendly)
        expected = [f"{basename}_{t}.png" for t in ['mel', 'mfcc', 'cqt', 'chroma']]
        if all(os.path.exists(os.path.join(output_dir, f)) for f in expected):
            return True  # sudah ada, lewati

        # Load audio
        y, sr = load_audio(wav_path)

        # Ekstrak 4 representasi
        specs = {
            'mel':    extract_mel(y, sr),
            'mfcc':   extract_mfcc(y, sr),
            'cqt':    extract_cqt(y, sr),
            'chroma': extract_chroma(y, sr),
        }

        # Simpan masing-masing sebagai PNG grayscale
        for spec_type, spec_data in specs.items():
            img = spectrogram_to_image(spec_data)
            out_path = os.path.join(output_dir, f"{basename}_{spec_type}.png")
            cv2.imwrite(out_path, img)

        return True

    except Exception as e:
        print(f"\nError memproses {os.path.basename(wav_path)}: {e}")
        return False


# ── VISUALISASI SAMPEL ────────────────────────────────────────────────────────

def visualize_sample(wav_path):
    """
    Tampilkan 4 spektrogram dari satu file .wav berdampingan.

    Berguna untuk cek visual sebelum melanjutkan ke step3.
    Pastikan gambar terlihat wajar — tidak semua hitam atau semua putih.
    """
    y, sr = load_audio(wav_path)
    specs = {
        'Mel Spectrogram': extract_mel(y, sr),
        'MFCC':            extract_mfcc(y, sr),
        'CQT':             extract_cqt(y, sr),
        'Chroma':          extract_chroma(y, sr),
    }

    fig, axes = plt.subplots(1, 4, figsize=(16, 4))
    fig.suptitle(f"Spektrogram: {os.path.basename(wav_path)}", fontsize=12)

    for ax, (title, S) in zip(axes, specs.items()):
        im = ax.imshow(S, aspect='auto', origin='lower', cmap='magma')
        ax.set_title(title)
        ax.set_xlabel('Frame')
        ax.set_ylabel('Bin')
        plt.colorbar(im, ax=ax, format='%+2.0f dB')

    plt.tight_layout()
    plt.show()


# ── MAIN ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("STEP 2: Ekstraksi Spektrogram")
    print("=" * 45)

    os.makedirs(SPEC_DIR, exist_ok=True)

    # Temukan semua file .wav
    wav_files = sorted(glob.glob(f"{RAW_DATA_DIR}/**/*.wav", recursive=True))
    if not wav_files:
        wav_files = sorted(glob.glob(f"{RAW_DATA_DIR}/*.wav"))

    if not wav_files:
        print("Tidak ada file .wav ditemukan. Jalankan step1 dulu.")
        sys.exit(1)

    print(f"Ditemukan {len(wav_files)} file .wav")
    print(f"Output ke: {SPEC_DIR}")

    # Visualisasi sampel pertama sebagai sanity check
    print("\nVisualisasi sampel pertama:")
    visualize_sample(wav_files[0])

    # Proses semua file
    success = 0
    fail = 0
    for wav_path in tqdm(wav_files, desc="Ekstraksi", unit="file"):
        ok = process_one_file(wav_path, SPEC_DIR)
        if ok:
            success += 1
        else:
            fail += 1

    print(f"\nSelesai: {success} berhasil, {fail} gagal")

    # Verifikasi jumlah output
    png_count = len(glob.glob(f"{SPEC_DIR}/*.png"))
    expected_count = len(wav_files) * 4  # 4 jenis per file
    print(f"PNG yang dihasilkan: {png_count} (expected: {expected_count})")

    if png_count >= expected_count * 0.99:  # toleransi 1%
        print("Step 2 selesai. Lanjut ke step3_dip_enhancement.py")
    else:
        print("Ada file yang hilang — periksa error di atas.")
