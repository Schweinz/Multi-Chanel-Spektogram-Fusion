"""
step1_prepare_data.py
=====================
STEP PERTAMA — jalankan ini sebelum yang lain.

Yang dilakukan di step ini:
1. Mount Google Drive supaya data tersimpan permanen
2. Install library yang diperlukan (librosa, opencv, dll.)
3. Setup kredensial Kaggle
4. Download dataset RAVDESS dari Kaggle
5. Verifikasi file sudah lengkap dan labelnya terbaca benar

Dataset yang dipakai:
RAVDESS (Ryerson Audio-Visual Database of Emotional Speech and Song)
— Livingstone & Russo, 2018, PLoS ONE
— 24 aktor, 8 emosi, total 1440 file ucapan (speech only)
— Dikerjakan ulang oleh Kaggle: uwrfkaggler/ravdess-emotional-speech-audio

Catatan nama file RAVDESS:
03-01-05-01-02-01-12.wav
││  │  │  │  │  │  └── aktor (01–24)
││  │  │  │  │  └───── repetisi (01=pertama, 02=kedua)
││  │  │  │  └──────── statement (01="Kids are talking", 02="Dogs are sitting")
││  │  │  └─────────── intensitas (01=normal, 02=strong)
││  │  └────────────── emosi (01–08) ← yang kita pakai
││  └───────────────── vocal channel (01=speech, 02=song)
│└──────────────────── modality (01=AV, 02=video only, 03=audio only)
└───────────────────── tidak relevan
"""

import os
import sys
import glob
import shutil
from collections import Counter

sys.path.append('/content/drive/MyDrive/SER_Research')
from config import (
    RAW_DATA_DIR, EMOTION_MAP, create_dirs
)


def mount_drive():
    """
    Mount Google Drive ke /content/drive.

    Colab menyimpan file di RAM sementara — begitu session habis,
    semua hilang. Dengan mount Drive, data tersimpan permanen.
    Fungsi ini aman dipanggil berkali-kali karena Drive yang sudah
    ter-mount tidak akan di-mount ulang, hanya muncul pesan warning.
    """
    try:
        from google.colab import drive
        drive.mount('/content/drive')
        print("Google Drive sudah ter-mount.")
    except Exception as e:
        print(f"Tidak bisa mount Drive: {e}")
        print("Kalau kamu menjalankan ini di luar Colab, ini normal.")


def install_libraries():
    """
    Install semua library yang diperlukan pipeline.

    librosa       — ekstraksi fitur audio (mel, mfcc, cqt, chroma)
    opencv        — image processing (CLAHE, unsharp masking, resize)
    tensorflow    — training model MobileNetV2
    scikit-learn  — split dataset stratified, metrik evaluasi
    tf-keras-vis  — Grad-CAM visualization
    tqdm          — progress bar
    seaborn       — heatmap confusion matrix
    """
    os.system(
        "pip install -q kaggle librosa opencv-python-headless "
        "tensorflow matplotlib seaborn scikit-learn tf-keras-vis tqdm"
    )
    print("Library sudah terinstall.")


def setup_kaggle():
    """
    Salin kaggle.json ke lokasi yang dikenali oleh Kaggle CLI (~/.kaggle/).

    Cara mendapatkan kaggle.json:
    1. Buka https://www.kaggle.com/account
    2. Scroll ke bagian API → klik "Create New Token"
    3. File kaggle.json akan terdownload
    4. Di panel kiri Colab, klik ikon folder → Upload → pilih kaggle.json
       (file akan masuk ke /content/kaggle.json)

    chmod 600 wajib karena Kaggle CLI menolak file yang bisa dibaca
    oleh user lain (security requirement).
    """
    kaggle_dir = os.path.expanduser("~/.kaggle")
    os.makedirs(kaggle_dir, exist_ok=True)

    src = "/content/kaggle.json"
    dst = os.path.join(kaggle_dir, "kaggle.json")

    if os.path.exists(src):
        shutil.copy(src, dst)
        os.chmod(dst, 0o600)
        print("Kaggle API key sudah dikonfigurasi.")
    else:
        print("File kaggle.json tidak ditemukan di /content/.")
        print("Upload dulu sebelum lanjut: Files (panel kiri) → Upload.")


def download_ravdess():
    """
    Download RAVDESS speech audio dari Kaggle dan pindahkan ke RAW_DATA_DIR.

    Kenapa pakai --unzip langsung?
    Supaya tidak perlu unzip manual. File .wav tersebar di subfolder
    Actor_01/ sampai Actor_24/, jadi setelah unzip kita loop semua
    subfolder dan copy flat ke RAW_DATA_DIR.

    Kenapa disalin flat (tidak ikut struktur subfolder)?
    Karena nama file RAVDESS sudah unik per aktor (ada kode aktor
    di bagian akhir nama file), jadi tidak ada risiko konflik nama.
    Pipeline kita hanya butuh nama file untuk parsing label,
    tidak butuh tahu file ada di folder aktor mana.
    """
    # Skip download kalau data sudah ada
    existing = glob.glob(f"{RAW_DATA_DIR}/**/*.wav", recursive=True)
    if existing:
        print(f"Data sudah ada ({len(existing)} file), skip download.")
        return

    print("Mendownload RAVDESS dari Kaggle...")
    ret = os.system(
        "kaggle datasets download "
        "-d uwrfkaggler/ravdess-emotional-speech-audio "
        "-p /content/ --unzip"
    )

    if ret != 0:
        print("Download gagal. Cek koneksi internet atau validitas kaggle.json.")
        return

    # Pindahkan semua .wav ke satu folder flat
    os.makedirs(RAW_DATA_DIR, exist_ok=True)
    wav_files = glob.glob(
        "/content/audio_speech_actors_01-24/**/*.wav", recursive=True
    )
    for f in wav_files:
        shutil.copy(f, os.path.join(RAW_DATA_DIR, os.path.basename(f)))

    print(f"{len(wav_files)} file .wav dipindahkan ke {RAW_DATA_DIR}")


def verify_dataset():
    """
    Cek apakah semua 1440 file ada dan labelnya terbaca benar.

    RAVDESS speech subset: 24 aktor × 60 ucapan/aktor = 1440 file.
    60 ucapan per aktor = 8 emosi × 2 intensitas × 2 statement
    × (1 repetisi untuk neutral/calm, 2 untuk yang lain) — bervariasi.

    Fungsi ini juga jadi sanity check bahwa kode parsing EMOTION_MAP
    bekerja benar sebelum masuk step ekstraksi yang lebih lama.
    """
    wav_files = glob.glob(f"{RAW_DATA_DIR}/**/*.wav", recursive=True)
    if not wav_files:
        wav_files = glob.glob(f"{RAW_DATA_DIR}/*.wav")

    print(f"\nVerifikasi dataset RAVDESS")
    print(f"{'─'*40}")
    print(f"Total file .wav ditemukan : {len(wav_files)}")

    if not wav_files:
        print("Tidak ada file .wav. Pastikan download berhasil.")
        return False

    # Hitung distribusi per emosi
    counts = Counter()
    unrecognized = []
    for f in wav_files:
        parts = os.path.basename(f).split('-')
        if len(parts) >= 3:
            code = parts[2]
            name = EMOTION_MAP.get(code, f"?({code})")
            counts[name] += 1
        else:
            unrecognized.append(os.path.basename(f))

    print(f"\nDistribusi per emosi:")
    for emo in EMOTION_MAP.values():
        n = counts.get(emo, 0)
        bar = '█' * (n // 10)
        print(f"  {emo:<12}: {n:>4}  {bar}")

    if unrecognized:
        print(f"\nPeringatan: {len(unrecognized)} file dengan nama tidak dikenali.")

    if len(wav_files) >= 1440:
        print(f"\nDataset lengkap ({len(wav_files)} file).")
    else:
        print(f"\nDataset mungkin tidak lengkap — expected 1440, dapat {len(wav_files)}.")

    return len(wav_files) > 0


if __name__ == "__main__":
    print("STEP 1: Persiapan Dataset RAVDESS")
    print("=" * 45)

    mount_drive()
    create_dirs()
    install_libraries()
    setup_kaggle()
    download_ravdess()

    ok = verify_dataset()
    if ok:
        print("\nStep 1 selesai. Lanjut ke step2_extract_spectrograms.py")
    else:
        print("\nStep 1 gagal. Periksa pesan error di atas.")
