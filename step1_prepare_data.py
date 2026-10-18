# ============================================================
# step1_prepare_data.py — Persiapan Dataset RAVDESS
# ============================================================
# Jalankan PERTAMA sebelum step lainnya.
# Script ini akan:
#   1. Mount Google Drive
#   2. Download dataset RAVDESS dari Kaggle
#   3. Ekstrak file .wav ke folder yang benar
#   4. Verifikasi jumlah file dan label
# ============================================================

import os
import sys
import zipfile
import glob
import shutil
from collections import Counter

# ── Tambahkan folder project ke path ────────────────────────
sys.path.append('/content/drive/MyDrive/SER_Research')
from config import *

# ── 1. MOUNT GOOGLE DRIVE ────────────────────────────────────
def mount_drive():
    try:
        from google.colab import drive
        drive.mount('/content/drive')
        print("✅ Google Drive berhasil di-mount.")
    except Exception as e:
        print(f"⚠️  Tidak bisa mount Drive (mungkin sudah ter-mount): {e}")

# ── 2. INSTALL LIBRARY ───────────────────────────────────────
def install_libraries():
    os.system("pip install -q kaggle librosa opencv-python-headless "
              "tensorflow matplotlib seaborn scikit-learn tf-keras-vis")
    print("✅ Library berhasil diinstall.")

# ── 3. SETUP KAGGLE ─────────────────────────────────────────
def setup_kaggle():
    """
    Upload kaggle.json dulu ke Colab sebelum menjalankan ini.
    Cara upload: Files (panel kiri) → Upload → pilih kaggle.json
    """
    kaggle_dir = os.path.expanduser("~/.kaggle")
    os.makedirs(kaggle_dir, exist_ok=True)
    if os.path.exists("/content/kaggle.json"):
        shutil.copy("/content/kaggle.json", f"{kaggle_dir}/kaggle.json")
        os.chmod(f"{kaggle_dir}/kaggle.json", 0o600)
        print("✅ Kaggle API key berhasil disimpan.")
    else:
        print("⚠️  File kaggle.json tidak ditemukan di /content/")
        print("    Upload dulu kaggle.json sebelum lanjut.")
        print("    Cara: Kaggle → Account → API → Create New Token")

# ── 4. DOWNLOAD RAVDESS ──────────────────────────────────────
def download_ravdess():
    """
    Download RAVDESS dari Kaggle jika belum ada.
    Dataset: uwrfkaggler/ravdess-emotional-speech-audio
    """
    zip_path = "/content/ravdess.zip"
    if os.path.exists(RAW_DATA_DIR) and len(glob.glob(f"{RAW_DATA_DIR}/**/*.wav", recursive=True)) > 0:
        print(f"✅ Data sudah ada di {RAW_DATA_DIR}, skip download.")
        return

    print("📥 Mendownload RAVDESS dari Kaggle...")
    ret = os.system(
        f"kaggle datasets download -d uwrfkaggler/ravdess-emotional-speech-audio "
        f"-p /content/ --unzip"
    )
    if ret != 0:
        print("❌ Download gagal. Cek koneksi atau kaggle.json.")
        return

    # Pindahkan ke RAW_DATA_DIR
    os.makedirs(RAW_DATA_DIR, exist_ok=True)
    wav_files = glob.glob("/content/audio_speech_actors_01-24/**/*.wav", recursive=True)
    for f in wav_files:
        dest = os.path.join(RAW_DATA_DIR, os.path.basename(f))
        shutil.copy(f, dest)
    print(f"✅ {len(wav_files)} file .wav dipindahkan ke {RAW_DATA_DIR}")

# ── 5. VERIFIKASI DATASET ────────────────────────────────────
def verify_dataset():
    """
    Verifikasi jumlah file dan distribusi label emosi.
    Format nama file RAVDESS:
    03-01-[emotion]-[intensity]-[statement]-[repetition]-[actor].wav
    Kolom ke-3 (index 2) = kode emosi
    """
    wav_files = glob.glob(f"{RAW_DATA_DIR}/**/*.wav", recursive=True)
    if len(wav_files) == 0:
        wav_files = glob.glob(f"{RAW_DATA_DIR}/*.wav")

    print(f"\n{'='*50}")
    print(f"📊 VERIFIKASI DATASET RAVDESS")
    print(f"{'='*50}")
    print(f"Total file .wav : {len(wav_files)}")

    if len(wav_files) == 0:
        print("❌ Tidak ada file .wav ditemukan!")
        return False

    # Hitung distribusi emosi
    emotion_counts = Counter()
    errors = []
    for f in wav_files:
        fname = os.path.basename(f)
        parts = fname.split('-')
        if len(parts) >= 3:
            emo_code = parts[2]
            emo_name = EMOTION_MAP.get(emo_code, f"unknown({emo_code})")
            emotion_counts[emo_name] += 1
        else:
            errors.append(fname)

    print(f"\nDistribusi emosi:")
    for emo, count in sorted(emotion_counts.items()):
        bar = '█' * (count // 10)
        print(f"  {emo:<12}: {count:>4} file  {bar}")

    if errors:
        print(f"\n⚠️  {len(errors)} file dengan format nama tidak dikenali.")

    total_expected = 1440
    if len(wav_files) >= total_expected:
        print(f"\n✅ Dataset lengkap ({len(wav_files)} file).")
    else:
        print(f"\n⚠️  Dataset mungkin tidak lengkap. "
              f"Expected ~{total_expected}, got {len(wav_files)}.")

    return len(wav_files) > 0

# ── MAIN ─────────────────────────────────────────────────────
if __name__ == "__main__":
    print("🚀 STEP 1: Persiapan Dataset RAVDESS")
    print("="*50)

    mount_drive()
    create_dirs()
    install_libraries()
    setup_kaggle()
    download_ravdess()
    ok = verify_dataset()

    if ok:
        print("\n✅ Step 1 selesai! Lanjut ke step2_extract_spectrograms.py")
    else:
        print("\n❌ Step 1 GAGAL. Periksa error di atas.")
