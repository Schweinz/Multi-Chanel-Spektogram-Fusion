# SER Research Pipeline
### Speech Emotion Recognition dengan Multi-Channel Spectrogram Fusion & DIP Enhancement

**Peneliti:** Jehezkiel Beny Brian | NIM: 103012300241  
**Program:** S1 Informatika, FIF, Telkom University

---

## 📁 Struktur File

```
SER_Research/          ← folder ini di Google Drive kamu
├── config.py
├── step1_prepare_data.py
├── step2_extract_spectrograms.py
├── step3_dip_enhancement.py
├── step4_build_dataset.py
├── step5_train.py
├── step6_evaluate_gradcam.py
├── requirements.txt
├── README.md
│
├── data/
│   ├── raw/           ← file .wav RAVDESS (1440 file)
│   ├── spectrograms/  ← gambar spektrogram asli (5760 PNG)
│   ├── enhanced/      ← gambar setelah DIP (5760 PNG)
│   └── fused/         ← dataset.npz (tensor 4-channel)
│
├── models/            ← model B1–B4 (.keras)
├── results/           ← grafik, tabel evaluasi
└── gradcam/           ← output Grad-CAM per channel
```

---

## 🚀 Cara Menjalankan (Google Colab)

### Persiapan Awal (Sekali Saja)

1. **Buka Google Colab:** https://colab.research.google.com
2. **Upload semua file `.py` ini ke Google Drive** di folder:  
   `My Drive/SER_Research/`
3. **Dapatkan Kaggle API Key:**
   - Login ke kaggle.com → Account → API → **Create New Token**
   - File `kaggle.json` akan terdownload
   - Upload `kaggle.json` ke Colab (panel kiri → Files → Upload)

---

### Step 1 — Persiapan Data

**File:** `step1_prepare_data.py`

```python
# Di sel baru Google Colab, jalankan:
%cd /content/drive/MyDrive/SER_Research
!python step1_prepare_data.py
```

**Yang dilakukan:**
- Mount Google Drive
- Install semua library yang dibutuhkan
- Download RAVDESS dari Kaggle (±600 MB)
- Verifikasi 1440 file .wav & distribusi 8 kelas emosi

**Output:** `data/raw/` berisi 1440 file .wav

**Estimasi waktu:** 5–10 menit (tergantung koneksi)

---

### Step 2 — Ekstraksi Spektrogram

**File:** `step2_extract_spectrograms.py`

```python
!python step2_extract_spectrograms.py
```

**Yang dilakukan:**
- Load setiap file .wav (dipotong/padding ke 3 detik)
- Ekstrak 4 jenis spektrogram: Mel, MFCC, CQT, Chroma
- Simpan sebagai gambar PNG 128×128 grayscale
- Visualisasi 4 spektrogram dari 1 sampel

**Output:** `data/spectrograms/` berisi 5760 PNG  
**Estimasi waktu:** 30–60 menit

---

### Step 3 — DIP Enhancement

**File:** `step3_dip_enhancement.py`

```python
!python step3_dip_enhancement.py
```

**Yang dilakukan:**
- **Tahap 1:** CLAHE (clipLimit=2.0, tileGrid=8×8)
  → Perkuat kontras lokal adaptif  
  *(Referensi: Buriboev et al., Sensors 2024)*
- **Tahap 2:** Unsharp Masking (sigma=1.0, strength=1.5)
  → Pertajam detail frekuensi halus  
  *(Referensi: Tan et al., IJACSA 2022)*
- Visualisasi perbandingan Before→CLAHE→CLAHE+UM

**Output:** `data/enhanced/` berisi 5760 PNG yang sudah di-enhance  
**Estimasi waktu:** 15–30 menit

---

### Step 4 — Bangun Dataset

**File:** `step4_build_dataset.py`

```python
!python step4_build_dataset.py
```

**Yang dilakukan:**
- Fusi 4 channel (Mel+MFCC+CQT+Chroma) → tensor 128×128×4
- Ekstrak label dari nama file RAVDESS
- Split stratified: 70% train / 15% val / 15% test
- Simpan sebagai `dataset.npz`
- Visualisasi distribusi kelas per split

**Output:** `data/fused/dataset.npz`  
**Estimasi waktu:** 10–20 menit

---

### Step 5 — Training Model

**File:** `step5_train.py`

```python
!python step5_train.py
```

**Yang dilakukan:**
- Training 4 kondisi ablasi:
  - **B1:** Single-channel Mel, tanpa DIP (baseline)
  - **B2:** Single-channel Mel, dengan DIP
  - **B3:** 4-channel fused, tanpa DIP
  - **B4:** 4-channel fused, dengan DIP ← **PROPOSED**
- Arsitektur: MobileNetV2 + Conv2D(1×1) channel projection
- 2-phase training: frozen (10 epoch) → fine-tune (max 100 epoch)
- Early stopping, ReduceLROnPlateau, ModelCheckpoint

**Output:** `models/B1_*.keras`, `B2_*.keras`, `B3_*.keras`, `B4_*.keras`  
**Estimasi waktu:** 2–6 jam (dengan GPU Colab T4)

> ⚡ **Tips:** Aktifkan GPU di Colab: Runtime → Change runtime type → T4 GPU

---

### Step 6 — Evaluasi & Grad-CAM

**File:** `step6_evaluate_gradcam.py`

```python
!python step6_evaluate_gradcam.py
```

**Yang dilakukan:**
- Evaluasi metrik: Accuracy, Precision, Recall, Macro F1 per kelas
- Confusion matrix per kondisi ablasi
- **Per-Channel Grad-CAM** (NOVELTY penelitian):
  - Visualisasi saliency map terpisah untuk tiap channel
  - Analisis channel importance per kelas emosi → heatmap 8×4
- Grafik perbandingan ablation study (B1 vs B2 vs B3 vs B4)

**Output:**
- `results/cm_b*.png` — confusion matrix
- `results/ablation_comparison.png` — grafik perbandingan
- `results/channel_importance_heatmap.png` — heatmap importance
- `results/evaluation_results.json` — angka metrik
- `gradcam/gradcam_*.png` — visualisasi per-channel Grad-CAM

**Estimasi waktu:** 30–60 menit

---

## 🏗️ Arsitektur Model (B3 & B4)

```
Input (128×128×4)
    ↓
Conv2D(3, kernel=1×1)    ← Channel projection: 4 → 3
    ↓
MobileNetV2 (pretrained ImageNet, trainable phase 2)
    ↓
GlobalAveragePooling2D
    ↓
Dropout(0.3)
    ↓
Dense(128, ReLU) + BatchNorm
    ↓
Dropout(0.3)
    ↓
Dense(8, Softmax)        ← 8 kelas emosi
```

---

## 📊 Dataset RAVDESS

- **Jumlah file:** 1440 file .wav
- **Aktor:** 24 (12 pria, 12 wanita)
- **Kelas emosi:** 8 (neutral, calm, happy, sad, angry, fearful, disgust, surprised)
- **Distribusi:** Seimbang (~180 sampel per kelas)
- **Sumber Kaggle:** `uwrfkaggler/ravdess-emotional-speech-audio`

---

## 🔬 Novelty Penelitian

1. **Multi-Channel Spectrogram Fusion:** Kombinasi Mel+MFCC+CQT+Chroma sebagai input 4-channel ke CNN
2. **DIP pada Level Gambar Spektrogram:** CLAHE + Unsharp Masking diterapkan sebagai preprocessing image
3. **Channel Projection Layer:** Conv2D(1×1) untuk adaptasi pretrained 3-channel model ke 4-channel input
4. **Per-Channel Grad-CAM:** Analisis kontribusi tiap jenis spektrogram terhadap prediksi emosi

---

## ⚙️ Konfigurasi

Semua parameter terpusat di `config.py`. Parameter yang bisa disesuaikan:

| Parameter | Default | Keterangan |
|-----------|---------|------------|
| `CLAHE_CLIP_LIMIT` | 2.0 | Batas amplifikasi CLAHE |
| `UM_SIGMA` | 1.0 | Sigma Gaussian blur untuk UM |
| `UM_STRENGTH` | 1.5 | Kekuatan sharpening UM |
| `BATCH_SIZE` | 32 | Ukuran batch training |
| `EPOCHS_FROZEN` | 10 | Epoch phase 1 (frozen) |
| `EPOCHS_UNFROZEN` | 100 | Max epoch phase 2 |
| `DROPOUT_RATE` | 0.3 | Dropout rate |
| `LEARNING_RATE` | 1e-4 | Learning rate awal |

---

## 📚 Referensi Kunci

- Buriboev et al. (2024). CLAHE-based enhancement. *Sensors*, DOI: 10.3390/s24237703
- Tan et al. (2022). Unsharp masking enhancement. *IJACSA*, Vol.13 No.8
- Howard et al. (2018). MobileNetV2. *CVPR 2018*
- Selvaraju et al. (2017). Grad-CAM. *ICCV 2017*
- Livingstone & Russo (2018). RAVDESS dataset. *PLoS ONE*
