"""
step6_evaluate_gradcam.py
==========================
STEP TERAKHIR — evaluasi performa model dan analisis per-channel dengan Grad-CAM.

Yang dilakukan di step ini:
  1. Muat semua model yang sudah ditraining di step5 (B1–B4)
  2. Hitung metrik evaluasi lengkap per kondisi:
       accuracy, precision, recall, F1 per kelas emosi
  3. Plot confusion matrix untuk setiap kondisi
  4. NOVELTY: Grad-CAM per channel untuk model B4
       → Biasanya Grad-CAM menghasilkan satu heatmap untuk seluruh input.
         Di sini kita pisahkan kontribusi tiap channel (Mel, MFCC, CQT, Chroma)
         untuk memahami: mana channel yang paling diperhatikan model
         untuk mengenali tiap emosi?
  5. Heatmap kepentingan channel × emosi (8 emosi × 4 channel)
  6. Bar chart perbandingan akurasi dan F1 antar kondisi ablasi

Output: folder RESULT_DIR dan GRADCAM_DIR
"""

import os
import sys
import json
import numpy as np
import cv2
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import (
    classification_report, confusion_matrix,
    accuracy_score, f1_score
)

sys.path.append('/content/drive/MyDrive/SER_Research')
from config import (
    FUSED_DIR, SPEC_DIR, MODEL_DIR, RESULT_DIR, GRADCAM_DIR,
    EMOTION_LABELS, N_CLASSES, CHANNEL_ORDER, N_CHANNELS
)

import tensorflow as tf
from tensorflow.keras.models import load_model
from tensorflow.keras.utils import to_categorical

# tf-keras-vis untuk Grad-CAM visualization
try:
    from tf_keras_vis.gradcam import Gradcam
    from tf_keras_vis.utils.scores import CategoricalScore
    from tf_keras_vis.utils.model_modifiers import ReplaceToLinear
    TF_KERAS_VIS_OK = True
except ImportError:
    print("tf-keras-vis tidak tersedia. Install dengan: pip install tf-keras-vis")
    print("Bagian Grad-CAM akan dilewati.")
    TF_KERAS_VIS_OK = False


# ── LOAD MODEL DAN DATA ───────────────────────────────────────────────────────

def load_models():
    """
    Load semua model .keras dari MODEL_DIR.

    Nama file mengikuti konvensi yang ditetapkan di step5:
    B1_mel_nodip_best.keras, B2_mel_dip_best.keras, dst.

    Kalau satu model tidak ditemukan, kondisi itu dilewati tapi
    kondisi lain tetap dievaluasi.
    """
    models = {}
    name_map = {
        'B1': 'B1_mel_nodip_best.keras',
        'B2': 'B2_mel_dip_best.keras',
        'B3': 'B3_4ch_nodip_best.keras',
        'B4': 'B4_4ch_dip_best.keras',
    }
    for cond, fname in name_map.items():
        path = os.path.join(MODEL_DIR, fname)
        if os.path.exists(path):
            models[cond] = load_model(path)
            print(f"Model {cond} dimuat dari {path}")
        else:
            print(f"Model {cond} tidak ditemukan di {path}")
    return models


def load_dataset_npz():
    """Load dataset split yang dibuat step4."""
    data_path = os.path.join(FUSED_DIR, "dataset.npz")
    if not os.path.exists(data_path):
        raise FileNotFoundError(
            f"File {data_path} tidak ditemukan. Jalankan step4 dulu."
        )
    data = np.load(data_path)
    return (
        data['X_train'], data['y_train'],
        data['X_val'],   data['y_val'],
        data['X_test'],  data['y_test']
    )


def prepare_x(X, condition):
    """
    Sesuaikan input shape berdasarkan kondisi ablasi.
    B1/B2: duplikasi channel Mel jadi 3 channel.
    B3/B4: pakai 4 channel langsung.
    """
    if condition in ('B1', 'B2'):
        mel = X[:, :, :, 0:1]
        return np.concatenate([mel, mel, mel], axis=-1)
    return X


# ── CONFUSION MATRIX ──────────────────────────────────────────────────────────

def plot_confusion_matrix(y_true, y_pred, title, save_name):
    """
    Plot confusion matrix dengan anotasi jumlah absolut dan warna normalized.

    Warna (normalized) memudahkan melihat pola distribusi error,
    angka (absolut) memudahkan menghitung berapa sampel yang salah.
    """
    cm = confusion_matrix(y_true, y_pred)
    cm_norm = cm.astype(float) / cm.sum(axis=1, keepdims=True)

    fig, ax = plt.subplots(figsize=(10, 8))
    sns.heatmap(
        cm_norm,
        annot=cm,          # angka absolut di dalam sel
        fmt='d',
        xticklabels=EMOTION_LABELS,
        yticklabels=EMOTION_LABELS,
        cmap='Blues',
        ax=ax
    )
    ax.set_title(title, fontsize=13, fontweight='bold')
    ax.set_xlabel('Predicted')
    ax.set_ylabel('True')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()

    save_path = os.path.join(RESULT_DIR, f"cm_{save_name}.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.close()
    print(f"  Confusion matrix disimpan: {save_path}")


# ── EVALUASI LENGKAP ──────────────────────────────────────────────────────────

def evaluate_model(model, X_test, y_test, condition, label):
    """
    Hitung semua metrik evaluasi untuk satu kondisi ablasi.

    Metrik yang dihitung:
      accuracy     → persentase prediksi benar dari semua sampel
      precision    → dari yang diprediksi kelas X, berapa yang benar X?
      recall       → dari semua sampel kelas X yang asli, berapa yang terdeteksi?
      F1 macro     → rata-rata F1 per kelas (tidak berbobot jumlah sampel)

    Kenapa macro F1, bukan weighted?
    Dengan macro, setiap kelas emosi berkontribusi sama besar ke metrik akhir.
    Ini lebih fair untuk evaluasi karena RAVDESS relatif seimbang antar kelas.
    """
    print(f"\n{'='*55}")
    print(f"Evaluasi {condition}: {label}")
    print("=" * 55)

    X_in = prepare_x(X_test, condition)
    y_oh = to_categorical(y_test, N_CLASSES)
    loss, acc = model.evaluate(X_in, y_oh, verbose=0)

    y_pred_proba = model.predict(X_in, verbose=0)
    y_pred = np.argmax(y_pred_proba, axis=1)

    print(classification_report(y_test, y_pred, target_names=EMOTION_LABELS))
    macro_f1 = f1_score(y_test, y_pred, average='macro')

    plot_confusion_matrix(
        y_test, y_pred,
        title=f"Confusion Matrix — {condition}: {label}",
        save_name=condition.lower()
    )

    return {
        'accuracy': float(acc),
        'loss':     float(loss),
        'macro_f1': float(macro_f1),
        'report':   classification_report(
            y_test, y_pred, target_names=EMOTION_LABELS, output_dict=True
        )
    }


# ── PER-CHANNEL GRAD-CAM ──────────────────────────────────────────────────────

def compute_gradcam_per_channel(model_b4, X_test, y_test, sample_idx=0):
    """
    Hitung Grad-CAM pada model B4 untuk satu sampel.

    Grad-CAM (Gradient-weighted Class Activation Mapping):
    Prinsip: kelas yang model prediksi paling bergantung pada fitur apa?
    Cara: ambil gradien loss terhadap feature map terakhir, lalu rata-ratakan
    gradien itu per channel feature map → bobot pentingnya.
    Kalikan bobot ke feature map → heatmap 2D.

    Referensi: Selvaraju et al., ICCV 2017.

    Kenapa layer 'global_average_pooling2d' sebagai penultimate layer?
    Grad-CAM bekerja terbaik pada layer konvolusional terakhir sebelum
    pooling. Di MobileNetV2, itu adalah conv output sebelum GAP.
    """
    if not TF_KERAS_VIS_OK:
        return None

    gradcam = Gradcam(
        model_b4,
        model_modifier=ReplaceToLinear(),
        clone=True
    )

    x = X_test[sample_idx:sample_idx + 1]   # shape: (1, 128, 128, 4)
    score = CategoricalScore([int(y_test[sample_idx])])

    cam = gradcam(score, x, penultimate_layer='global_average_pooling2d')
    heatmap = cam[0]   # shape: (128, 128)
    return heatmap


def visualize_gradcam_per_channel(model_b4, X_test, y_test, n_samples=5):
    """
    Untuk setiap sampel: tampilkan gambar spektrogram 4 channel
    berdampingan dengan Grad-CAM overlay-nya.

    Layout:
      Baris atas:  4 gambar spektrogram asli (mel, mfcc, cqt, chroma)
      Baris bawah: 4 Grad-CAM overlay (heatmap di-blend ke atas gambar)

    Blend: 60% gambar asli + 40% colormap jet dari heatmap.
    Ini mempertahankan struktur spektrogram tapi memberi warna merah
    di area yang "paling diperhatikan" model.
    """
    if not TF_KERAS_VIS_OK:
        print("tf-keras-vis tidak tersedia, skip visualisasi Grad-CAM.")
        return

    os.makedirs(GRADCAM_DIR, exist_ok=True)

    # Pilih sampel representatif — ambil 1-2 dari setiap kelas emosi
    samples_per_class = max(1, n_samples // N_CLASSES)
    selected = []
    for cls in range(N_CLASSES):
        idx_cls = np.where(y_test == cls)[0]
        if len(idx_cls) > 0:
            selected.extend(idx_cls[:samples_per_class].tolist())

    print(f"Grad-CAM per channel untuk {min(len(selected), n_samples)} sampel...")

    for idx in selected[:n_samples]:
        emotion = EMOTION_LABELS[y_test[idx]]
        heatmap = compute_gradcam_per_channel(model_b4, X_test, y_test, idx)
        if heatmap is None:
            continue

        # Normalisasi heatmap ke [0, 1]
        h_min, h_max = heatmap.min(), heatmap.max()
        heatmap = (heatmap - h_min) / (h_max - h_min + 1e-8)

        fig, axes = plt.subplots(2, 4, figsize=(20, 10))
        fig.suptitle(
            f'Per-Channel Grad-CAM  |  Sampel #{idx}  |  Emosi: {emotion}',
            fontsize=13, fontweight='bold'
        )

        for col, ch_name in enumerate(CHANNEL_ORDER):
            ch_img = X_test[idx, :, :, col]   # gambar satu channel, [0,1]

            # Baris atas: gambar spektrogram asli
            axes[0, col].imshow(ch_img, cmap='viridis', aspect='auto', origin='lower')
            axes[0, col].set_title(f'{ch_name.upper()}\nSpectrogram')
            axes[0, col].axis('off')

            # Baris bawah: Grad-CAM overlay
            # Konversi grayscale ke RGB untuk blending
            ch_rgb = plt.cm.viridis(ch_img)[:, :, :3]
            heatmap_rgb = plt.cm.jet(heatmap)[:, :, :3]
            overlay = np.clip(0.6 * ch_rgb + 0.4 * heatmap_rgb, 0, 1)

            axes[1, col].imshow(overlay, aspect='auto', origin='lower')
            axes[1, col].set_title(f'{ch_name.upper()}\nGrad-CAM Overlay')
            axes[1, col].axis('off')

        plt.tight_layout()
        save_path = os.path.join(GRADCAM_DIR, f"gradcam_{emotion}_idx{idx}.png")
        plt.savefig(save_path, dpi=100, bbox_inches='tight')
        plt.close()

    print(f"Grad-CAM disimpan di: {GRADCAM_DIR}")


def analyze_channel_importance(model_b4, X_test, y_test):
    """
    Hitung kepentingan tiap channel untuk tiap kelas emosi.

    Caranya: jalankan Grad-CAM untuk setiap sampel, lalu kalikan
    heatmap dengan gambar channel. Nilai tinggi di hasil kali artinya
    "di sini area penting DAN channelnya aktif" → model benar-benar
    memperhatikan fitur yang ada di lokasi itu.

    Hasil: matriks 8×4 → (8 emosi) × (4 channel)
    Diplot sebagai heatmap untuk menjawab:
    "Untuk emosi 'angry', channel mana yang paling berkontribusi?"
    """
    if not TF_KERAS_VIS_OK:
        return None

    print("\nAnalisis kepentingan channel per emosi...")
    gradcam = Gradcam(
        model_b4,
        model_modifier=ReplaceToLinear(),
        clone=True
    )

    # importance[emosi_idx, channel_idx]
    importance = np.zeros((N_CLASSES, N_CHANNELS))

    for cls in range(N_CLASSES):
        idx_list = np.where(y_test == cls)[0][:10]  # max 10 sampel
        if len(idx_list) == 0:
            continue

        score = CategoricalScore([cls])
        for idx in idx_list:
            x = X_test[idx:idx + 1]
            cam = gradcam(score, x, penultimate_layer='global_average_pooling2d')
            heatmap = cam[0]

            for ch in range(N_CHANNELS):
                # Bobot per channel = rata-rata (channel × heatmap)
                ch_weight = float(np.mean(X_test[idx, :, :, ch] * heatmap))
                importance[cls, ch] += ch_weight

        importance[cls] /= len(idx_list)

    # Plot heatmap kepentingan
    fig, ax = plt.subplots(figsize=(10, 8))
    sns.heatmap(
        importance,
        xticklabels=[c.upper() for c in CHANNEL_ORDER],
        yticklabels=EMOTION_LABELS,
        cmap='YlOrRd',
        annot=True, fmt='.3f',
        ax=ax
    )
    ax.set_title(
        'Kepentingan Channel per Emosi\n(rata-rata Grad-CAM activation weight)',
        fontsize=12, fontweight='bold'
    )
    ax.set_xlabel('Spectrogram Channel')
    ax.set_ylabel('Emotion Class')
    plt.tight_layout()

    save_path = os.path.join(RESULT_DIR, "channel_importance_heatmap.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.close()
    print(f"Channel importance heatmap disimpan ke {save_path}")

    return importance


# ── PERBANDINGAN ABLATION ─────────────────────────────────────────────────────

def plot_ablation_comparison(results_dict):
    """
    Bar chart perbandingan accuracy dan macro F1 antar kondisi B1–B4.

    Visualisasi ini adalah Tabel IV dari paper dalam bentuk grafik.
    Diharapkan B4 (metode yang diusulkan) menunjukkan performa tertinggi,
    dengan peningkatan bertahap: B1 → B2 (tambah DIP) dan B1 → B3 (tambah channel).
    """
    conditions = list(results_dict.keys())
    accuracies = [results_dict[c]['accuracy'] * 100 for c in conditions]
    f1s        = [results_dict[c]['macro_f1'] * 100 for c in conditions]

    labels_desc = {
        'B1': 'B1\nSingle-Mel\nNo DIP',
        'B2': 'B2\nSingle-Mel\n+ DIP',
        'B3': 'B3\n4-Channel\nNo DIP',
        'B4': 'B4\n4-Channel\n+ DIP\n(PROPOSED)',
    }

    x = np.arange(len(conditions))
    width = 0.35

    fig, ax = plt.subplots(figsize=(10, 7))
    bars1 = ax.bar(x - width / 2, accuracies, width,
                   label='Accuracy (%)', color='steelblue', alpha=0.85)
    bars2 = ax.bar(x + width / 2, f1s, width,
                   label='Macro F1 (%)', color='darkorange', alpha=0.85)

    ax.set_xlabel('Kondisi Ablasi')
    ax.set_ylabel('Persentase (%)')
    ax.set_title(
        'Ablation Study: Accuracy & Macro F1\n'
        'Speech Emotion Recognition — MobileNetV2',
        fontsize=12, fontweight='bold'
    )
    ax.set_xticks(x)
    ax.set_xticklabels([labels_desc.get(c, c) for c in conditions], fontsize=9)
    ax.legend()
    ax.set_ylim(0, 110)
    ax.yaxis.grid(True, alpha=0.3)

    for bar in bars1:
        ax.annotate(
            f'{bar.get_height():.1f}%',
            xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
            xytext=(0, 3), textcoords='offset points',
            ha='center', fontsize=9
        )
    for bar in bars2:
        ax.annotate(
            f'{bar.get_height():.1f}%',
            xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
            xytext=(0, 3), textcoords='offset points',
            ha='center', fontsize=9
        )

    plt.tight_layout()
    save_path = os.path.join(RESULT_DIR, "ablation_comparison.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.close()
    print(f"Grafik ablation disimpan ke {save_path}")


# ── MAIN ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("STEP 6: Evaluasi & Per-Channel Grad-CAM")
    print("=" * 55)

    os.makedirs(RESULT_DIR, exist_ok=True)
    os.makedirs(GRADCAM_DIR, exist_ok=True)

    # Load dataset
    X_train, y_train, X_val, y_val, X_test, y_test = load_dataset_npz()

    # Rebuild test set tanpa DIP untuk B1/B3
    from step4_build_dataset import build_full_dataset, split_dataset
    print("\nMemuat data tanpa DIP untuk B1/B3...")
    X_raw, y_raw = build_full_dataset(enhanced_dir=SPEC_DIR)
    _, _, _, _, X_ts_raw, y_ts_raw = split_dataset(X_raw, y_raw)

    # Load semua model
    print("\nMemuat model...")
    models = load_models()

    # Evaluasi per kondisi
    results = {}
    test_data_map = {
        'B1': (X_ts_raw, y_ts_raw),
        'B2': (X_test,   y_test),
        'B3': (X_ts_raw, y_ts_raw),
        'B4': (X_test,   y_test),
    }
    desc_map = {
        'B1': 'Single-Mel, No DIP',
        'B2': 'Single-Mel, DIP',
        'B3': '4-Channel, No DIP',
        'B4': '4-Channel, DIP (PROPOSED)',
    }

    for cond in ['B1', 'B2', 'B3', 'B4']:
        if cond not in models:
            print(f"Lewati {cond}: model tidak ditemukan.")
            continue
        X_ts, y_ts = test_data_map[cond]
        results[cond] = evaluate_model(
            models[cond], X_ts, y_ts, cond, desc_map[cond]
        )

    # Plot perbandingan
    if results:
        plot_ablation_comparison(results)

    # Simpan hasil evaluasi ke JSON
    eval_path = os.path.join(RESULT_DIR, "evaluation_results.json")
    with open(eval_path, 'w') as f:
        json.dump(
            {c: {'accuracy': r['accuracy'], 'macro_f1': r['macro_f1'], 'loss': r['loss']}
             for c, r in results.items()},
            f, indent=2
        )
    print(f"\nHasil evaluasi disimpan ke {eval_path}")

    # Grad-CAM per channel untuk B4 (metode yang diusulkan)
    if 'B4' in models:
        print("\n" + "=" * 55)
        print("Grad-CAM Per-Channel Analysis (Model B4 — PROPOSED)")
        print("=" * 55)

        # Visualisasi per sampel
        visualize_gradcam_per_channel(
            models['B4'], X_test, y_test, n_samples=8
        )

        # Heatmap kepentingan channel × emosi
        importance = analyze_channel_importance(
            models['B4'], X_test, y_test
        )
        if importance is not None:
            np.save(
                os.path.join(RESULT_DIR, "channel_importance.npy"),
                importance
            )
            print("Channel importance matrix disimpan ke channel_importance.npy")

    print("\n" + "=" * 55)
    print("Step 6 selesai! Pipeline penelitian LENGKAP.")
    print(f"  Semua hasil evaluasi ada di: {RESULT_DIR}")
    print(f"  Semua Grad-CAM ada di      : {GRADCAM_DIR}")
