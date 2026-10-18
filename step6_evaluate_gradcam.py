# ============================================================
# step6_evaluate_gradcam.py — Evaluasi & Per-Channel Grad-CAM
# ============================================================
# Jalankan setelah step5_train.py
#
# Script ini akan:
#   1. Evaluasi metrik lengkap (accuracy, precision, recall, F1)
#      per kelas emosi untuk semua kondisi ablasi
#   2. Plot confusion matrix per kondisi
#   3. Grad-CAM per channel pada model B4 (PROPOSED)
#      → Novelty: visualisasi saliency map per channel
#        (Mel, MFCC, CQT, Chroma secara terpisah)
#   4. Tabel perbandingan ablation study
# ============================================================

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
from config import *

import tensorflow as tf
from tensorflow.keras.models import load_model
from tensorflow.keras.utils import to_categorical

# tf-keras-vis untuk Grad-CAM
try:
    from tf_keras_vis.gradcam import Gradcam
    from tf_keras_vis.utils.scores import CategoricalScore
    from tf_keras_vis.utils.model_modifiers import ReplaceToLinear
    TF_KERAS_VIS_OK = True
except ImportError:
    print("⚠️  tf-keras-vis tidak tersedia. Install dengan:")
    print("    pip install tf-keras-vis")
    TF_KERAS_VIS_OK = False


# ── LOAD MODEL & DATA ─────────────────────────────────────────

def load_models():
    """Load semua model ablation yang sudah ditraining."""
    models = {}
    names = {
        'B1': 'B1_mel_nodip_best.keras',
        'B2': 'B2_mel_dip_best.keras',
        'B3': 'B3_4ch_nodip_best.keras',
        'B4': 'B4_4ch_dip_best.keras',
    }
    for cond, fname in names.items():
        path = os.path.join(MODEL_DIR, fname)
        if os.path.exists(path):
            models[cond] = load_model(path)
            print(f"✅ Model {cond} dimuat dari {path}")
        else:
            print(f"⚠️  Model {cond} tidak ditemukan di {path}")
    return models


def load_dataset_npz():
    data_path = os.path.join(FUSED_DIR, "dataset.npz")
    data = np.load(data_path)
    return (
        data['X_train'], data['y_train'],
        data['X_val'],   data['y_val'],
        data['X_test'],  data['y_test']
    )


def prepare_x(X, condition):
    """Sesuaikan input shape berdasarkan kondisi ablasi."""
    if condition in ('B1', 'B2'):
        mel = X[:, :, :, 0:1]
        return np.concatenate([mel, mel, mel], axis=-1)
    return X


# ── CONFUSION MATRIX ─────────────────────────────────────────

def plot_confusion_matrix(y_true, y_pred, title, save_name):
    cm = confusion_matrix(y_true, y_pred)
    cm_norm = cm.astype(float) / cm.sum(axis=1, keepdims=True)

    fig, ax = plt.subplots(figsize=(10, 8))
    sns.heatmap(cm_norm, annot=cm, fmt='d',
                xticklabels=EMOTION_LABELS,
                yticklabels=EMOTION_LABELS,
                cmap='Blues', ax=ax)
    ax.set_title(title, fontsize=13, fontweight='bold')
    ax.set_xlabel('Predicted')
    ax.set_ylabel('True')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()

    save_path = os.path.join(RESULT_DIR, f"cm_{save_name}.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.show()
    print(f"  ✅ Confusion matrix disimpan: {save_path}")


# ── EVALUASI LENGKAP ──────────────────────────────────────────

def evaluate_model(model, X_test, y_test, condition, label):
    """Hitung semua metrik evaluasi untuk satu kondisi."""
    print(f"\n{'='*55}")
    print(f"📊 Evaluasi {condition}: {label}")
    print("="*55)

    X_in = prepare_x(X_test, condition)
    y_oh = to_categorical(y_test, N_CLASSES)
    loss, acc = model.evaluate(X_in, y_oh, verbose=0)

    y_pred_proba = model.predict(X_in, verbose=0)
    y_pred = np.argmax(y_pred_proba, axis=1)

    # Classification report
    report = classification_report(
        y_test, y_pred,
        target_names=EMOTION_LABELS,
        output_dict=True
    )
    print(classification_report(y_test, y_pred, target_names=EMOTION_LABELS))

    macro_f1 = f1_score(y_test, y_pred, average='macro')

    # Confusion matrix
    plot_confusion_matrix(
        y_test, y_pred,
        title=f"Confusion Matrix — {condition}: {label}",
        save_name=f"{condition.lower()}"
    )

    return {
        'accuracy': acc,
        'loss': loss,
        'macro_f1': macro_f1,
        'report': report
    }


# ── PER-CHANNEL GRAD-CAM (NOVELTY) ───────────────────────────

def compute_gradcam_per_channel(model_b4, X_sample, true_label, sample_idx=0):
    """
    NOVELTY: Grad-CAM per channel untuk model 4-channel.

    Standard Grad-CAM menghasilkan satu heatmap untuk seluruh input.
    Di sini kita isolasi kontribusi masing-masing channel dengan cara:
    1. Jalankan Grad-CAM pada input 4-channel penuh → heatmap gabungan
    2. Kalikan heatmap dengan masing-masing channel gambar
       → saliency per channel yang terpisah

    Ini memberikan pemahaman:
    - Channel mana yang paling informatif untuk tiap emosi?
    - Frekuensi mana yang "diperhatikan" model per channel?

    Referensi konsep: Selvaraju et al., ICCV 2017 (Grad-CAM)
    Ekstensi per-channel: kontribusi orisinal penelitian ini.
    """
    if not TF_KERAS_VIS_OK:
        print("⚠️  tf-keras-vis tidak tersedia, skip Grad-CAM.")
        return None

    # Inisialisasi Grad-CAM
    gradcam = Gradcam(
        model_b4,
        model_modifier=ReplaceToLinear(),
        clone=True
    )

    x = X_sample[sample_idx:sample_idx+1]  # shape (1, 128, 128, 4)
    score = CategoricalScore([int(true_label[sample_idx])])

    # Hitung Grad-CAM → heatmap 2D (128, 128)
    cam = gradcam(score, x,
                  penultimate_layer='global_average_pooling2d')
    heatmap = cam[0]  # (128, 128)

    return heatmap


def visualize_gradcam_per_channel(model_b4, X_test, y_test,
                                   n_samples=5):
    """
    Visualisasi Grad-CAM per channel untuk beberapa sampel test.
    Untuk setiap sampel: tampilkan original channel + Grad-CAM overlay
    """
    if not TF_KERAS_VIS_OK:
        print("⚠️  tf-keras-vis tidak tersedia, skip visualisasi.")
        return

    os.makedirs(GRADCAM_DIR, exist_ok=True)

    # Pilih sampel yang diverse dari setiap kelas
    samples_per_class = max(1, n_samples // N_CLASSES)
    selected = []
    for cls in range(N_CLASSES):
        idx_cls = np.where(y_test == cls)[0]
        if len(idx_cls) > 0:
            chosen = idx_cls[:samples_per_class]
            selected.extend(chosen.tolist())

    print(f"\n📊 Grad-CAM per channel untuk {len(selected)} sampel...")

    for idx in selected[:n_samples]:
        emotion = EMOTION_LABELS[y_test[idx]]
        heatmap = compute_gradcam_per_channel(model_b4, X_test, y_test, idx)
        if heatmap is None:
            continue

        # Normalize heatmap ke [0,1]
        heatmap = (heatmap - heatmap.min()) / (heatmap.max() - heatmap.min() + 1e-8)

        # Plot: 4 channel + 4 overlay
        fig, axes = plt.subplots(2, 4, figsize=(20, 10))
        fig.suptitle(
            f'Per-Channel Grad-CAM | Sampel #{idx} | Emosi: {emotion}',
            fontsize=13, fontweight='bold'
        )

        for col, ch_name in enumerate(CHANNEL_ORDER):
            # Baris atas: gambar spektrogram original
            ch_img = X_test[idx, :, :, col]
            axes[0, col].imshow(ch_img, cmap='viridis',
                                aspect='auto', origin='lower')
            axes[0, col].set_title(f'{ch_name.upper()}\nSpectrogram')
            axes[0, col].axis('off')

            # Baris bawah: Grad-CAM overlay
            # Buat RGB dari spektrogram untuk overlay
            ch_rgb = plt.cm.viridis(ch_img)[:, :, :3]
            # Overlay heatmap
            heatmap_colored = plt.cm.jet(heatmap)[:, :, :3]
            overlay = (0.6 * ch_rgb + 0.4 * heatmap_colored)
            overlay = np.clip(overlay, 0, 1)

            axes[1, col].imshow(overlay, aspect='auto', origin='lower')
            axes[1, col].set_title(f'{ch_name.upper()}\nGrad-CAM Overlay')
            axes[1, col].axis('off')

        plt.tight_layout()
        save_path = os.path.join(GRADCAM_DIR, f"gradcam_{emotion}_{idx}.png")
        plt.savefig(save_path, dpi=100, bbox_inches='tight')
        plt.close()

    print(f"✅ Grad-CAM disimpan di folder: {GRADCAM_DIR}")


def analyze_channel_importance(model_b4, X_test, y_test):
    """
    Analisis statistik: rata-rata aktivasi Grad-CAM per channel
    untuk setiap kelas emosi.
    Menghasilkan tabel 8 emosi × 4 channel → heatmap 2D.
    """
    if not TF_KERAS_VIS_OK:
        return

    print("\n📊 Analisis Channel Importance per Emosi...")
    gradcam = Gradcam(
        model_b4,
        model_modifier=ReplaceToLinear(),
        clone=True
    )

    # importance[emosi, channel] = rata-rata saliency
    importance = np.zeros((N_CLASSES, N_CHANNELS))

    for cls in range(N_CLASSES):
        idx_cls = np.where(y_test == cls)[0][:10]  # Max 10 sampel per kelas
        if len(idx_cls) == 0:
            continue
        score = CategoricalScore([cls])
        for idx in idx_cls:
            x = X_test[idx:idx+1]
            cam = gradcam(score, x,
                          penultimate_layer='global_average_pooling2d')
            heatmap = cam[0]
            # Rata-rata aktivasi untuk tiap channel (berbobot gambar)
            for ch in range(N_CHANNELS):
                ch_weight = np.mean(X_test[idx, :, :, ch] * heatmap)
                importance[cls, ch] += ch_weight
        importance[cls] /= len(idx_cls)

    # Plot heatmap importance
    fig, ax = plt.subplots(figsize=(10, 8))
    sns.heatmap(
        importance,
        xticklabels=CHANNEL_ORDER,
        yticklabels=EMOTION_LABELS,
        cmap='YlOrRd', annot=True, fmt='.3f',
        ax=ax
    )
    ax.set_title('Channel Importance per Emosi\n(rata-rata Grad-CAM activation)',
                 fontsize=12, fontweight='bold')
    ax.set_xlabel('Spectrogram Channel')
    ax.set_ylabel('Emotion Class')
    plt.tight_layout()

    save_path = os.path.join(RESULT_DIR, "channel_importance_heatmap.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.show()
    print(f"✅ Channel importance heatmap disimpan ke {save_path}")

    return importance


# ── TABEL PERBANDINGAN ABLATION ───────────────────────────────

def plot_ablation_comparison(results_dict):
    """
    Plot bar chart perbandingan accuracy & F1 antar kondisi ablasi.
    """
    conditions = list(results_dict.keys())
    accuracies = [results_dict[c]['accuracy'] * 100 for c in conditions]
    f1_scores  = [results_dict[c]['macro_f1'] * 100 for c in conditions]
    labels_desc = {
        'B1': 'B1\nSingle-Mel\nNo DIP',
        'B2': 'B2\nSingle-Mel\nDIP',
        'B3': 'B3\n4-Channel\nNo DIP',
        'B4': 'B4\n4-Channel\nDIP\n(PROPOSED)',
    }

    x = np.arange(len(conditions))
    width = 0.35

    fig, ax = plt.subplots(figsize=(10, 7))
    bars1 = ax.bar(x - width/2, accuracies, width, label='Accuracy (%)',
                   color='steelblue', alpha=0.85)
    bars2 = ax.bar(x + width/2, f1_scores,  width, label='Macro F1 (%)',
                   color='darkorange', alpha=0.85)

    ax.set_xlabel('Kondisi Ablasi')
    ax.set_ylabel('Persentase (%)')
    ax.set_title('Ablation Study: Perbandingan Accuracy & Macro F1\n'
                 'Speech Emotion Recognition dengan MobileNetV2',
                 fontsize=12, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels([labels_desc.get(c, c) for c in conditions],
                        fontsize=9)
    ax.legend()
    ax.set_ylim(0, 110)
    ax.yaxis.grid(True, alpha=0.3)

    for bar in bars1:
        ax.annotate(f'{bar.get_height():.1f}%',
                    xy=(bar.get_x() + bar.get_width()/2, bar.get_height()),
                    xytext=(0, 3), textcoords='offset points',
                    ha='center', fontsize=9)
    for bar in bars2:
        ax.annotate(f'{bar.get_height():.1f}%',
                    xy=(bar.get_x() + bar.get_width()/2, bar.get_height()),
                    xytext=(0, 3), textcoords='offset points',
                    ha='center', fontsize=9)

    plt.tight_layout()
    save_path = os.path.join(RESULT_DIR, "ablation_comparison.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.show()
    print(f"✅ Grafik ablation disimpan ke {save_path}")


# ── MAIN ─────────────────────────────────────────────────────

if __name__ == "__main__":
    print("🚀 STEP 6: Evaluasi & Per-Channel Grad-CAM")
    print("="*55)

    os.makedirs(RESULT_DIR, exist_ok=True)
    os.makedirs(GRADCAM_DIR, exist_ok=True)

    # Load data
    X_train, y_train, X_val, y_val, X_test, y_test = load_dataset_npz()

    # Load raw data untuk B1/B3
    from step4_build_dataset import build_full_dataset, split_dataset
    X_raw, y_raw = build_full_dataset(enhanced_dir=SPEC_DIR)
    _, _, _, _, X_ts_raw, y_ts_raw = split_dataset(X_raw, y_raw)

    # Load semua model
    models = load_models()

    # Evaluasi semua kondisi
    results = {}
    test_data = {
        'B1': (X_ts_raw, y_ts_raw),
        'B2': (X_test,   y_test),
        'B3': (X_ts_raw, y_ts_raw),
        'B4': (X_test,   y_test),
    }
    for cond, desc in [
        ('B1', 'Single-Mel, No DIP'),
        ('B2', 'Single-Mel, DIP'),
        ('B3', '4-Channel, No DIP'),
        ('B4', '4-Channel, DIP (PROPOSED)'),
    ]:
        if cond not in models:
            print(f"⚠️  Skip {cond}: model tidak ditemukan.")
            continue
        X_ts, y_ts = test_data[cond]
        results[cond] = evaluate_model(models[cond], X_ts, y_ts, cond, desc)

    # Plot perbandingan ablation
    if results:
        plot_ablation_comparison(results)

    # Simpan hasil evaluasi
    eval_save = {}
    for cond, res in results.items():
        eval_save[cond] = {
            'accuracy': float(res['accuracy']),
            'macro_f1': float(res['macro_f1']),
            'loss': float(res['loss'])
        }
    eval_path = os.path.join(RESULT_DIR, "evaluation_results.json")
    with open(eval_path, 'w') as f:
        json.dump(eval_save, f, indent=2)
    print(f"\n✅ Hasil evaluasi disimpan ke {eval_path}")

    # Grad-CAM per channel (B4 = proposed method)
    if 'B4' in models:
        print("\n" + "="*55)
        print("🎯 Grad-CAM Per-Channel Analysis (Model B4 — PROPOSED)")
        print("="*55)
        visualize_gradcam_per_channel(models['B4'], X_test, y_test, n_samples=8)
        importance = analyze_channel_importance(models['B4'], X_test, y_test)
        if importance is not None:
            np.save(os.path.join(RESULT_DIR, "channel_importance.npy"), importance)

    print("\n" + "="*55)
    print("✅ Step 6 selesai! Pipeline penelitian COMPLETE.")
    print(f"   Semua hasil ada di: {RESULT_DIR}")
    print(f"   Semua Grad-CAM ada di: {GRADCAM_DIR}")
