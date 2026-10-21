"""
step5_train.py
==============
STEP KELIMA — training model MobileNetV2 dengan ablation study 4 kondisi.

Arsitektur model (untuk kondisi B3 dan B4 — 4-channel input):

  Input (128 × 128 × 4)
       ↓
  Conv2D(3 filter, kernel 1×1, padding='same')   ← channel projection
       ↓
  MobileNetV2 (pretrained ImageNet, input 128×128×3)
       ↓
  GlobalAveragePooling2D
       ↓
  Dropout(0.3)
       ↓
  Dense(128, relu) + BatchNormalization
       ↓
  Dropout(0.3)
       ↓
  Dense(8, softmax)   ← 8 kelas emosi RAVDESS

Kenapa Conv2D(1×1) sebagai channel projection?
MobileNetV2 pretrained ImageNet hanya bisa menerima input 3 channel (RGB).
Kita punya 4 channel. Solusi naif: duplikasi satu channel atau buang satu.
Solusi yang kita pakai: tambah lapisan Conv2D dengan kernel 1×1 yang belajar
bobot optimal untuk menggabungkan 4 channel menjadi 3. Kernel 1×1 artinya
operasi ini hanya di dimensi channel, tidak spatial — seperti weighted sum
antar channel. Bobot ini dilatih bersama sisa model, jadi model sendiri yang
belajar "mana channel yang paling informatif untuk setiap posisi spatial."

Training 2 fase:
  Phase 1 (FROZEN, lr=1e-4, 10 epoch):
    MobileNetV2 backbone dibekukan. Hanya Conv2D projection + head yang dilatih.
    Kenapa? Bobot head masih random dan gradiennya besar. Kalau backbone ikut
    dilatih dari awal, gradien besar ini bisa merusak bobot ImageNet yang
    sudah bagus.

  Phase 2 (UNFROZEN, lr=1e-5, max 100 epoch + early stopping):
    Semua layer dilatih. LR diturunkan 10× supaya tidak merusak bobot ImageNet.
    EarlyStopping patience=15 epoch, ReduceLROnPlateau patience=5 epoch.
"""

import os
import sys
import numpy as np
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

sys.path.append('/content/drive/MyDrive/SER_Research')
from config import (
    FUSED_DIR, SPEC_DIR, MODEL_DIR, RESULT_DIR,
    N_CLASSES, BATCH_SIZE,
    EPOCHS_FROZEN, EPOCHS_UNFROZEN,
    LEARNING_RATE, DROPOUT_RATE
)

import tensorflow as tf
from tensorflow.keras.applications import MobileNetV2
from tensorflow.keras.layers import (
    Input, Conv2D, GlobalAveragePooling2D,
    Dropout, Dense, BatchNormalization
)
from tensorflow.keras.models import Model
from tensorflow.keras.optimizers import Adam
from tensorflow.keras.callbacks import (
    EarlyStopping, ModelCheckpoint, ReduceLROnPlateau, CSVLogger
)
from tensorflow.keras.utils import to_categorical

print(f"TensorFlow versi: {tf.__version__}")
print(f"GPU tersedia: {tf.config.list_physical_devices('GPU')}")


# ── LOAD DATASET ──────────────────────────────────────────────────────────────

def load_dataset_npz():
    """
    Load dataset yang sudah difusion dan displit di step4.

    Kalau file tidak ada, kemungkinan step4 belum dijalankan
    atau path FUSED_DIR salah di config.py.
    """
    data_path = os.path.join(FUSED_DIR, "dataset.npz")
    if not os.path.exists(data_path):
        raise FileNotFoundError(
            f"File {data_path} tidak ditemukan. Jalankan step4 dulu."
        )
    data = np.load(data_path)
    X_train = data['X_train']
    y_train = data['y_train']
    X_val   = data['X_val']
    y_val   = data['y_val']
    X_test  = data['X_test']
    y_test  = data['y_test']
    print(f"Dataset dimuat: train={len(X_train)}, val={len(X_val)}, test={len(X_test)}")
    return X_train, y_train, X_val, y_val, X_test, y_test


# ── PERSIAPAN DATA PER KONDISI ABLASI ─────────────────────────────────────────

def prepare_ablation_data(X, condition):
    """
    Siapkan data sesuai kondisi ablasi.

    B1 dan B2: ambil hanya channel Mel (index 0), duplikasi 3× jadi 3 channel.
    Kenapa duplikasi? Karena MobileNetV2 butuh tepat 3 channel. Trik ini
    umum dipakai di transfer learning untuk grayscale → pretrained RGB model.

    B3 dan B4: pakai semua 4 channel, konversi 4→3 dilakukan di model via
    Conv2D(1×1) yang bisa dilatih.

    B1 = Mel tanpa DIP (datanya dari load_b1_b3_data)
    B2 = Mel dengan DIP (datanya dari load_dataset_npz)
    B3 = 4-channel tanpa DIP (datanya dari load_b1_b3_data)
    B4 = 4-channel dengan DIP (datanya dari load_dataset_npz) ← metode yang diusulkan
    """
    if condition in ('B1', 'B2'):
        # Ambil channel 0 (Mel), shape: (N, 128, 128, 1)
        mel_ch = X[:, :, :, 0:1]
        # Duplikasi jadi 3 channel: (N, 128, 128, 3)
        return np.concatenate([mel_ch, mel_ch, mel_ch], axis=-1)
    else:
        # B3, B4: kembalikan 4 channel apa adanya
        return X


def load_b1_b3_data():
    """
    B1 dan B3 butuh data tanpa DIP enhancement.
    Dataset utama (dataset.npz) sudah memakai gambar dari ENHANCED_DIR.
    Kita harus rebuild dataset dari SPEC_DIR (gambar asli sebelum enhancement).

    Import dari step4 supaya tidak duplikasi kode split.
    """
    from step4_build_dataset import build_full_dataset, split_dataset
    print("Membangun dataset B1/B3 dari SPEC_DIR (tanpa DIP)...")
    X_raw, y_raw = build_full_dataset(enhanced_dir=SPEC_DIR)
    return split_dataset(X_raw, y_raw)


# ── ARSITEKTUR MODEL ──────────────────────────────────────────────────────────

def build_model_4ch(n_classes=N_CLASSES):
    """
    Model untuk 4-channel input (kondisi B3 dan B4).

    Lapisan Conv2D(3, 1×1) adalah channel projection:
      Input:  (batch, 128, 128, 4) → 4 channel spektrogram
      Output: (batch, 128, 128, 3) → 3 channel untuk MobileNetV2

    Mengapa kernel 1×1?
    Kernel 1×1 = operasi yang hanya melihat satu piksel pada satu waktu.
    Artinya ia tidak menangkap pola spatial, hanya melakukan weighted
    combination antar channel. Ini persis yang kita mau: kita tidak
    ingin layer ini belajar pola spatial (itu tugas MobileNetV2),
    hanya belajar bagaimana mengkombinasikan 4 channel menjadi 3.

    Mengembalikan (model, base_model) karena base_model perlu
    di-unfreeze terpisah di phase 2.
    """
    inp = Input(shape=(128, 128, 4), name='input_4ch')

    # Channel projection: 4 channel → 3 channel (dapat dilatih)
    x = Conv2D(3, (1, 1), padding='same', name='channel_projection')(inp)

    # MobileNetV2 backbone
    # include_top=False → hapus layer klasifikasi bawaan (kita ganti sendiri)
    # weights='imagenet'  → pakai bobot pretrained, bukan dari nol
    base = MobileNetV2(
        input_shape=(128, 128, 3),
        include_top=False,
        weights='imagenet'
    )
    base.trainable = False   # Frozen di phase 1

    # training=False penting: BatchNorm di MobileNetV2 menggunakan
    # running statistics (bukan batch statistics) saat backbone frozen.
    # Kalau tidak diset, BatchNorm bisa merusak statistik yang sudah dilatih.
    x = base(x, training=False)

    x = GlobalAveragePooling2D(name='global_avg_pool')(x)
    x = Dropout(DROPOUT_RATE)(x)
    x = Dense(128, activation='relu', name='fc1')(x)
    x = BatchNormalization()(x)
    x = Dropout(DROPOUT_RATE)(x)
    out = Dense(n_classes, activation='softmax', name='output')(x)

    model = Model(inp, out, name='SER_MobileNetV2_4ch')
    return model, base


def build_model_3ch(n_classes=N_CLASSES):
    """
    Model untuk single-channel Mel input (kondisi B1 dan B2).

    Sama seperti build_model_4ch tapi tanpa channel projection.
    Input sudah 3 channel (duplikasi dari Mel).
    """
    base = MobileNetV2(
        input_shape=(128, 128, 3),
        include_top=False,
        weights='imagenet'
    )
    base.trainable = False

    inp = Input(shape=(128, 128, 3), name='input_3ch')
    x = base(inp, training=False)
    x = GlobalAveragePooling2D(name='global_avg_pool')(x)
    x = Dropout(DROPOUT_RATE)(x)
    x = Dense(128, activation='relu', name='fc1')(x)
    x = BatchNormalization()(x)
    x = Dropout(DROPOUT_RATE)(x)
    out = Dense(n_classes, activation='softmax', name='output')(x)

    model = Model(inp, out, name='SER_MobileNetV2_3ch')
    return model, base


# ── TRAINING 2 FASE ───────────────────────────────────────────────────────────

def train_model(model, base_model, X_tr, y_tr, X_vl, y_vl,
                model_name, model_dir):
    """
    Latih model dengan strategi 2 fase:

    Phase 1 — FROZEN backbone:
      Backbone MobileNetV2 dibekukan, hanya layer baru yang dilatih.
      Tujuan: stabilkan bobot layer baru sebelum fine-tuning.
      10 epoch sudah cukup karena kita hanya melatih head (~200K parameter).

    Phase 2 — UNFROZEN (fine-tuning):
      Seluruh model dilatih dengan LR 10× lebih kecil.
      LR kecil (1e-5) penting: kita tidak mau "mengacak" bobot ImageNet
      yang sudah bagus, hanya menyesuaikannya ke domain spektrogram.

    Callback yang dipakai:
      EarlyStopping patience=15  → berhenti kalau val_loss tidak turun 15 epoch
      ReduceLROnPlateau patience=5→ kurangi LR 50% kalau val_loss stagnasi 5 epoch
      ModelCheckpoint            → simpan model terbaik (val_accuracy tertinggi)
      CSVLogger                  → rekam history ke CSV untuk analisis
    """
    os.makedirs(model_dir, exist_ok=True)

    # One-hot encoding label
    y_tr_oh = to_categorical(y_tr, N_CLASSES)
    y_vl_oh = to_categorical(y_vl, N_CLASSES)

    # ── Phase 1: Backbone frozen ─────────────────────────────────
    print(f"\n[{model_name}] Phase 1: Training head (backbone frozen)...")
    model.compile(
        optimizer=Adam(LEARNING_RATE),
        loss='categorical_crossentropy',
        metrics=['accuracy']
    )

    callbacks_p1 = [
        CSVLogger(os.path.join(model_dir, f"{model_name}_phase1.csv")),
        ModelCheckpoint(
            os.path.join(model_dir, f"{model_name}_best.keras"),
            monitor='val_accuracy',
            save_best_only=True,
            verbose=0
        ),
    ]
    hist1 = model.fit(
        X_tr, y_tr_oh,
        validation_data=(X_vl, y_vl_oh),
        batch_size=BATCH_SIZE,
        epochs=EPOCHS_FROZEN,
        callbacks=callbacks_p1,
        verbose=1
    )

    # ── Phase 2: Unfreeze semua layer ────────────────────────────
    print(f"\n[{model_name}] Phase 2: Fine-tuning semua layer...")
    base_model.trainable = True

    # Setelah unfreeze, perlu compile ulang supaya perubahan trainable
    # terdaftar di optimizer
    model.compile(
        optimizer=Adam(LEARNING_RATE / 10),   # LR dikecilkan 10×
        loss='categorical_crossentropy',
        metrics=['accuracy']
    )

    callbacks_p2 = [
        EarlyStopping(
            monitor='val_loss', patience=15,
            restore_best_weights=True, verbose=1
        ),
        ReduceLROnPlateau(
            monitor='val_loss', factor=0.5,
            patience=5, min_lr=1e-7, verbose=1
        ),
        ModelCheckpoint(
            os.path.join(model_dir, f"{model_name}_best.keras"),
            monitor='val_accuracy',
            save_best_only=True, verbose=0
        ),
        CSVLogger(os.path.join(model_dir, f"{model_name}_phase2.csv")),
    ]
    hist2 = model.fit(
        X_tr, y_tr_oh,
        validation_data=(X_vl, y_vl_oh),
        batch_size=BATCH_SIZE,
        epochs=EPOCHS_UNFROZEN,
        callbacks=callbacks_p2,
        verbose=1
    )

    return hist1, hist2


# ── VISUALISASI KURVA TRAINING ────────────────────────────────────────────────

def plot_training_history(hist1, hist2, model_name):
    """
    Plot kurva accuracy dan loss dari phase 1 dan phase 2 secara bersamaan.

    Garis vertikal abu-abu menandai batas antara phase 1 (frozen) dan
    phase 2 (unfrozen). Biasanya ada "lompatan" performa tepat setelah
    garis ini karena backbone mulai bisa menyesuaikan diri.
    """
    acc      = hist1.history['accuracy']     + hist2.history['accuracy']
    val_acc  = hist1.history['val_accuracy'] + hist2.history['val_accuracy']
    loss     = hist1.history['loss']         + hist2.history['loss']
    val_loss = hist1.history['val_loss']     + hist2.history['val_loss']
    epochs   = range(1, len(acc) + 1)
    phase1_end = len(hist1.history['accuracy'])

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
    fig.suptitle(f'Training History — {model_name}', fontsize=13)

    for ax, (train_vals, val_vals, metric) in zip([ax1, ax2], [
        (acc, val_acc, 'Accuracy'),
        (loss, val_loss, 'Loss')
    ]):
        ax.plot(epochs, train_vals, label=f'Train {metric}')
        ax.plot(epochs, val_vals,   label=f'Val {metric}')
        ax.axvline(
            x=phase1_end, color='gray', linestyle='--',
            label=f'Phase 2 start (epoch {phase1_end + 1})'
        )
        ax.set_xlabel('Epoch')
        ax.set_ylabel(metric)
        ax.legend()
        ax.set_title(metric)

    plt.tight_layout()
    out_path = os.path.join(RESULT_DIR, f"training_{model_name}.png")
    plt.savefig(out_path, dpi=100, bbox_inches='tight')
    plt.close()
    print(f"Kurva training disimpan ke {out_path}")


# ── MAIN: ABLATION STUDY ──────────────────────────────────────────────────────

if __name__ == "__main__":
    print("STEP 5: Training MobileNetV2 — Ablation Study B1 sampai B4")
    print("=" * 60)

    os.makedirs(MODEL_DIR, exist_ok=True)
    os.makedirs(RESULT_DIR, exist_ok=True)

    # Load dataset dengan DIP (untuk B2 dan B4)
    X_train, y_train, X_val, y_val, X_test, y_test = load_dataset_npz()

    # Load dataset tanpa DIP (untuk B1 dan B3)
    X_tr_raw, y_tr_raw, X_vl_raw, y_vl_raw, X_ts_raw, y_ts_raw = load_b1_b3_data()

    # Dict untuk menyimpan hasil per kondisi
    ablation_results = {}

    # ── B1: Single Mel, Tanpa DIP ─────────────────────────────────
    print("\n" + "=" * 60)
    print("B1: Single-channel Mel, TANPA DIP — baseline sederhana")
    X_tr_b1 = prepare_ablation_data(X_tr_raw, 'B1')
    X_vl_b1 = prepare_ablation_data(X_vl_raw, 'B1')
    model_b1, base_b1 = build_model_3ch()
    model_b1.summary()
    h1a, h1b = train_model(
        model_b1, base_b1,
        X_tr_b1, y_tr_raw,
        X_vl_b1, y_vl_raw,
        'B1_mel_nodip', MODEL_DIR
    )
    plot_training_history(h1a, h1b, 'B1_mel_nodip')
    ablation_results['B1'] = {
        'model':  model_b1,
        'X_test': prepare_ablation_data(X_ts_raw, 'B1'),
        'y_test': y_ts_raw
    }

    # ── B2: Single Mel, Dengan DIP ────────────────────────────────
    print("\n" + "=" * 60)
    print("B2: Single-channel Mel, DENGAN DIP")
    X_tr_b2 = prepare_ablation_data(X_train, 'B2')
    X_vl_b2 = prepare_ablation_data(X_val,   'B2')
    model_b2, base_b2 = build_model_3ch()
    h2a, h2b = train_model(
        model_b2, base_b2,
        X_tr_b2, y_train,
        X_vl_b2, y_val,
        'B2_mel_dip', MODEL_DIR
    )
    plot_training_history(h2a, h2b, 'B2_mel_dip')
    ablation_results['B2'] = {
        'model':  model_b2,
        'X_test': prepare_ablation_data(X_test, 'B2'),
        'y_test': y_test
    }

    # ── B3: 4-Channel, Tanpa DIP ──────────────────────────────────
    print("\n" + "=" * 60)
    print("B3: 4-channel, TANPA DIP")
    model_b3, base_b3 = build_model_4ch()
    h3a, h3b = train_model(
        model_b3, base_b3,
        X_tr_raw, y_tr_raw,
        X_vl_raw, y_vl_raw,
        'B3_4ch_nodip', MODEL_DIR
    )
    plot_training_history(h3a, h3b, 'B3_4ch_nodip')
    ablation_results['B3'] = {
        'model':  model_b3,
        'X_test': X_ts_raw,
        'y_test': y_ts_raw
    }

    # ── B4: 4-Channel, Dengan DIP (METODE YANG DIUSULKAN) ─────────
    print("\n" + "=" * 60)
    print("B4: 4-channel, DENGAN DIP — METODE YANG DIUSULKAN")
    model_b4, base_b4 = build_model_4ch()
    h4a, h4b = train_model(
        model_b4, base_b4,
        X_train, y_train,
        X_val, y_val,
        'B4_4ch_dip', MODEL_DIR
    )
    plot_training_history(h4a, h4b, 'B4_4ch_dip')
    ablation_results['B4'] = {
        'model':  model_b4,
        'X_test': X_test,
        'y_test': y_test
    }

    # ── Ringkasan Akurasi Test ────────────────────────────────────
    print("\n" + "=" * 60)
    print("RINGKASAN ABLATION STUDY — Akurasi pada Test Set")
    print(f"{'Kondisi':<12} {'Test Accuracy':>14}")
    print("-" * 28)

    summary = {}
    for cond, result in ablation_results.items():
        mdl   = result['model']
        X_ts  = result['X_test']
        y_ts  = result['y_test']
        y_ts_oh = to_categorical(y_ts, N_CLASSES)
        _, acc = mdl.evaluate(X_ts, y_ts_oh, verbose=0)
        summary[cond] = round(float(acc), 4)
        print(f"{cond:<12} {acc * 100:>13.2f}%")

    summary_path = os.path.join(RESULT_DIR, "ablation_summary.json")
    with open(summary_path, 'w') as f:
        json.dump(summary, f, indent=2)
    print(f"\nRingkasan disimpan ke {summary_path}")
    print("\nStep 5 selesai. Lanjut ke step6_evaluate_gradcam.py")
