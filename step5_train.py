# ============================================================
# step5_train.py — Training MobileNetV2 + Ablation Study
# ============================================================
# Jalankan setelah step4_build_dataset.py
#
# Arsitektur:
#   Input (128×128×4)
#   → Conv2D(3, 1×1, padding='same')   ← channel projection 4→3
#   → MobileNetV2 (pretrained ImageNet, frozen saat phase 1)
#   → GlobalAveragePooling2D
#   → Dropout(0.3)
#   → Dense(128, relu)
#   → Dropout(0.3)
#   → Dense(8, softmax)   ← 8 kelas emosi
#
# Ablation Study (4 kondisi):
#   B1: Single-channel Mel, tanpa DIP (baseline sederhana)
#   B2: Single-channel Mel, dengan DIP  (test DIP saja)
#   B3: 4-channel fused, tanpa DIP      (test multi-channel saja)
#   B4: 4-channel fused, dengan DIP     (PROPOSED METHOD)
# ============================================================

import os
import sys
import numpy as np
import glob
import cv2
import json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns

sys.path.append('/content/drive/MyDrive/SER_Research')
from config import *

# TensorFlow / Keras
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

print(f"TF version: {tf.__version__}")
print(f"GPU: {tf.config.list_physical_devices('GPU')}")


# ── LOAD DATASET ─────────────────────────────────────────────

def load_dataset_npz():
    """Load dataset dari file .npz yang dibuat step4."""
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
    print(f"✅ Dataset dimuat: train={len(X_train)}, val={len(X_val)}, test={len(X_test)}")
    return X_train, y_train, X_val, y_val, X_test, y_test


# ── PERSIAPAN DATA PER KONDISI ABLASI ────────────────────────

def prepare_ablation_data(X, condition):
    """
    Siapkan data sesuai kondisi ablasi.
    condition: 'B1', 'B2', 'B3', 'B4'

    B1: Hanya channel 0 (Mel asli) → RGB palsu dengan tile 3 channel
    B2: Hanya channel 0 (Mel+DIP)  → sama, tapi dari enhanced
    B3: 4-channel, tanpa DIP        → pakai data dari SPEC_DIR
    B4: 4-channel, dengan DIP       → data X dari step4 (ENHANCED)
    """
    if condition in ('B1', 'B2'):
        # Ambil hanya channel Mel (index 0), duplikasi jadi 3 channel
        mel_ch = X[:, :, :, 0:1]  # (N, 128, 128, 1)
        return np.concatenate([mel_ch, mel_ch, mel_ch], axis=-1)  # (N, 128, 128, 3)
    else:
        # B3 dan B4 pakai 4 channel, channel projection dilakukan di model
        return X  # (N, 128, 128, 4)


def load_b1_b3_data():
    """
    B1 & B3 butuh data tanpa DIP (dari SPEC_DIR, sebelum enhancement).
    Fungsi ini membangun ulang dataset dari SPEC_DIR.
    """
    from step4_build_dataset import build_full_dataset, split_dataset
    print("⚙️  Membangun dataset B1/B3 dari SPEC_DIR (tanpa DIP)...")
    X_raw, y_raw = build_full_dataset(enhanced_dir=SPEC_DIR)
    return split_dataset(X_raw, y_raw)


# ── ARSITEKTUR MODEL ──────────────────────────────────────────

def build_model_4ch(n_classes=N_CLASSES):
    """
    Model untuk 4-channel input (B3, B4).
    Conv2D(3, 1×1) memproyeksikan 4 channel → 3 channel
    sebelum masuk MobileNetV2 yang pretrained pada 3-channel ImageNet.
    """
    inp = Input(shape=(128, 128, 4), name='input_4ch')

    # Channel projection: 4 → 3 (belajar bobot channel terbaik)
    x = Conv2D(3, (1, 1), padding='same', name='channel_proj')(inp)

    # MobileNetV2 backbone (pretrained ImageNet, tanpa head)
    base = MobileNetV2(
        input_shape=(128, 128, 3),
        include_top=False,
        weights='imagenet'
    )
    base.trainable = False  # Frozen di phase 1

    x = base(x, training=False)
    x = GlobalAveragePooling2D(name='gap')(x)
    x = Dropout(DROPOUT_RATE)(x)
    x = Dense(128, activation='relu', name='fc1')(x)
    x = BatchNormalization()(x)
    x = Dropout(DROPOUT_RATE)(x)
    out = Dense(n_classes, activation='softmax', name='output')(x)

    model = Model(inp, out, name='SER_MobileNetV2_4ch')
    return model, base


def build_model_3ch(n_classes=N_CLASSES):
    """
    Model untuk single-channel Mel input (B1, B2).
    Input: 3 channel (duplicate dari 1 channel Mel).
    """
    base = MobileNetV2(
        input_shape=(128, 128, 3),
        include_top=False,
        weights='imagenet'
    )
    base.trainable = False

    inp = Input(shape=(128, 128, 3), name='input_3ch')
    x = base(inp, training=False)
    x = GlobalAveragePooling2D(name='gap')(x)
    x = Dropout(DROPOUT_RATE)(x)
    x = Dense(128, activation='relu', name='fc1')(x)
    x = BatchNormalization()(x)
    x = Dropout(DROPOUT_RATE)(x)
    out = Dense(n_classes, activation='softmax', name='output')(x)

    model = Model(inp, out, name='SER_MobileNetV2_3ch')
    return model, base


# ── TRAINING 2-PHASE ─────────────────────────────────────────

def train_model(model, base_model, X_tr, y_tr, X_vl, y_vl,
                model_name, model_dir):
    """
    Training 2 fase:
    Phase 1 (FROZEN, 10 epoch):
        Hanya layer baru yang dilatih, backbone dibekukan.
        Tujuan: inisialisasi bobot layer baru yang stabil.
    Phase 2 (UNFROZEN, max 100 epoch + early stopping):
        Seluruh model dilatih dengan LR kecil (fine-tuning).
        Tujuan: adaptasi backbone ke domain spektrogram emosi.
    """
    os.makedirs(model_dir, exist_ok=True)

    # One-hot encoding
    y_tr_oh = to_categorical(y_tr, N_CLASSES)
    y_vl_oh = to_categorical(y_vl, N_CLASSES)

    # ── Phase 1: Frozen backbone ──
    print(f"\n⚙️  [{model_name}] Phase 1: Training layer baru (backbone frozen)...")
    model.compile(
        optimizer=Adam(LEARNING_RATE),
        loss='categorical_crossentropy',
        metrics=['accuracy']
    )

    cb_p1 = [
        CSVLogger(os.path.join(model_dir, f"{model_name}_phase1.csv")),
        ModelCheckpoint(
            os.path.join(model_dir, f"{model_name}_best.keras"),
            monitor='val_accuracy', save_best_only=True, verbose=0
        ),
    ]
    hist1 = model.fit(
        X_tr, y_tr_oh,
        validation_data=(X_vl, y_vl_oh),
        batch_size=BATCH_SIZE,
        epochs=EPOCHS_FROZEN,
        callbacks=cb_p1,
        verbose=1
    )

    # ── Phase 2: Unfreeze semua layer ──
    print(f"\n⚙️  [{model_name}] Phase 2: Fine-tuning seluruh model...")
    base_model.trainable = True
    model.compile(
        optimizer=Adam(LEARNING_RATE / 10),
        loss='categorical_crossentropy',
        metrics=['accuracy']
    )

    cb_p2 = [
        EarlyStopping(monitor='val_loss', patience=15,
                      restore_best_weights=True, verbose=1),
        ReduceLROnPlateau(monitor='val_loss', factor=0.5,
                          patience=5, min_lr=1e-7, verbose=1),
        ModelCheckpoint(
            os.path.join(model_dir, f"{model_name}_best.keras"),
            monitor='val_accuracy', save_best_only=True, verbose=0
        ),
        CSVLogger(os.path.join(model_dir, f"{model_name}_phase2.csv")),
    ]
    hist2 = model.fit(
        X_tr, y_tr_oh,
        validation_data=(X_vl, y_vl_oh),
        batch_size=BATCH_SIZE,
        epochs=EPOCHS_UNFROZEN,
        callbacks=cb_p2,
        verbose=1
    )

    return hist1, hist2


# ── PLOT KURVA TRAINING ───────────────────────────────────────

def plot_training_history(hist1, hist2, model_name):
    """
    Plot gabungan kurva accuracy & loss dari phase 1 dan phase 2.
    """
    acc  = hist1.history['accuracy']  + hist2.history['accuracy']
    val_acc = hist1.history['val_accuracy'] + hist2.history['val_accuracy']
    loss = hist1.history['loss']      + hist2.history['loss']
    val_loss = hist1.history['val_loss']    + hist2.history['val_loss']
    epochs = range(1, len(acc) + 1)
    phase1_end = len(hist1.history['accuracy'])

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
    fig.suptitle(f'Training History — {model_name}', fontsize=13)

    for ax, (train, val, label) in zip([ax1, ax2], [
        (acc, val_acc, 'Accuracy'),
        (loss, val_loss, 'Loss')
    ]):
        ax.plot(epochs, train, label=f'Train {label}')
        ax.plot(epochs, val, label=f'Val {label}')
        ax.axvline(x=phase1_end, color='gray', linestyle='--',
                   label=f'Phase 2 start (ep {phase1_end+1})')
        ax.set_xlabel('Epoch')
        ax.set_ylabel(label)
        ax.legend()
        ax.set_title(label)

    plt.tight_layout()
    save_path = os.path.join(RESULT_DIR, f"training_{model_name}.png")
    plt.savefig(save_path, dpi=100, bbox_inches='tight')
    plt.show()
    print(f"✅ Kurva training disimpan ke {save_path}")


# ── MAIN: ABLATION STUDY ─────────────────────────────────────

if __name__ == "__main__":
    print("🚀 STEP 5: Training MobileNetV2 — Ablation Study (B1–B4)")
    print("="*60)

    os.makedirs(MODEL_DIR, exist_ok=True)
    os.makedirs(RESULT_DIR, exist_ok=True)

    # Load dataset dengan DIP (B2, B4)
    X_train, y_train, X_val, y_val, X_test, y_test = load_dataset_npz()

    # Load dataset tanpa DIP (B1, B3)
    X_tr_raw, y_tr_raw, X_vl_raw, y_vl_raw, X_ts_raw, y_ts_raw = load_b1_b3_data()

    ablation_results = {}

    # ── B1: Single Mel, No DIP ────────────────────────────────
    print("\n" + "="*60)
    print("🔬 B1: Single-channel Mel (tanpa DIP) — BASELINE")
    X_tr_b1 = prepare_ablation_data(X_tr_raw, 'B1')
    X_vl_b1 = prepare_ablation_data(X_vl_raw, 'B1')
    model_b1, base_b1 = build_model_3ch()
    model_b1.summary()
    h1a, h1b = train_model(model_b1, base_b1,
                            X_tr_b1, y_tr_raw, X_vl_b1, y_vl_raw,
                            'B1_mel_nodip', MODEL_DIR)
    plot_training_history(h1a, h1b, 'B1_mel_nodip')
    ablation_results['B1'] = {'model': model_b1, 'X_test': prepare_ablation_data(X_ts_raw, 'B1'), 'y_test': y_ts_raw}

    # ── B2: Single Mel, With DIP ──────────────────────────────
    print("\n" + "="*60)
    print("🔬 B2: Single-channel Mel (dengan DIP)")
    X_tr_b2 = prepare_ablation_data(X_train, 'B2')
    X_vl_b2 = prepare_ablation_data(X_val,   'B2')
    model_b2, base_b2 = build_model_3ch()
    h2a, h2b = train_model(model_b2, base_b2,
                            X_tr_b2, y_train, X_vl_b2, y_val,
                            'B2_mel_dip', MODEL_DIR)
    plot_training_history(h2a, h2b, 'B2_mel_dip')
    ablation_results['B2'] = {'model': model_b2, 'X_test': prepare_ablation_data(X_test, 'B2'), 'y_test': y_test}

    # ── B3: 4-Channel, No DIP ─────────────────────────────────
    print("\n" + "="*60)
    print("🔬 B3: 4-channel (tanpa DIP)")
    model_b3, base_b3 = build_model_4ch()
    h3a, h3b = train_model(model_b3, base_b3,
                            X_tr_raw, y_tr_raw, X_vl_raw, y_vl_raw,
                            'B3_4ch_nodip', MODEL_DIR)
    plot_training_history(h3a, h3b, 'B3_4ch_nodip')
    ablation_results['B3'] = {'model': model_b3, 'X_test': X_ts_raw, 'y_test': y_ts_raw}

    # ── B4: 4-Channel, With DIP (PROPOSED) ───────────────────
    print("\n" + "="*60)
    print("🔬 B4: 4-channel (dengan DIP) ← PROPOSED METHOD")
    model_b4, base_b4 = build_model_4ch()
    h4a, h4b = train_model(model_b4, base_b4,
                            X_train, y_train, X_val, y_val,
                            'B4_4ch_dip', MODEL_DIR)
    plot_training_history(h4a, h4b, 'B4_4ch_dip')
    ablation_results['B4'] = {'model': model_b4, 'X_test': X_test, 'y_test': y_test}

    # ── Ringkasan Akurasi Test ─────────────────────────────────
    print("\n" + "="*60)
    print("📊 RINGKASAN ABLATION STUDY")
    print(f"{'Kondisi':<10} {'Test Accuracy':>14}")
    print("-"*26)
    summary = {}
    for cond, result in ablation_results.items():
        model = result['model']
        X_ts  = result['X_test']
        y_ts  = result['y_test']
        y_ts_oh = to_categorical(y_ts, N_CLASSES)
        _, acc = model.evaluate(X_ts, y_ts_oh, verbose=0)
        summary[cond] = acc
        print(f"{cond:<10} {acc*100:>13.2f}%")

    # Simpan ringkasan
    summary_path = os.path.join(RESULT_DIR, "ablation_summary.json")
    with open(summary_path, 'w') as f:
        json.dump(summary, f, indent=2)
    print(f"\n✅ Ringkasan disimpan ke {summary_path}")
    print("\n✅ Step 5 selesai! Lanjut ke step6_evaluate_gradcam.py")
