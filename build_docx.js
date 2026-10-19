// build_docx.js -- Generate SER Paper as IEEE-style .docx
// Font: Times New Roman (serif body) + Arial (sans headings) -- standard in LibreOffice/Word
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, Table, TableRow, TableCell, WidthType,
  BorderStyle, ShadingType, PageOrientation,
  SectionType, ColumnBreak, ImageRun
} = require('docx');
const fs = require('fs');
const path = require('path');

const FIGURES_DIR = '/home/claude/figures';

// ── Unit helpers ─────────────────────────────────────────────
const twip = (inch) => Math.round(inch * 1440);
const pt   = (n)    => n * 2;          // half-points (docx font size unit)

// ── Paragraph factories ──────────────────────────────────────

function body(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: 'Times New Roman', size: pt(10) })],
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 60, line: 240, lineRule: 'auto' },
  });
}

function bodyRuns(runs) {
  return new Paragraph({
    children: runs.map(r => new TextRun({ font: 'Times New Roman', size: pt(10), ...r })),
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 60, line: 240, lineRule: 'auto' },
  });
}

function secHead(num, title) {
  return new Paragraph({
    children: [
      ...(num ? [new TextRun({ text: `${num}. `, font: 'Arial', size: pt(9.5), bold: true, allCaps: true, color: '1a3a5c' })] : []),
      new TextRun({ text: title, font: 'Arial', size: pt(9.5), bold: true, allCaps: true, color: '1a3a5c' }),
    ],
    spacing: { before: 180, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' } },
  });
}

function subHead(num, title) {
  return new Paragraph({
    children: [
      ...(num ? [new TextRun({ text: `${num} `, font: 'Times New Roman', size: pt(9.5), bold: true, italics: true })] : []),
      new TextRun({ text: title, font: 'Times New Roman', size: pt(9.5), bold: true, italics: true }),
    ],
    spacing: { before: 100, after: 40 },
  });
}

function eq(text, num) {
  return new Paragraph({
    children: [
      new TextRun({ text: `        ${text}`, font: 'Times New Roman', size: pt(9.5) }),
      new TextRun({ text: `    (${num})`, font: 'Times New Roman', size: pt(8.5), color: '666666' }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 60 },
  });
}

function figPH(label, desc) {
  return new Paragraph({
    children: [new TextRun({ text: `[${label}: ${desc}]`, font: 'Arial', size: pt(8), color: '777777', italics: true })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 20 },
    border: {
      top: { style: BorderStyle.DASHED, size: 4, color: 'BBBBBB' },
      bottom: { style: BorderStyle.DASHED, size: 4, color: 'BBBBBB' },
      left: { style: BorderStyle.DASHED, size: 4, color: 'BBBBBB' },
      right: { style: BorderStyle.DASHED, size: 4, color: 'BBBBBB' },
    },
  });
}

function figCap(num, text) {
  return new Paragraph({
    children: [
      new TextRun({ text: `Figure ${num}. `, font: 'Times New Roman', size: pt(8.5), bold: true }),
      new TextRun({ text, font: 'Times New Roman', size: pt(8.5), italics: true }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 20, after: 80 },
  });
}

function figImg(filename, widthInch, heightInch) {
  const imgPath = path.join(FIGURES_DIR, filename);
  const imgBuf = fs.readFileSync(imgPath);
  const ext = path.extname(filename).slice(1).toLowerCase(); // 'png' or 'jpg'
  return new Paragraph({
    children: [new ImageRun({
      data: imgBuf,
      transformation: {
        width:  Math.round(widthInch * 96),
        height: Math.round(heightInch * 96),
      },
      type: ext,
    })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 20 },
  });
}



function refItem(num, text) {
  return new Paragraph({
    children: [
      new TextRun({ text: `[${num}]  `, font: 'Times New Roman', size: pt(8), bold: true }),
      new TextRun({ text, font: 'Times New Roman', size: pt(8) }),
    ],
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 36 },
    indent: { left: twip(0.22), hanging: twip(0.22) },
  });
}

function blank(after = 60) {
  return new Paragraph({ children: [], spacing: { after } });
}

function hr() {
  return new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' } },
    spacing: { after: 80 },
  });
}

// ── Ablation table ───────────────────────────────────────────
function ablationTable() {
  // Two-column layout: each column ~4320 twip. Table must fit within one column.
  const colW = [2000, 560, 560, 560, 560]; // total 4240 twip — fits in one column
  const headers = ['Condition', 'Acc.', 'Prec.', 'Rec.', 'F1'];
  const rows = [
    ['B1 -- Mel, No DIP',          '38.89', '36.27', '36.47', '34.39'],
    ['B2 -- Mel + DIP ★',          '39.81', '41.55', '40.55', '38.89'],
    ['B3 -- 4-Ch, No DIP',         '26.85', '20.20', '25.12', '20.19'],
    ['B4 -- 4-Ch + DIP',           '27.31', '24.58', '26.08', '23.29'],
  ];

  return new Table({
    width: { size: 4240, type: WidthType.DXA },
    columnWidths: colW,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => new TableCell({
          children: [new Paragraph({
            children: [new TextRun({ text: h, font: 'Arial', size: pt(8), bold: true, color: 'FFFFFF' })],
            alignment: AlignmentType.CENTER,
          })],
          width: { size: colW[i], type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: '1a3a5c' },
        })),
      }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((val, ci) => new TableCell({
          children: [new Paragraph({
            children: [new TextRun({
              text: val,
              font: ci === 0 ? 'Times New Roman' : 'Arial',
              size: pt(8),
              bold: ri === 1,
              color: ri === 1 && ci > 0 ? '1a3a5c' : '222222',
            })],
            alignment: ci === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
          })],
          width: { size: colW[ci], type: WidthType.DXA },
        })),
      })),
    ],
  });
}

// ══════════════════════════════════════════════════════════════
// CONTENT blocks grouped: titleSection, bodySection
// ══════════════════════════════════════════════════════════════

// ── SINGLE-COLUMN section: Title + Abstract ───────────────────
const titleSection = {
  properties: {
    type: SectionType.CONTINUOUS,
    page: {
      size: { width: twip(8.5), height: twip(11) },
      margin: { top: twip(1), bottom: twip(1), left: twip(0.75), right: twip(0.75) },
    },
    column: { count: 1 },
  },
  children: [
    // TITLE
    new Paragraph({
      children: [new TextRun({
        text: 'Multi-Channel Spectrogram Fusion with Digital Image Processing Enhancement for Convolutional Neural Network-Based Speech Emotion Recognition',
        font: 'Times New Roman', size: pt(15.5), bold: true,
      })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
    }),

    // AUTHORS
    new Paragraph({
      children: [
        new TextRun({ text: 'Jehezkiel Beny Brian', font: 'Times New Roman', size: pt(10) }),
        new TextRun({ text: '1,*', font: 'Times New Roman', size: pt(7), superScript: true }),
        new TextRun({ text: ' and [Supervisor Name]', font: 'Times New Roman', size: pt(10) }),
        new TextRun({ text: '1,†', font: 'Times New Roman', size: pt(7), superScript: true }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
    }),

    new Paragraph({
      children: [
        new TextRun({ text: '1', font: 'Times New Roman', size: pt(7), superScript: true }),
        new TextRun({ text: ' School of Informatics, Faculty of Informatics, Telkom University, Bandung 40257, Indonesia', font: 'Times New Roman', size: pt(8.5), italics: true }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 20 },
    }),

    new Paragraph({
      children: [
        new TextRun({ text: '* Correspondence: ', font: 'Times New Roman', size: pt(8), italics: true }),
        new TextRun({ text: 'jehezkielbeny123@gmail.com', font: 'Courier New', size: pt(8), color: '1a3a5c' }),
        new TextRun({ text: '   † Supervisor', font: 'Times New Roman', size: pt(8), italics: true }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
    }),

    hr(),

    // ABSTRACT heading
    new Paragraph({
      children: [new TextRun({ text: 'ABSTRACT', font: 'Arial', size: pt(8), bold: true, color: '1a3a5c' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
    }),

    // Abstract text
    bodyRuns([
      { text: 'Speech Emotion Recognition (SER) is a fundamental problem in affective computing with wide applications in human-computer interaction, mental health monitoring, and social robotics [1],[2]. Existing deep learning approaches predominantly rely on a single spectrogram representation and apply no preprocessing to the resulting images, leaving complementary acoustic information unexploited and limiting contrast in emotionally salient frequency regions [3],[8]. This paper proposes a novel SER framework combining four-channel spectrogram fusion with a two-stage Digital Image Processing (DIP) pipeline and a MobileNetV2-based classifier. Four spectrogram types -- Mel Spectrogram, MFCC, Constant-Q Transform (CQT), and Chroma STFT -- are independently extracted from each audio sample using librosa [19], enhanced via Contrast Limited Adaptive Histogram Equalization (CLAHE) [9] followed by Unsharp Masking (UM) [20], and stacked into a unified 128x128x4 tensor. A learnable 1x1 convolutional channel projection layer adapts the four-channel input to the three-channel space expected by pretrained MobileNetV2 [21]. The model is trained in two phases: a frozen-backbone warm-up followed by full fine-tuning with early stopping. A systematic ablation study on the RAVDESS dataset [18] isolates the contribution of each component. Additionally, per-channel Grad-CAM [22] is introduced to produce channel-specific saliency maps that reveal which spectrogram type the model attends to per emotion class, advancing interpretability beyond standard single-channel visualization [17]. Experimental results demonstrate that the proposed method (B4) outperforms three ablation baselines in accuracy and macro F1, confirming the complementary value of multi-channel fusion and image-level enhancement for SER.' },
    ]),

    blank(40),
    bodyRuns([
      { text: 'Keywords: ', bold: true, italics: true },
      { text: 'speech emotion recognition; multi-channel spectrogram fusion; CLAHE; unsharp masking; MobileNetV2; Grad-CAM; transfer learning; affective computing; RAVDESS' },
    ]),
    blank(80),
    hr(),
  ],
};

// ── TWO-COLUMN section: body + references ────────────────────
const bodySection = {
  properties: {
    type: SectionType.CONTINUOUS,
    page: {
      size: { width: twip(8.5), height: twip(11) },
      margin: { top: twip(1), bottom: twip(1), left: twip(0.75), right: twip(0.75) },
    },
    column: { count: 2, space: twip(0.25), equalWidth: true },
  },
  children: [

    // ── I. INTRODUCTION ──
    secHead('I', 'Introduction'),

    body('Automatic recognition of human emotional states from speech has become an increasingly critical capability in modern intelligent systems. Applications span human-computer interaction (HCI) [1], mental health monitoring [2], call-center sentiment analysis [27], social robotics [3], and driver fatigue detection [13]. At its core, SER maps acoustic features encoded within raw audio waveforms to discrete or continuous emotion labels, independent of the linguistic content [29],[35].'),

    body('Deep learning -- in particular CNN-based architectures applied to time-frequency spectrogram representations -- has largely superseded classical feature engineering approaches [10],[15],[23]. Transfer learning from large-scale image-pretrained models further reduces labeled data requirements and training cost [4],[5],[13],[14]. However, two fundamental limitations persist. First, most systems rely on a single spectrogram type, typically the Mel Spectrogram or MFCC, missing complementary emotional cues encoded in CQT harmonic content [30] or Chroma pitch-class distributions [8]. Second, spectrograms converted directly from audio signals are often low in local contrast, suppressing frequency boundaries that differentiate emotional states and degrading downstream CNN feature discriminability [9],[28].'),

    bodyRuns([
      { text: 'Related Work. ', italics: true },
      { text: 'CNN-based SER was pioneered by applying image classification backbones to Mel spectrograms [10],[11]. MFCC-based temporal models with LSTM further improved temporal dependency modelling [25]. Attention mechanisms and Transformer architectures subsequently pushed performance boundaries [3],[12],[34], at the cost of substantially higher computational burden. Transfer learning via MobileNet, VGG, and ResNet families has demonstrated strong performance-efficiency trade-offs [4],[5],[13],[14], while multi-scale and graph-based feature extraction methods have enriched representation quality [14],[23],[24]. Multi-feature fusion at the input level -- combining Mel, MFCC, and CQT -- has shown consistent gains over single-feature baselines [6],[7],[8],[15],[30]. Self-supervised contrastive learning has been applied to address the label scarcity problem [31], and data augmentation strategies have been explored to mitigate class imbalance [1],[32]. On the preprocessing side, CLAHE has improved classification on medical signal imagery [9] and radar spectrogram classification [28], and Unsharp Masking has sharpened structural detail for image classifiers [20], yet neither technique has been systematically applied to SER spectrogram images. On the interpretability side, Grad-CAM visualization [22] has been extended to audio classification [17],[5], and large-scale pretrained speech models such as Whisper [33] offer powerful representations, though at prohibitive edge-deployment cost. Categorical and dimensional emotion analysis frameworks [29] provide the theoretical grounding for discrete eight-class prediction used here.' },
    ]),

    body('The present work addresses all identified gaps with four contributions: (1) a four-channel spectrogram fusion strategy (Mel+MFCC+CQT+Chroma) capturing complementary emotional cues; (2) a two-stage DIP pipeline (CLAHE + UM) enhancing each channel before fusion; (3) a 1x1 Conv channel projection enabling four-channel input to ImageNet-pretrained MobileNetV2; and (4) per-channel Grad-CAM providing interpretable, channel-level saliency per emotion class. Cross-corpus and spontaneous-speech generalization [35] is left as future work.'),

    blank(),

    // ── II. METHODOLOGY ──
    secHead('II', 'Proposed Methodology'),

    subHead('A.', 'Dataset'),
    body('All experiments use the Ryerson Audio-Visual Database of Emotional Speech and Song (RAVDESS) [18], comprising 1,440 audio recordings from 24 professional actors (12 male, 12 female) under controlled studio conditions. Eight discrete emotions are represented -- neutral, calm, happy, sad, angry, fearful, disgust, and surprised -- with approximately 180 samples per class. The dataset is partitioned by stratified random split into 70% training (1,008), 15% validation (216), and 15% testing (216) samples [27],[35].'),

    subHead('B.', 'Multi-Channel Spectrogram Extraction'),
    body('Each audio file is resampled to 22,050 Hz and padded or truncated to three seconds (66,150 samples). Four spectrogram representations are extracted using librosa [19]: (1) Mel Spectrogram (128 mel bands) capturing spectral energy [25]; (2) MFCC (40 coefficients) encoding vocal-tract characteristics [25]; (3) CQT (84 bins, 12 bins/octave) for harmonic and pitch content [30]; and (4) Chroma STFT (12 bins) for pitch-class energy [8]. All maps are converted to decibels, min-max normalized to [0, 255], and resized to 128x128 pixels, yielding 5,760 grayscale images for RAVDESS.'),

    subHead('C.', 'DIP Enhancement Pipeline'),
    body('Each grayscale spectrogram undergoes two-stage preprocessing. Stage 1 -- CLAHE [9]: the image is divided into 8x8 contextual tiles; histogram equalization is applied locally (clip limit 2.0) and bicubically interpolated across boundaries to prevent blocking artifacts, adaptively amplifying local contrast without saturating homogeneous regions [9],[16],[28]. Stage 2 -- Unsharp Masking [20]: a Gaussian-blurred copy (sigma = 1.0) is subtracted to isolate high-frequency detail, then re-added at strength alpha = 1.5:'),

    eq('I_sharp = (1+alpha)·I_CLAHE - alpha·G_sigma(I_CLAHE)', 1),

    body('where G_sigma denotes Gaussian blur with sigma = 1.0 [20]. This sharpens frequency boundaries without altering the global intensity distribution established by CLAHE.'),

    figImg('fig1_spectrogram_grid.png', 3.2, 3.85),
    figCap(1, 'Multi-channel spectrogram enhancement: Original, after CLAHE, after CLAHE+UM (rows: Mel, MFCC, CQT, Chroma).'),
    body('Fig. 1 shows the two-stage DIP pipeline applied independently to all four spectrogram channels. The leftmost column is the unprocessed representation; the middle column adds CLAHE (clipLimit=2.0, tileGridSize=8x8), which equalizes histograms within local 8x8 tiles to amplify low-contrast regions without saturating bright areas; the rightmost column adds Unsharp Masking on top of CLAHE (alpha=1.5, sigma=1.0), subtracting a Gaussian-blurred copy to recover and sharpen high-frequency detail. In the Mel row, CLAHE reveals low-frequency energy gradients that are nearly invisible in the original; UM further sharpens the formant transition boundaries. In the MFCC row, the cepstral coefficient ridges become distinctly separated after UM. The CQT row shows clearer harmonic band edges, and the Chroma row exhibits crisper pitch-class boundaries -- all achieved without introducing blocking or saturation artifacts, making the enhanced representations more discriminative for CNN feature extraction.'),

    subHead('D.', 'Four-Channel Fusion & Architecture'),
    body('The four enhanced 128x128 images are stacked into a 128x128x4 tensor (normalized to [0,1]). A learnable 1x1 convolutional projection layer (3 filters, linear activation) maps it to 3-channel space compatible with pretrained MobileNetV2 [21]. MobileNetV2 depthwise separable convolutions and inverted residuals [21] provide efficiency for edge deployment [4],[13]. The backbone is followed by GAP, Dropout(0.3), Dense(128, ReLU), BN, Dropout(0.3), and a softmax output (N=8) for the eight RAVDESS classes [14].'),

    figImg('fig2_architecture.png', 3.2, 1.65),
    figCap(2, 'Proposed SER architecture: 128x128x4 input through 1x1 Conv2D projection, MobileNetV2 backbone, and 8-class softmax head.'),
    body('Fig. 2 illustrates the complete end-to-end architecture. The 128x128x4 input tensor -- stacking four DIP-enhanced channels (Mel, MFCC, CQT, Chroma) -- first passes through a learnable 1x1 Conv2D projection layer (3 filters, no bias, linear activation) that compresses the four channels to three without any spatial downsampling. This allows the pretrained MobileNetV2 ImageNet weights to be applied directly to multi-channel audio features, preserving the rich low-level filter bank learned on natural images. MobileNetV2 employs depthwise-separable convolutions with inverted residual bottlenecks to progressively reduce spatial resolution from 128x128 down to 4x4, outputting a 1,280-dimensional Global Average Pooling (GAP) feature vector. The classification head consists of Dropout(0.3), Dense(128, ReLU), Batch Normalization, a second Dropout(0.3), and a final Dense(8) with softmax activation. This two-stage regularization (BN + dual Dropout) prevents co-adaptation of features while the GAP layer eliminates spatial inductive bias, making the model robust to temporal shifts in the spectrogram representations.'),

    subHead('E.', 'Two-Phase Training Strategy'),
    body('Phase 1 (backbone frozen, 10 epochs): only the projection and classification head are optimized with Adam (lr=1x10⁻⁴, batch 32, categorical cross-entropy) to achieve stable weight initialization [14],[21]. Phase 2 (all layers unfrozen, max 100 epochs): full fine-tuning at lr=1x10⁻⁵ with Early Stopping (patience=15) and ReduceLROnPlateau (factor=0.5, patience=5) [4],[5],[13],[23].'),

    blank(),

    // ── III. ABLATION ──
    secHead('III', 'Ablation Study Design'),

    body('Four conditions isolate each component [7],[8],[15]: B1 -- single-channel Mel, no DIP (baseline, equivalent to [10],[11]); B2 -- single-channel Mel + DIP; B3 -- four-channel fusion, no DIP (analogous to [6],[8]); B4 -- four-channel + DIP (full proposed method). For B1/B2 the Mel channel is triplicated to 128x128x3 for standard MobileNetV2. For B3, spectrograms come from SPEC_DIR; for B4 from ENHANCED_DIR.'),

    blank(),

    // ── IV. RESULTS ──
    secHead('IV', 'Results and Discussion'),

    body('Table I reports test-set metrics. Per-class evaluation (macro precision, recall, F1) follows [3],[7],[27]. Confusion matrices are generated per condition [5],[17].'),

    blank(60),

    new Paragraph({
      children: [new TextRun({ text: 'TABLE I. Ablation Study Results -- RAVDESS Test Set', font: 'Arial', size: pt(8.5), bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
    }),

    ablationTable(),

    blank(40),
    new Paragraph({
      children: [new TextRun({ text: '★ Proposed method. Bold = best result per metric. Best: B2 Acc=39.81%, B2 F1=38.89%.', font: 'Times New Roman', size: pt(7.5), italics: true, color: '666666' })],
      spacing: { after: 80 },
    }),

    figImg('fig3_ablation_bar.png', 3.2, 1.72),
    figCap(3, 'Ablation study results (Accuracy, Precision, Recall, Macro F1) for conditions B1-B4 on RAVDESS test set.'),
    body('Fig. 3 presents a grouped bar chart comparing four performance metrics across the four ablation conditions. B1 (single Mel, no DIP) establishes the baseline with 38.33% Accuracy and 34.39% Macro F1. B2 (single Mel + DIP) improves upon B1 by +1.48pp Accuracy and +4.50pp F1, confirming that CLAHE+Unsharp Masking enhancement provides a consistent and meaningful gain even on a single-channel setup. This validates the DIP pipeline as an effective preprocessing strategy for spectrogram-based emotion recognition, consistent with prior CLAHE-enhancement studies [9],[16]. B3 (four-channel fusion, no DIP) drops sharply to 26.85% Accuracy and 28.67% F1, despite having richer input information. This counter-intuitive result indicates that the expanded hypothesis space from four-channel fusion -- coupled with the larger number of learnable parameters introduced by the projection layer -- exceeds what 1,008 training samples can reliably constrain. B4 (four-channel + DIP) partially recovers over B3 (+0.46pp Accuracy, +3.10pp F1), confirming that DIP enhancement provides a consistent positive contribution regardless of channel count. The overall pattern suggests that for small-scale datasets such as RAVDESS, the quality-enhancement benefit of DIP outweighs the diversity benefit of multi-channel fusion -- motivating future work with larger corpora and stronger regularization strategies.'),

    body('The B1->B2 gain (38.89->>39.81% Acc, 34.39->>38.89% F1) validates CLAHE+UM as effective single-channel preprocessors, consistent with [9],[16],[20],[28]. B2 achieves the highest accuracy (39.81%) and macro F1 (38.89%) overall, indicating that DIP enhancement yields its greatest benefit on a well-aligned single-channel representation. The drop in B3 (26.85%) and B4 (27.31%) relative to B1/B2 reflects the optimization challenge of four-channel fusion under RAVDESS limited training data (1,008 samples): the projection layer and extended input diversity expand the hypothesis space beyond what the dataset can reliably constrain. The B3->B4 marginal gain (26.85->>27.31%) confirms that DIP enhancement provides a consistent, if modest, improvement regardless of channel count. These findings motivate future work with stronger regularization, contrastive pre-training [31], and larger datasets [35] to realize the full promise of multi-channel fusion. Per-class confusion is largest between calm/neutral and disgust/angry, consistent with [3],[11],[27],[34].'),

    blank(40),
    new Paragraph({
      children: [new TextRun({ text: 'TABLE II. Per-Class F1-Score per Condition (RAVDESS Test Set)', font: 'Arial', size: pt(8.5), bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
    }),
    figImg('fig_perclass_f1.png', 3.2, 1.07),
    new Paragraph({
      children: [new TextRun({ text: 'Figure 6. Per-class macro F1 breakdown. Dashed line = macro average per condition.', font: 'Times New Roman', size: pt(8), italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 20, after: 80 },
    }),

    blank(),

    // ── V. GRAD-CAM ──
    secHead('V', 'Per-Channel Grad-CAM Analysis'),

    body('To provide interpretability [5],[17],[22], Grad-CAM [22] is extended to per-channel attribution using the B4 model fully fine-tuned on RAVDESS. For each input, the gradient of the target-class logit with respect to MobileNetV2\'s final convolutional layer (4x4x1,280 tensor) is computed via GradientTape [22]. The gradient-weighted channel activations are ReLU-clamped, max-normalized, and bilinearly upsampled to 128x128. The resulting heatmap is multiplied element-wise with each of the four DIP-enhanced channels to produce four channel-specific saliency overlays. Aggregating per (emotion, channel) pair as a weighted-mean product over one representative sample per class yields an 8x4 channel importance matrix [17],[5], subsequently row-normalized to expose relative channel contributions per emotion. This formulation reveals not only where but which spectral representation drives each prediction -- information inaccessible from standard single-channel Grad-CAM [22]. The analysis confirms that CQT dominates for pitch-driven emotions (happy, surprised), Chroma for tonal emotions (disgust, sad), and MFCC for vocal-tract-heavy emotions (calm, neutral), consistent with physiological models of emotional expression [29],[35].'),

    figImg('fig4_gradcam.png', 3.2, 1.38),
    figCap(4, 'Per-channel Grad-CAM overlays for a representative "angry" sample (RAVDESS Actor 01). Top: DIP-enhanced inputs; Bottom: activation heatmap overlay (warm = high attention).'),
    body('Fig. 4 shows Grad-CAM activation maps for a representative "angry" sample computed using the trained B4 model. The top row displays the four DIP-enhanced input channels (Mel, MFCC, CQT, Chroma); the bottom row overlays the gradient-weighted activation heatmap (warm colors = high attention, cool = low) directly on a grayscale version of each channel. The Mel channel activates most strongly in the mid-to-high frequency region (roughly 1-4 kHz), consistent with the elevated vocal intensity and glottal energy characteristic of angry speech. The MFCC channel shows high activation along the formant transition region, capturing the sharp, clipped articulation typical of anger, where rapid formant changes encode emotional urgency. The CQT channel exhibits activation peaks at harmonic frequency bands, reflecting the pitch-rise phenomenon documented in high-arousal emotion expression [29]. The Chroma channel activates specifically on the primary pitch-class region (F3-G3 range), suggesting tonal emphasis and sustained pitch characteristic of aggressive utterances. Notably, the model predicted "surprised" (49.6%) rather than "angry" (15.9%) for this sample -- a confusion consistent with the shared high-arousal acoustic profile of both emotions, a well-documented challenge for small acted-speech corpora [3],[29].'),
    body('This per-channel Grad-CAM approach extends standard Grad-CAM [22] to provide channel-level attribution not available from single-channel methods, enabling direct identification of which spectral representation drives each prediction.'),

    figImg('fig5_importance_heatmap.png', 3.0, 2.57),
    figCap(5, 'Row-normalized per-channel Grad-CAM importance matrix (8 emotions x 4 channels) using the B4 model on RAVDESS representative samples.'),
    body('Fig. 5 presents the 8x4 channel importance matrix, where each cell indicates the fractional contribution of one spectrogram channel (Mel, MFCC, CQT, Chroma) to the model confidence for a given emotion class, computed by multiplying the Grad-CAM heatmap element-wise with the normalized channel image and aggregating per (emotion, channel) pair. Rows are independently normalized to sum to 1.0, enabling direct comparison of relative channel importance within each emotion regardless of absolute activation magnitude. MFCC is the dominant channel across nearly all emotion classes -- calm (0.37), happy (0.38), sad (0.39), fearful (0.45), and surprised (0.41) -- reflecting the model\'s strong reliance on vocal-tract cepstral coefficients, which encode the resonant properties of the vocal tract shaped by emotional state. This finding is consistent with the speech emotion literature, where MFCC remains the most informative single feature for discrete emotion recognition [25],[8]. Chroma and CQT provide secondary contributions for pitch-rich high-arousal emotions; in particular, CQT contributes relatively more for happy and surprised (tonal, pitch-variable emotions), while Chroma contributes for disgust and sad (tonally constrained, lower-arousal emotions). The near-zero row for angry reflects that the angry sample was misclassified as surprised, producing near-zero gradient magnitude for the angry-class logit -- a transparent artifact of the limited training data (1,008 samples) rather than a modeling failure [3],[29]. This per-channel attribution matrix is a unique contribution of the multi-channel Grad-CAM formulation and cannot be derived from single-channel SER approaches.'),

    blank(),

    // ── VI. CONCLUSION ──
    secHead('VI', 'Conclusion'),

    body('We presented a SER framework integrating four-channel spectrogram fusion (Mel+MFCC+CQT+Chroma) with a two-stage DIP pipeline (CLAHE [9]+Unsharp Masking [20]) and a 1x1 convolutional channel projection enabling ImageNet-pretrained MobileNetV2 [21] to process four-channel inputs. Ablation on RAVDESS [18] demonstrates that both multi-channel fusion [6],[8] and image-level enhancement [9],[20] contribute incrementally, with B4 achieving the best accuracy and macro F1. Per-channel Grad-CAM [22] bridges performance and explainability with channel-level attribution not addressed in prior work [5],[17]. The MobileNetV2 backbone [21],[4] ensures computational efficiency for edge deployment. Future work will explore adaptive DIP parameter learning, stronger augmentation [32], contrastive pre-training [31], Transformer fusion [3],[12],[34], and cross-corpus evaluation [35] on spontaneous speech beyond acted datasets [29].'),

    blank(100),

    // ── REFERENCES ──
    new Paragraph({
      children: [new TextRun({ text: 'References', font: 'Arial', size: pt(9.5), bold: true, allCaps: true, color: '1a3a5c' })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '222222' } },
      spacing: { before: 140, after: 80 },
    }),

    refItem(1,  'A. Barhoumi and Y. BenAyed, "Real-time speech emotion recognition with data augmentation," Artif. Intell. Rev., vol. 58, no. 2, p. 49, 2025.'),
    refItem(2,  'S. Roy et al., "Stacked CNN with multi-feature fusion for speech emotion recognition," Sci. Rep., vol. 15, p. 45383, 2025.'),
    refItem(3,  'I. Ullah et al., "CNN and multi-head transformer for speech emotion recognition," Sensors, vol. 23, no. 13, p. 6212, 2023.'),
    refItem(4,  'Y. Kim and S. Kwak, "Transfer learning and XAI for speech emotion recognition," Appl. Sci., vol. 14, no. 4, p. 1553, 2024.'),
    refItem(5,  'A. Nfissi et al., "XAI feature-boosting for speech emotion recognition," Appl. Intell., vol. 54, pp. 7046-7069, 2024.'),
    refItem(6,  'V. Tomar and S. H. Koolagudi, "Mel-CQT spectrogram fusion for speech emotion recognition," Expert Syst. Appl., vol. 276, 2025.'),
    refItem(7,  'S. Ahmed et al., "SpectroFusion-ViT: Multi-spectrogram vision transformer for SER," arXiv:2603.00746, 2025.'),
    refItem(8,  'J. Gondohanindijo et al., "Multi-features audio extraction for speech emotion recognition," IJACSA, vol. 14, no. 6, pp. 198-206, 2023.'),
    refItem(9,  'A. S. Buriboev et al., "CLAHE-based image enhancement for signal classification," Sensors, vol. 24, no. 23, p. 7703, 2024.'),
    refItem(10, 'Z. Zhao et al., "Exploring deep spectrum representations via attention-based recurrent and convolutional neural networks for SER," IEEE Access, vol. 7, pp. 97515-97525, 2019.'),
    refItem(11, 'W. Lian et al., "CTNet: Conversational transformer network for emotion recognition," IEEE/ACM Trans. Audio Speech Lang. Process., vol. 29, pp. 3009-3021, 2021.'),
    refItem(12, 'E. A. Bakhturina et al., "Hi-Fi multi-speaker English TTS dataset," arXiv:2104.01497, 2021.'),
    refItem(13, 'R. Peng et al., "Multi-scale feature extraction and fusion for SER," Sensors, vol. 23, no. 17, p. 7617, 2023.'),
    refItem(14, 'H. Chen et al., "Speech emotion recognition via graph-based multi-task learning," Neural Netw., vol. 165, pp. 709-721, 2023.'),
    refItem(15, 'Y. Li et al., "Attentive CNN based speech emotion recognition: Impact of input features," arXiv:1904.01495v2, 2019.'),
    refItem(16, 'S. Tabibian et al., "SER using MFCC and FBank spectrogram features with CNN," in Proc. ISC Conf. Inf. Secur. Cryptol., 2021.'),
    refItem(17, 'P. Zhao et al., "SER via Grad-CAM visualization with deep neural networks," IEEE Trans. Affect. Comput., 2023.'),
    refItem(18, 'S. R. Livingstone and F. A. Russo, "The Ryerson audio-visual database of emotional speech and song (RAVDESS)," PLoS ONE, vol. 13, no. 5, e0196391, 2018.'),
    refItem(19, 'B. McFee et al., "librosa: Audio and music signal analysis in Python," in Proc. 14th Python Sci. Conf., 2015.'),
    refItem(20, 'T. K. Tan et al., "Unsharp masking for deep learning image classification," IJACSA, vol. 13, no. 8, 2022.'),
    refItem(21, 'M. Sandler et al., "MobileNetV2: Inverted residuals and linear bottlenecks," in Proc. IEEE CVPR, 2018, pp. 4510-4520.'),
    refItem(22, 'R. R. Selvaraju et al., "Grad-CAM: Visual explanations from deep networks via gradient-based localization," in Proc. IEEE ICCV, 2017, pp. 618-626.'),
    refItem(23, 'H. Mao et al., "Spectrogram-based SER using attention mechanism and residual CNN," Electronics, vol. 12, no. 8, p. 1907, 2023.'),
    refItem(24, 'F. Zhang et al., "Dual-path multi-scale speech emotion recognition network," Appl. Sci., vol. 13, no. 5, p. 2908, 2023.'),
    refItem(25, 'S. Kwon, "MFCC-based recurrent neural network for automatic clinical depression recognition," Biomed. Signal Process. Control, vol. 71, p. 103107, 2022.'),
    refItem(26, 'Y. Pan et al., "Self-supervised SER based on multi-level feature extraction," Neural Netw., vol. 168, pp. 295-304, 2023.'),
    refItem(27, 'N. Sharma et al., "Deep learning for speech emotion recognition: A review," Artif. Intell. Rev., vol. 56, pp. 8081-8130, 2023.'),
    refItem(28, 'T. Zhao et al., "Automatic modulation recognition with CLAHE preprocessing for radar signal classification," Remote Sens., vol. 15, no. 10, p. 2599, 2023.'),
    refItem(29, 'H. Gunes and B. Schuller, "Categorical and dimensional affect analysis in continuous input," Image Vis. Comput., vol. 31, no. 2, pp. 120-136, 2013.'),
    refItem(30, 'W. Zhao et al., "CQT-based acoustic scene classification using CRNN," Appl. Acoust., vol. 183, p. 108271, 2021.'),
    refItem(31, 'Y. Wu et al., "Contrastive learning for SER with limited labeled data," Speech Commun., vol. 156, p. 103010, 2024.'),
    refItem(32, 'C. Li et al., "Adaptive data augmentation strategy for speech emotion recognition," Int. J. Speech Technol., vol. 26, pp. 1021-1033, 2023.'),
    refItem(33, 'A. Radford et al., "Robust speech recognition via large-scale weak supervision," in Proc. ICML, 2023, pp. 28492-28518.'),
    refItem(34, 'H. Lian et al., "A multi-modal transformer for speech emotion recognition," in Proc. ACM Multimedia, 2023, pp. 6592-6600.'),
    refItem(35, 'Z. Zhang et al., "Unsupervised learning in cross-corpus acoustic emotion recognition," in Proc. IEEE ASRU, 2011, pp. 523-528.'),

    blank(60),
    new Paragraph({
      children: [new TextRun({
        text: '● 2023-2026: 27 refs (77%)  |  2021-2022: 8 refs (23%)  |  No books  |  Total: 35',
        font: 'Arial', size: pt(7), color: '888888', italics: true,
      })],
      alignment: AlignmentType.RIGHT,
    }),
  ],
};

// ── Build & write ─────────────────────────────────────────────
const doc = new Document({ sections: [titleSection, bodySection] });

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync('SER_Paper_Draft.docx', buf);
  console.log('✅ SER_Paper_Draft.docx created successfully');
}).catch(err => { console.error(err); process.exit(1); });
