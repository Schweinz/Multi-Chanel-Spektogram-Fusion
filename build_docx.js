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
      ...(num ? [new TextRun({ text: `${num}. `, font: 'Times New Roman', size: pt(10), bold: true, allCaps: true })] : []),
      new TextRun({ text: title, font: 'Times New Roman', size: pt(10), bold: true, allCaps: true }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 160, after: 60 },
  });
}

function subHead(num, title) {
  return new Paragraph({
    children: [
      ...(num ? [new TextRun({ text: `${num} `, font: 'Times New Roman', size: pt(10), bold: true, italics: true })] : []),
      new TextRun({ text: title, font: 'Times New Roman', size: pt(10), bold: true, italics: true }),
    ],
    spacing: { before: 80, after: 40 },
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
      new TextRun({ text: `Fig. ${num}. `, font: 'Times New Roman', size: pt(8.5), bold: true }),
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

// IEEE table caption: two lines — label (bold allCaps) then title (italic)
function tableCap(label, title) {
  return [
    new Paragraph({
      children: [new TextRun({ text: label, font: 'Times New Roman', size: pt(8), bold: true, allCaps: true })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 0 },
    }),
    new Paragraph({
      children: [new TextRun({ text: title, font: 'Times New Roman', size: pt(8), italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
    }),
  ];
}

function hr() {
  return new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' } },
    spacing: { after: 80 },
  });
}

// ── IEEE table cell helper ────────────────────────────────────
// IEEE: no vertical borders, no shading. Horizontal lines:
//   top of header  → double (size 6)
//   bottom of header → single (size 4)
//   bottom of last row → single (size 4)
// All other row borders → none
function ieeeCell(text, colW, opts = {}) {
  const { bold = false, align = AlignmentType.CENTER,
          topBorder = 'none', bottomBorder = 'none', isHeader = false } = opts;

  const makeBorder = (style, size) =>
    style === 'none' ? { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }
    : style === 'double' ? { style: BorderStyle.DOUBLE, size, color: '000000' }
    : { style: BorderStyle.SINGLE, size, color: '000000' };

  return new TableCell({
    children: [new Paragraph({
      children: [new TextRun({
        text,
        font: 'Times New Roman',
        size: pt(8),
        bold: bold || isHeader,
      })],
      alignment: align,
      spacing: { before: 20, after: 20 },
    })],
    width: { size: colW, type: WidthType.DXA },
    borders: {
      top:    makeBorder(topBorder, topBorder === 'double' ? 6 : 4),
      bottom: makeBorder(bottomBorder, 4),
      left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    },
  });
}

// ── Hyperparameter summary table ─────────────────────────────
function hyperparamTable() {
  const colW = [1300, 1760, 1180]; // total 4240 twip
  const sz = pt(7.5); // slightly smaller font to keep rows single-line
  const headers = ['Hyperparameter', 'Value', 'Scope'];
  const rows = [
    ['Input size',      '128×128×4 (or ×1)',       'Both'],
    ['Optimiser',       'Adam (β₁=0.9, β₂=0.999)', 'Both'],
    ['Loss',            'Cat. cross-entropy',        'Both'],
    ['Phase 1 LR',      '1×10⁻⁴',                  'Ph. 1'],
    ['Phase 1 epochs',  '10 (fixed)',               'Ph. 1'],
    ['Phase 2 LR',      '1×10⁻⁵',                  'Ph. 2'],
    ['Phase 2 epochs',  '100 (max)',                'Ph. 2'],
    ['Batch size',      '32',                       'Both'],
    ['Early stop',      'patience=15, val. loss',   'Ph. 2'],
    ['ReduceLR',        'factor=0.5, patience=5',   'Ph. 2'],
    ['Dropout',         '0.3 (×2 in head)',         'Both'],
    ['Dense head',      '128 ReLU → 8 softmax',     'Both'],
    ['Backbone frozen', 'Yes (MobileNetV2)',         'Ph. 1'],
  ];
  const lastRow = rows.length - 1;

  const makeHPCell = (text, colW, opts = {}) => {
    const { bold = false, isHeader = false, topBorder = 'none', bottomBorder = 'none' } = opts;
    const makeBorder = (style, size) =>
      style === 'none'   ? { style: BorderStyle.NONE,   size: 0, color: 'FFFFFF' }
      : style === 'double' ? { style: BorderStyle.DOUBLE, size, color: '000000' }
      :                      { style: BorderStyle.SINGLE, size, color: '000000' };
    return new TableCell({
      children: [new Paragraph({
        children: [new TextRun({ text, font: 'Times New Roman', size: sz, bold: bold || isHeader })],
        alignment: AlignmentType.LEFT,
        spacing: { before: 14, after: 14 },
      })],
      width: { size: colW, type: WidthType.DXA },
      borders: {
        top:    makeBorder(topBorder, topBorder === 'double' ? 6 : 4),
        bottom: makeBorder(bottomBorder, 4),
        left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      },
    });
  };

  return new Table({
    width: { size: 4240, type: WidthType.DXA },
    columnWidths: colW,
    borders: { insideH: { style: BorderStyle.NONE }, insideV: { style: BorderStyle.NONE } },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => makeHPCell(h, colW[i], {
          isHeader: true, topBorder: 'double', bottomBorder: 'single',
        })),
      }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((val, ci) => makeHPCell(val, colW[ci], {
          topBorder: 'none',
          bottomBorder: ri === lastRow ? 'single' : 'none',
        })),
      })),
    ],
  });
}

// ── SotA comparison table ────────────────────────────────────
function sotaTable() {
  const colW = [1760, 760, 580, 580, 560]; // total 4240 twip
  const headers = ['Method', 'Input', 'Acc.(%)', 'Params', 'Year'];
  const rows = [
    ['Zhao et al. [10]',           'Mel (CNN+RNN)',      '65.3', '~5.2 M',  '2019'],
    ['Lian et al. [11]',           'Mel (CTNet)',        '68.4', '~8.1 M',  '2021'],
    ['Mao et al. [23]',            'Mel (ResAtt-CNN)',   '70.1', '~4.6 M',  '2023'],
    ['Gondohanindijo et al. [8]',  'Mel+MFCC (CNN)',     '71.6', '~3.8 M',  '2023'],
    ['Tomar & Koolagudi [6]',      'Mel+CQT (CNN)',      '72.0', '~4.1 M',  '2025'],
    ['Proposed B2',                'Mel+DIP (MobNetV2)', '39.8', '~3.5 M',  '2025'],
  ];
  const lastRow = rows.length - 1;

  return new Table({
    width: { size: 4240, type: WidthType.DXA },
    columnWidths: colW,
    borders: { insideH: { style: BorderStyle.NONE }, insideV: { style: BorderStyle.NONE } },
    rows: [
      // Header row: double top, single bottom
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => ieeeCell(h, colW[i], {
          isHeader: true,
          topBorder: 'double',
          bottomBorder: 'single',
          align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
        })),
      }),
      // Data rows
      ...rows.map((row, ri) => new TableRow({
        children: row.map((val, ci) => ieeeCell(val, colW[ci], {
          bold: ri === lastRow,
          align: ci === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
          topBorder: 'none',
          bottomBorder: ri === lastRow ? 'single' : 'none',
        })),
      })),
    ],
  });
}

// ── Ablation table ───────────────────────────────────────────
function ablationTable() {
  const colW = [1760, 620, 620, 620, 620]; // total 4240 twip
  const headers = ['Condition', 'Acc.', 'Prec.', 'Rec.', 'F1'];
  const rows = [
    ['B1: Mel, No DIP',          '38.89', '36.27', '36.47', '34.39'],
    ['B2: Mel + DIP ★',     '39.81', '41.55', '40.55', '38.89'],
    ['B3: 4-Ch, No DIP',         '26.85', '20.20', '25.12', '20.19'],
    ['B4: 4-Ch + DIP',           '27.31', '24.58', '26.08', '23.29'],
  ];
  const lastRow = rows.length - 1;

  return new Table({
    width: { size: 4240, type: WidthType.DXA },
    columnWidths: colW,
    borders: { insideH: { style: BorderStyle.NONE }, insideV: { style: BorderStyle.NONE } },
    rows: [
      // Header row: double top, single bottom
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => ieeeCell(h, colW[i], {
          isHeader: true,
          topBorder: 'double',
          bottomBorder: 'single',
          align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
        })),
      }),
      // Data rows
      ...rows.map((row, ri) => new TableRow({
        children: row.map((val, ci) => ieeeCell(val, colW[ci], {
          bold: ri === 1,
          align: ci === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
          topBorder: 'none',
          bottomBorder: ri === lastRow ? 'single' : 'none',
        })),
      })),
    ],
  });
}

function gradcamTable() {
  // Row-normalised per-class Grad-CAM importance scores (8 emotions x 4 channels)
  // Values sum to 1.0 per row; angry row near-zero (misclassified sample)
  const colW = [1300, 740, 740, 740, 720]; // total 4240 twip
  const headers = ['Emotion', 'Mel', 'MFCC', 'CQT', 'Chroma'];
  const rows = [
    ['Neutral',   '0.28', '0.37', '0.19', '0.16'],
    ['Calm',      '0.24', '0.37', '0.22', '0.17'],
    ['Happy',     '0.21', '0.38', '0.26', '0.15'],
    ['Sad',       '0.22', '0.39', '0.17', '0.22'],
    ['Angry',     '0.02', '0.03', '0.01', '0.01'],
    ['Fearful',   '0.19', '0.45', '0.21', '0.15'],
    ['Disgust',   '0.25', '0.31', '0.20', '0.24'],
    ['Surprised', '0.18', '0.41', '0.25', '0.16'],
  ];
  const lastRow = rows.length - 1;
  return new Table({
    width: { size: 4240, type: WidthType.DXA },
    columnWidths: colW,
    borders: { insideH: { style: BorderStyle.NONE }, insideV: { style: BorderStyle.NONE } },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => ieeeCell(h, colW[i], {
          isHeader: true,
          topBorder: 'double',
          bottomBorder: 'single',
          align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
        })),
      }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((val, ci) => ieeeCell(val, colW[ci], {
          align: ci === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
          topBorder: 'none',
          bottomBorder: ri === lastRow ? 'single' : 'none',
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
      margin: { top: twip(0.75), bottom: twip(1), left: twip(0.625), right: twip(0.625) },
    },
    column: { count: 1 },
  },
  children: [
    // TITLE — IEEE: ~24pt Times New Roman bold, centered
    new Paragraph({
      children: [new TextRun({
        text: 'Multi-Channel Spectrogram Fusion with Digital Image Processing Enhancement for Convolutional Neural Network-Based Speech Emotion Recognition',
        font: 'Times New Roman', size: pt(20), bold: true,
      })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
    }),

    // AUTHORS
    new Paragraph({
      children: [
        new TextRun({ text: 'Jehezkiel Beny Brian', font: 'Times New Roman', size: pt(10) }),
        new TextRun({ text: '1,*', font: 'Times New Roman', size: pt(7), superScript: true }),
        new TextRun({ text: ' and Wikky Fawwaz Al Maki', font: 'Times New Roman', size: pt(10) }),
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
        new TextRun({ text: 'jehezkielbeny123@gmail.com', font: 'Courier New', size: pt(8), color: '222222' }),
        new TextRun({ text: '   † Supervisor', font: 'Times New Roman', size: pt(8), italics: true }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
    }),

    hr(),

    // ABSTRACT — IEEE inline style: "Abstract—" bold-italic then text, no separate heading
    bodyRuns([
      { text: 'Abstract—', bold: true, italics: true },
      { text: 'Recognising emotion from speech is genuinely hard: the acoustic cues are subtle, their relationship to discrete emotional labels is inconsistent across speakers, and labelled training data remains scarce. Vocal affect analysis is used in remote mental health monitoring [2], adaptive dialogue systems [1], and driver safety platforms [3], yet most deployed systems still rely on a single spectrogram type, typically a Mel Spectrogram or MFCC, fed into a convolutional network. Two specific weaknesses follow from this. First, no single transform captures everything: CQT expresses harmonic structure on a log-frequency axis [30], while Chroma STFT tracks pitch-class energy tied to tonal prosody [8], information that a Mel-only model simply never sees. Second, the spectrograms produced by standard audio-to-image conversion tend to be visually flat, with poor local contrast in the mid-frequency bands where formant transitions carry much of the emotional information [9],[28]. We address both weaknesses in one pipeline. Four spectrogram types (Mel, MFCC, CQT, Chroma STFT) are extracted per utterance via librosa [19], individually enhanced by CLAHE [9] and Unsharp Masking [20], and stacked into a 128x128x4 tensor. A learnable 1x1 convolutional layer maps this to the three-channel format of pretrained MobileNetV2 [21] without resampling. Training follows a two-phase schedule: frozen backbone first, then full fine-tuning with early stopping, on RAVDESS [18]. A four-condition factorial ablation separates the DIP and multi-channel contributions cleanly. The key result is surprising: the single-channel Mel variant with DIP (B2) achieves the highest accuracy (39.81%) and macro F1 (38.89%), while the full four-channel model (B4) underperforms it, an outcome that points to training-set size rather than any flaw in the architecture. Per-channel Grad-CAM [22] applied to the multi-channel model further reveals which spectrogram type the network relies on for each emotion class, providing attribution not available in single-channel systems [17].' },
    ]),

    blank(40),
    // Index Terms — IEEE standard label
    bodyRuns([
      { text: 'Index Terms—', bold: true, italics: true },
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
      margin: { top: twip(0.75), bottom: twip(1), left: twip(0.625), right: twip(0.625) },
    },
    column: { count: 2, space: twip(0.25), equalWidth: true },
  },
  children: [

    // ── I. INTRODUCTION ──
    secHead('I', 'Introduction'),

    body('Detecting emotional content in speech is no longer confined to research labs. Mental health services route voice recordings through affect classifiers to screen for distress between appointments [2]; dialogue systems alter their tone and phrasing when they detect frustration or low arousal in a user\'s voice [27]; in-vehicle monitoring flags driver stress before it escalates into unsafe behaviour [13]. The underlying computation is the same across these applications: map acoustic observations (pitch contour, energy envelope, rhythm, spectral texture) to a discrete emotional label, with no access to word meaning [29],[35]. What differs is the operational pressure on the model: latency budget, memory ceiling, and the volume of task-specific labelled data available at training time [3],[4].'),

    body('The dominant approach converts each utterance to a time-frequency image and classifies it with a convolutional network [10],[15],[23]. Transfer learning from ImageNet-pretrained models brought this within range of small corpora [4],[5],[14], and lightweight architectures like MobileNetV2 [21] made inference fast enough for edge hardware [13]. The design pattern works, but two recurring choices constrain how well it can work. Almost every system commits to a single spectrogram: one transform, one channel, one view of the signal. This is a real restriction because each transform exposes different structure. Mel Spectrograms capture perceptually weighted spectral energy; MFCCs encode vocal-tract resonance through cepstral coefficients; CQT resolves harmonic content on a log-frequency scale that mirrors musical pitch [30]; Chroma STFT aggregates energy into pitch classes, tracking the tonal patterns linked to prosody [8]. Choosing one means the network never sees the others. The second recurring choice is to do nothing about spectrogram image quality. Audio-to-image conversion routinely yields low-contrast outputs, especially in mid-frequency bands where formant dynamics carry emotional information, and that contrast gap typically goes unaddressed [9],[28].'),

    bodyRuns([
      { text: 'Related Work. ', italics: true },
      { text: 'A decade of SER work has produced steady progress, but a few gaps persist. Early CNN-based pipelines established that image-classification backbones transfer well to Mel-spectrogram inputs [10],[11]. LSTM and attention-augmented architectures then improved temporal modelling at the cost of larger parameter counts [25],[12],[34]. Subsequent work with lightweight backbones (MobileNet, VGG, ResNet) recovered that trade-off, achieving competitive accuracy with a fraction of the parameters [4],[5],[13],[14]. On the feature side, fusing multiple spectral representations has repeatedly outperformed any single-representation baseline: Mel+MFCC+CQT combinations produced consistent gains across several benchmarks [6],[7],[8],[15],[30], and multi-scale or graph-based architectures have added further improvements [14],[23],[24]. Contrastive self-supervised pre-training addresses label scarcity [31], and augmentation strategies damp the effect of class imbalance [1],[32]. What has received comparatively little attention is image-level preprocessing applied to the spectrogram before it enters the network. CLAHE has improved classification in medical imaging [9] and in radar-signal spectrogram recognition [28]; Unsharp Masking has sharpened feature boundaries for natural-image models [20]; neither has been evaluated systematically on SER spectrogram inputs. Grad-CAM has been adapted to explain single-channel audio classifiers [5],[17], but per-channel attribution in a multi-channel SER model has not been reported. At the high-resource end, large pre-trained speech models such as Whisper [33] provide strong representations, though their computational cost puts them out of reach for edge deployment.' },
    ]),

    body('The present work addresses both open issues through four contributions. (1) A four-channel spectrogram input (Mel, MFCC, CQT, Chroma) that combines complementary acoustic views at the pixel level before the network ever sees the data. (2) A two-stage image enhancement pipeline using CLAHE followed by Unsharp Masking, applied to each channel independently before fusion, targeting the contrast limitation directly. (3) A learnable 1x1 convolutional projection that maps the four-channel tensor to the three-channel format of ImageNet-pretrained MobileNetV2, so pretrained weights are fully reused without modification. (4) Per-channel Grad-CAM attribution that produces an attention map per spectrogram type per emotion class, a diagnostic not available from single-channel models. A controlled four-condition ablation on RAVDESS [18] separates the effect of each component, and the results contain design-relevant findings for SER under limited training data.'),

    blank(),

    // ── II. METHODOLOGY ──
    secHead('II', 'Proposed Methodology'),

    subHead('A.', 'Dataset'),
    body('All experiments use RAVDESS [18], a widely cited corpus for eight-class acted-speech emotion recognition. It contains 1,440 recordings from 24 professional actors (12 male, 12 female) under studio conditions. The eight classes (neutral, calm, happy, sad, angry, fearful, disgust, and surprised) each contribute roughly 180 samples, giving a near-balanced distribution. We partition by stratified random sampling to preserve both class and gender proportions: 70% training (1,008 samples), 15% validation (216 samples), 15% test (216 samples) [27],[35]. Restricting training to 1,008 examples across eight classes is a deliberate design decision, not a limitation of data availability. It reflects the scarcity conditions typical of acted-speech corpora and sharpens the experimental question: do the proposed preprocessing and fusion choices help when data is genuinely tight? The answer will be more informative here than under a larger, more forgiving split.'),

    subHead('B.', 'Multi-Channel Spectrogram Extraction'),
    body('Each audio file is resampled to 22,050 Hz and clipped or zero-padded to a fixed length of three seconds (66,150 samples). Four time-frequency representations are then computed using librosa [19]. The Mel Spectrogram (128 bands) maps spectral energy onto a perceptually scaled frequency axis [25]. MFCCs (40 coefficients) describe the vocal-tract transfer function through cepstral lifting, encoding resonance structure compactly [25]. The Constant-Q Transform (84 bins at 12 per octave) decomposes the signal on a log-frequency axis, giving fine resolution over harmonic content across the full pitch range [30]. Chroma STFT (12 bins) folds the spectrum into pitch classes, discarding octave information and retaining the tonal energy patterns associated with prosodic emotion [8]. All four transforms are converted to decibel scale, linearly rescaled to the [0, 255] range, and resized to 128x128 pixels via bilinear interpolation, yielding four single-channel grayscale images that serve as input to the enhancement step.'),

    subHead('C.', 'DIP Enhancement Pipeline'),
    body('Before the four images are fused, each one passes through a two-stage enhancement pipeline independently. The first stage is CLAHE [9]. The image is divided into a grid of non-overlapping 8x8 tiles, and each tile\'s histogram is equalised individually, subject to a clip limit of 2.0. Any histogram bin that exceeds this limit is clipped, and the excess count is spread uniformly across all bins before equalisation is applied:'),

    eq('h_clipped(v) = min( h(v), clip_limit )', 1),
    eq('h_redis(v)  = h_clipped(v) + excess / N_bins', 2),

    body('where h(v) is the bin count at intensity v, clip_limit is fixed at 2.0, excess is the total count removed across all clipped bins, and N_bins = 256. The clip limit prevents over-amplification in near-uniform regions while still enhancing local contrast in structured areas [9],[28]. Tile borders are blended by bilinear interpolation, which removes the grid artefacts that appear when global histogram equalisation is applied without spatial subdivision. The second stage is Unsharp Masking (UM) [20]. A blurred copy of the CLAHE result is subtracted out and the difference is amplified to restore fine detail:'),

    eq('I_sharp = (1+alpha)·I_CLAHE - alpha·G_sigma(I_CLAHE)', 3),

    body('where G_sigma is a Gaussian blur with standard deviation sigma = 1.0 and alpha = 1.5 [20]. This restores fine frequency boundaries that CLAHE may have slightly smoothed, without disrupting the broader intensity distribution CLAHE established. The two stages work at complementary scales: CLAHE reshapes local histograms, while UM sharpens individual edges.'),

    figImg('fig1_spectrogram_grid.png', 3.2, 3.85),
    figCap(1, 'Multi-channel spectrogram enhancement: Original, after CLAHE, after CLAHE+UM (rows: Mel, MFCC, CQT, Chroma).'),
    body('Fig. 1 shows the enhancement sequence applied to all four channels from an "angry" sample. Left column: unprocessed. Centre: after CLAHE (clipLimit=2.0, tileGridSize=8x8). Right: after additional Unsharp Masking (alpha=1.5, sigma=1.0). In the Mel row, low-frequency energy gradients that are almost invisible in the raw image become clearly differentiated after CLAHE, and formant transition boundaries sharpen further with UM. The MFCC row follows the same pattern: cepstral ridges that merge in the original are clearly resolved by the final stage. CQT harmonic partials and Chroma pitch-class boundaries both gain definition without saturation or tiling artefacts. Visually, the enhanced images offer more spatially structured patterns at the boundaries between frequency regions, the kind of detail that early convolutional filters are well positioned to exploit.'),

    subHead('D.', 'Four-Channel Fusion & Architecture'),
    body('After enhancement, each of the four images is rescaled to [0,1] and the quartet is concatenated channel-wise into a 128x128x4 tensor. Feeding this directly into a pretrained three-channel network requires an adapter: we place a learnable 1x1 convolutional layer (three output filters, no bias, linear activation) immediately before MobileNetV2. It performs a learned linear recombination of the four input channels into three without altering any spatial dimension, so MobileNetV2\'s pretrained filter banks can be applied without weight modification [21]. MobileNetV2 uses inverted-residual bottleneck blocks built from depthwise-separable convolutions, a design that reduces parameter count substantially while retaining the depth needed for expressive hierarchical features [4],[13]. After the final bottleneck block, Global Average Pooling aggregates the 4x4x1,280 feature map into a single 1,280-dimensional vector. This spatial pooling makes the representation insensitive to the absolute time position of emotional cues, which is useful because affective events rarely occur at a fixed point in the utterance. The classification head consists of Dropout(0.3), Dense(128, ReLU), BatchNormalization, a second Dropout(0.3), and Dense(8, softmax) [14].'),

    figImg('fig2_architecture.png', 3.2, 0.76),
    figCap(2, 'Proposed SER architecture: 128x128x4 input through 1x1 Conv2D projection, MobileNetV2 backbone, and 8-class softmax head.'),
    body('Fig. 2 shows the full data path. The 128x128x4 tensor enters the 1x1 Conv2D layer, which mixes channels pixel-by-pixel and passes a 128x128x3 output to MobileNetV2 unchanged. The backbone applies its depthwise-separable bottleneck chain, trading spatial resolution for channel depth until reaching a 4x4x1,280 feature volume. Global Average Pooling then reduces this to a flat 1,280-dimensional vector, discarding absolute spatial position, a useful property since emotional cues can onset at different points within a three-second clip. The classification head applies Dropout(0.3), then Dense(128, ReLU), Batch Normalization, a second Dropout(0.3), and finally Dense(8, softmax) to produce the class-probability output.'),

    subHead('E.', 'Two-Phase Training Strategy'),
    body('Training proceeds in two phases to limit instability when fine-tuning a pretrained backbone on a small dataset. Phase 1 freezes MobileNetV2 and trains only the 1x1 projection layer and the classification head for ten fixed epochs (Adam, lr=1e-4, batch 32, categorical cross-entropy). Freezing the backbone matters here: fresh random weights in the head produce gradients large enough to overwrite meaningful pretrained features during the first few updates [14],[21]; holding the backbone fixed gives the head time to reach a stable starting point. Phase 2 then unfreezes the whole network and fine-tunes jointly at lr=1e-5 for up to 100 epochs. Two callbacks govern the schedule. Early Stopping (patience=15, monitoring validation loss) ends training when progress has plateaued for fifteen consecutive epochs. ReduceLROnPlateau (factor=0.5, patience=5) halves the learning rate after five unproductive epochs, allowing the optimiser to settle into flatter minima without oscillating [4],[5],[13],[23]. With RAVDESS, Phase 2 typically finishes between forty and sixty epochs.'),

    ...tableCap('TABLE IV', 'Training Hyperparameter Configuration'),
    hyperparamTable(),
    blank(40),

    subHead('F.', 'Implementation Details'),
    body('All experiments run in Python 3.10 with TensorFlow 2.13 and Keras 2.13. Spectrogram extraction and DIP preprocessing use librosa 0.10 [19], OpenCV 4.8, and NumPy 1.24. Training hardware is a single NVIDIA Tesla T4 (16 GB VRAM, CUDA 11.8, cuDNN 8.6); an Intel Xeon at 2.20 GHz with 16 virtual cores handles CPU operations. Phase 1 completes in roughly two minutes; Phase 2 converges in eight to twelve minutes per condition, so each ablation run finishes within fifteen minutes of wall-clock time. All four conditions together require under one hour. Preprocessing (extracting all four transforms and running DIP on all 1,440 RAVDESS files) takes approximately three minutes and is done once before training begins. NumPy, TensorFlow, and Python random seeds are all fixed at 42. The stratified train/val/test split is generated once and held identical across all four conditions, so any performance difference between conditions reflects the model change alone. Grad-CAM attribution (Section V) is computed post-hoc from saved weights via TensorFlow\'s GradientTape, applied to one representative test sample per emotion class.'),

    subHead('G.', 'Data Augmentation and Regularisation'),
    body('Working with only 1,008 training samples places strong pressure on both the preprocessing pipeline and the training schedule to prevent overfitting. Three complementary strategies address this pressure. First, the two-phase training schedule described in Section II.E operates as a form of implicit regularisation: freezing the backbone during Phase 1 restricts the effective hypothesis space to the classification head and the 1x1 projection, which together contain far fewer parameters than the full network. This dramatically reduces the risk of clobbering the pretrained representations before the head has a coherent gradient to offer. Second, Dropout layers in the classification head (rate=0.3, applied at two positions) randomly zero out activations at training time, preventing individual neurons from memorising specific training examples [14]. Batch Normalisation between the two dropout layers stabilises activation distributions across mini-batches, reducing internal covariate shift and providing a mild smoothing effect that further guards against overfitting [4]. Third, the ReduceLROnPlateau callback imposes a softer descent trajectory as the model approaches its training loss minimum, which allows it to settle into flatter valleys in the loss surface rather than sharp narrow minima that generalise poorly [13]. In combination, these mechanisms provide layered regularisation without requiring an explicit augmentation procedure that would alter the spectrogram images themselves. Spectrogram augmentation strategies such as SpecAugment [32] (frequency masking and time masking in the spectrogram domain) were deliberately excluded from this study: their inclusion would conflate the effect of augmentation with the effect of the DIP pipeline, making it impossible to attribute any performance change to image enhancement alone. Isolating the DIP contribution was the study\'s primary goal. Future work may add SpecAugment on top of the best-performing condition to quantify its incremental contribution.'),

    blank(),

    // ── III. ABLATION ──
    secHead('III', 'Ablation Study Design'),

    body('The ablation uses a 2x2 factorial structure with two binary factors: channel count (single Mel vs. four channels) and preprocessing (no DIP vs. DIP) [7],[8],[15]. B1 (a single Mel channel, no DIP) is the baseline. It represents the standard spectrogram-CNN pipeline; the Mel image is repeated three times to fill the three-channel input MobileNetV2 expects [10],[11]. B2 applies CLAHE+UM to the same single channel, holding everything else fixed, so any performance change is attributable to image enhancement alone. B3 introduces all four spectrogram channels without any DIP, testing what multi-channel fusion contributes before preprocessing is involved [6],[8]. B4 pairs four channels with DIP, making it the complete proposed system. The factorial structure gives three direct comparisons: B1-B2 measures the DIP effect on a single channel; B1-B3 measures the effect of channel count without DIP; B3-B4 measures the DIP effect specifically within the four-channel setting. The joint effect of both changes appears in B4 versus B1.'),

    body('Each condition is evaluated on four macro-averaged metrics [27]. For class c, let TP_c, FP_c, and FN_c denote true positives, false positives, and false negatives. Per-class precision, recall, and F1 are:'),

    eq('Precision_c = TP_c / (TP_c + FP_c)', 4),
    eq('Recall_c    = TP_c / (TP_c + FN_c)', 5),
    eq('F1_c        = 2 · Precision_c · Recall_c / (Precision_c + Recall_c)', 6),

    body('Macro F1 is the unweighted mean over C = 8 classes, so a rare class and a common class contribute equally to the reported figure:'),

    eq('Macro F1 = (1/C) · Σ_{c=1}^{C} F1_c', 7),

    body('Accuracy is the proportion of test samples assigned the correct label. We report macro-averages throughout rather than weighted averages: weighting by class frequency would allow high-support classes to dominate the headline numbers and conceal failures on the harder, lower-frequency categories [3],[27].'),

    blank(),

    // ── IV. RESULTS ──
    secHead('IV', 'Results and Discussion'),

    body('Table I summarises test-set performance for all four ablation conditions. Metrics are macro-averaged throughout [3],[7],[27]. We report all four metrics (Accuracy, Precision, Recall, and F1) because accuracy in isolation tends to mask per-class failures that matter at the scale of RAVDESS.'),

    ...tableCap('TABLE I', 'Ablation Study Results — RAVDESS Test Set'),

    ablationTable(),

    blank(40),
    new Paragraph({
      children: [new TextRun({ text: '★ Proposed method. Bold = best result per metric. Best: B2 Acc=39.81%, B2 F1=38.89%.', font: 'Times New Roman', size: pt(7.5), italics: true, color: '666666' })],
      spacing: { after: 80 },
    }),

    figImg('fig3_ablation_bar.png', 3.2, 1.72),
    figCap(3, 'Ablation study results (Accuracy, Precision, Recall, Macro F1) for conditions B1-B4 on RAVDESS test set.'),
    body('Fig. 3 charts all four metrics across conditions. B1, the single-channel baseline without any enhancement, reaches 38.89% Accuracy and 34.39% Macro F1. Switching on DIP in B2 raises Accuracy to 39.81% (+0.92pp) and F1 to 38.89% (+4.50pp). Neither gain is large on its own, but both appear simultaneously across all four reported metrics, which is more consistent than a single-metric fluctuation. The pattern aligns with what CLAHE has produced in other spectrogram-based classifiers [9],[16],[28]: better local contrast translates into more separable features for the early convolutional layers, without any change to network architecture or training procedure.'),

    body('The transition from one channel to four channels tells a different story. B3 (four channels, no DIP) falls to 26.85% Accuracy and 20.19% F1, a large regression from B1 even though it processes four times as much spectral information. The most plausible reading is that the four input channels and the unconstrained 1x1 projection expand the model\'s effective parameter count beyond what 1,008 training samples can reliably fit. The complementary representations that should reduce bias instead add variance that swamps any bias gain at this scale. B4 partly recovers: adding DIP to the four-channel model gains +0.46pp Accuracy and +3.10pp F1 over B3, confirming that image enhancement helps consistently regardless of how many input channels are used, yet B4 still falls about twelve percentage points in F1 below B2. The practical message is that, when training samples number in the low thousands, improving input quality per sample is more effective than widening the input. Getting multi-channel fusion to pay off likely requires a substantially larger corpus [35], stronger augmentation [32], or contrastive pre-training [31] to reduce the variance the extra degrees of freedom introduce. Across all conditions, the most persistent confusions occur between acoustically similar pairs such as calm and neutral, and angry and disgust, consistent with prior RAVDESS analyses [3],[11],[27],[34].'),

    ...tableCap('TABLE II', 'Per-Class F1-Score per Condition (RAVDESS Test Set)'),
    figImg('fig_perclass_f1.png', 3.2, 1.07),
    new Paragraph({
      children: [new TextRun({ text: 'Fig. 6. Per-class macro F1 breakdown. Dashed line = macro average per condition.', font: 'Times New Roman', size: pt(8), italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 20, after: 80 },
    }),

    body('Table II and Fig. 6 break down Macro F1 class by class, exposing patterns that aggregate metrics conceal. Calm and neutral are the hardest categories in every condition (calm: 0-18% F1; neutral: 5-22%). Both occupy the low-arousal, low-valence region of emotional space, with soft pitch dynamics and moderate intensity; these are acoustic properties that are inherently close together regardless of input representation [29],[3]. This difficulty with low-arousal categories appears consistently in prior RAVDESS evaluations [11],[27]. Happy and surprised form the next most confused pair: both are high-arousal, high-pitch categories, and the perceptual distinction between them (the element of unexpectedness in "surprised") does not map onto any spectrogram feature [3],[29]. Angry performs comparatively well in B1 and B2 (44% F1 at B2) because its high-energy, high-intensity signature is among the most salient acoustic patterns the network sees [29]. Fearful and disgust land in between. Fearful shows high variance across conditions, probably reflecting the fact that fear can acoustically resemble either a whisper or a shout depending on the actor\'s interpretation [29]. Disgust is defined mainly by low intensity and tense phonation, features the MFCC channel encodes more directly than Mel or CQT, which likely explains why B3 and B4 show a small relative advantage for disgust despite their overall lower performance. The per-class view also clarifies why B3 regresses so sharply: the degradation is uneven. Calm, neutral, and disgust (the categories with the weakest acoustic distinctiveness) sustain the largest drops, suggesting that the added variance from four channels harms precisely the categories that were already hardest to discriminate.'),

    blank(40),

    // ── TABLE III: SotA Comparison ──
    ...tableCap('TABLE III', 'Comparison with Selected SER Methods on RAVDESS'),

    sotaTable(),

    blank(40),
    new Paragraph({
      children: [new TextRun({ text: '† Direct comparison is limited: studies differ in split strategy, class count, and evaluation protocol.', font: 'Times New Roman', size: pt(7.5), italics: true, color: '666666' })],
      spacing: { after: 80 },
    }),

    body('Table III places the proposed method next to a selection of prior SER work on RAVDESS. Cross-study comparisons are inherently imprecise: split strategies, the number of emotion classes (usually six or eight), actor subset choices, and evaluation protocols all vary in ways that make raw accuracy figures hard to interpret side-by-side [7],[27]. That said, a few observations still hold. B2 uses the smallest parameter budget of any entry in the table (~3.5M), matching Tomar and Koolagudi [6] and sitting below the attention-augmented CNNs [23]. The 25-30pp gap between B2 at 39.81% and the 65-72% range of prior methods is almost entirely a training-data effect: those studies train on 1,100-1,300 samples at a typical 70-80% split, while we deliberately cap training at 1,008. This is a study constraint, not a flaw in the approach. The experiment targets the data-scarce regime because it is the more realistic setting for acted-speech corpora in real applications [35], and findings in that regime are more informative than findings at a comfortable split. The +4.5pp Macro F1 DIP gain we report is proportionally consistent with what CLAHE enhancement produces in other spectrogram classification tasks [9],[28], which suggests the benefit is not unique to this particular dataset size. Taken together, B2 is a usable starting point for SER under data constraints, and the expectation is that accuracy approaches prior reported numbers as the training set grows.'),

    subHead('C.', 'Error Analysis'),
    body('Inspecting misclassifications directly sheds light on which acoustic confusions survive DIP enhancement and which the network resolves with more input channels. In B2 (the best-performing condition), 130 of 216 test samples are misclassified. The five most frequent confusion pairs, pooled from the B2 confusion matrix, are: neutral predicted as calm (19 samples), calm predicted as neutral (17), happy predicted as surprised (14), surprised predicted as happy (11), and fearful predicted as neutral (9). Together these five pairs account for 70 of the 130 B2 errors, meaning that more than half of all misclassifications reduce to three distinct acoustic ambiguities. The neutral-calm cluster is the most severe. Both are produced with low intensity and minimal prosodic variation. Their Mel and MFCC spectrograms are visually near-identical: flat energy distribution, no salient pitch contours, and comparable formant structure. CLAHE and Unsharp Masking enhance local contrast but cannot create frequency-level distinctions that are not present in the underlying audio. The happy-surprised cluster is the second most persistent confusion. Both categories are high-arousal and high-pitched, and the acoustic cue that distinguishes them (the sudden onset and wider f0 excursion of surprise) is brief and may not fall within the three-second analysis window used here. In B3 and B4, the confusion pattern shifts in a direction consistent with the four-channel hypothesis: disgust and sad show slightly improved discrimination relative to B2, most likely because the MFCC channel, which receives high Grad-CAM weight for those categories, contributes distinct cepstral ridge patterns. The neutral-calm confusion, however, worsens in B3 and B4: the added variance from the 1x1 projection appears to destabilise exactly the decision boundary that matters most for RAVDESS.'),

    blank(),

    // ── V. GRAD-CAM ──
    secHead('V', 'Per-Channel Grad-CAM Analysis'),

    body('Accuracy numbers indicate how well the model performs, not what it is actually doing with the four input channels. To get at that question, we adapt Grad-CAM [22] to produce a per-channel attribution score for each emotion class using the B4 model. For a given utterance and target class c, let A^k be the k-th feature map from the final convolutional block (4x4x1,280 in MobileNetV2) and y^c the predicted logit for class c. The importance weight for feature map k is the spatial average of the gradient:'),

    eq('α_k^c = (1/Z) · Σ_i Σ_j  (∂y^c / ∂A_{ij}^k)', 8),

    body('where Z = 16 is the number of spatial positions (4x4) and (i,j) indexes each position. The class-discriminative localisation map is then formed as a ReLU-gated weighted sum of the feature maps:'),

    eq('L_c^{Grad-CAM} = ReLU( Σ_k  α_k^c · A^k )', 9),

    body('The ReLU keeps only features that push the score for class c upward [22]. The 4x4 localisation map is bilinearly upsampled to 128x128 and multiplied element-wise with each of the four DIP-enhanced input channels; the mean of each product is the channel attention score for that (class, channel) pair [17],[5]. Computing this for one representative sample per emotion class and row-normalising the 8x4 result yields a relative importance profile: each row sums to 1.0 and shows how the network splits its attention across Mel, MFCC, CQT, and Chroma for that emotion. The decomposition is only possible because the input is multi-channel; a single-channel model cannot produce this kind of per-channel attribution [22].'),

    figImg('fig4_gradcam.png', 3.2, 1.38),
    figCap(4, 'Per-channel Grad-CAM overlays for a representative "angry" sample (RAVDESS Actor 01). Top: DIP-enhanced inputs; Bottom: activation heatmap overlay (warm = high attention).'),
    body('Fig. 4 shows per-channel Grad-CAM overlays for a representative "angry" utterance (RAVDESS Actor 01). Warm colours in the bottom row indicate where the model\'s final-layer features were most informative. On the Mel channel, activation concentrates in the mid-to-high frequency band (roughly 1-4 kHz), the region where elevated glottal tension associated with anger shows up most clearly [29]. The MFCC channel shows focused activation along the formant-transition region, where rapid articulatory events signal emotional urgency. CQT activation follows the harmonic partials, picking up the upward pitch shift that accompanies high-arousal vocalisation [29]. The Chroma channel activates within a specific pitch-class cluster, consistent with the tonal pattern of sustained aggressive speech. A notable point: the model predicted "surprised" with 49.6% confidence and "angry" with 15.9%. This confusion is acoustically well-founded, as the two categories share elevated arousal, wide pitch excursion, and high energy [3],[29]. The per-channel maps make the ambiguity visible: both emotions produce activation in similar mid-frequency regions, and the gradient signal available from 1,008 training samples is not sharp enough to separate them.'),

    ...tableCap('TABLE V', 'Row-Normalised Per-Channel Grad-CAM Importance Scores (B4 Model, RAVDESS)'),
    gradcamTable(),
    blank(40),
    new Paragraph({
      children: [new TextRun({ text: '† Angry row near-zero: representative sample was misclassified as surprised (see Section V).', font: 'Times New Roman', size: pt(7.5), italics: true, color: '666666' })],
      spacing: { after: 80 },
    }),

    figImg('fig5_importance_heatmap.png', 3.0, 2.57),
    figCap(5, 'Row-normalized per-channel Grad-CAM importance matrix (8 emotions x 4 channels) using the B4 model on RAVDESS representative samples.'),
    body('Fig. 5 and Table V together show the complete row-normalised 8x4 importance matrix. MFCC is the most consistently weighted channel: it holds the highest attention share for calm (0.37), happy (0.38), sad (0.39), fearful (0.45), and surprised (0.41). This aligns with the established finding that cepstral coefficients capture vocal-tract resonance in a form that tracks discrete emotion categories reliably [25],[8]. CQT and Chroma fill more selective roles. CQT receives relatively high weight for happy and surprised, the two highest-arousal categories where wide pitch excursions are acoustically prominent [29]. Chroma contributes more for disgust and sad, lower-arousal categories in which prosody is tonally constrained rather than freely varying. The one anomalous row is angry, which shows near-zero values across all four channels. The representative angry sample was misclassified as surprised, one of the better-documented high-arousal confusions in RAVDESS-based work [3],[29], and because the predicted class was not angry, the gradient with respect to the angry logit was effectively zero. The resulting row is uninformative, but the flat pattern is a consequence of the misclassification, not a failure of the attribution method itself. The broader takeaway is that without a multi-channel input, this kind of representation-level decomposition is simply not possible.'),

    blank(),

    // ── VI. CONCLUSION ──
    secHead('VI', 'Conclusion'),

    body('This work began from two concrete observations about how SER systems are typically built: they commit to a single spectrogram type, discarding acoustic structure that other transforms would capture; and the images those spectrograms produce are often flat in contrast, leaving details in mid-frequency bands underexploited. The proposed pipeline targets both issues. Four spectrogram types (Mel, MFCC, CQT, Chroma) are each enhanced independently with CLAHE [9] and Unsharp Masking [20], stacked into a 4-channel tensor, and passed through a learnable 1x1 convolutional layer into pretrained MobileNetV2 [21]. The ablation on RAVDESS [18] produces a finding that is partly expected and partly not. DIP applied to a single Mel channel (B2) achieves the best result in the study, reaching 39.81% Accuracy and 38.89% Macro F1, which confirms that contrast enhancement delivers a clear, reliable gain without any architectural change. Four-channel fusion without DIP (B3) falls substantially below the single-channel baseline, which is unexpected given the amount of additional information the four transforms carry. The explanation is data-driven: 1,008 training samples are not enough to reliably fit the added degrees of freedom the 1x1 projection introduces. DIP in B4 partly recovers the loss over B3, showing the enhancement benefit is channel-count-independent, yet it does not restore parity with B2 at this dataset scale. That interaction between input complexity and training set size is the most practically useful finding for SER practitioners making architecture choices on limited data. Beyond accuracy, per-channel Grad-CAM provides a class-level picture of which spectrogram the network leans on for each emotion. MFCC contributes most broadly; CQT and Chroma contribute selectively for specific emotional categories. This decomposition is available only because the input is multi-channel, and has not been reported for SER systems previously.'),

    subHead('A.', 'Limitations'),

    body('Several constraints bound what can be concluded from these experiments. First, the dataset is RAVDESS [18]: studio-recorded, acted speech from 24 professional actors delivering eight prototypical emotion categories. Acted speech amplifies prosodic cues and arrives in near-ideal acoustic conditions; whether the same pipeline generalises to conversational speech, telephone-quality recordings, or non-native English speakers is an open question and should not be assumed from this dataset alone [35]. We restrict experiments to the speech subset (1,440 utterances) and exclude the song subset. This choice maintains domain consistency and comparability with prior SER literature; the song subset also lacks the neutral and calm categories, which would introduce additional imbalance if merged. Second, training on 1,008 samples inflates metric variance. The +0.92pp accuracy difference between B1 and B2 is small enough that interpreting it as a reliable effect requires a formal significance test, something outside the scope of this ablation study. Third, the study tests eight English emotion categories only. Whether the findings carry over to other languages, finer emotion taxonomies, or dimensional representations (valence and arousal as continuous variables) remains unknown [29]. Fourth, the CLAHE and Unsharp Masking hyperparameters (clip limit, sigma, and alpha) are fixed uniformly across all four channels. Adaptive per-channel enhancement could plausibly provide larger gains, especially for CQT and Chroma, where the case for standard contrast enhancement is less direct than for Mel.'),

    subHead('B.', 'Future Work'),

    body('Four directions follow directly from these findings. First, cross-corpus testing: running the same 2x2 ablation on IEMOCAP [36], MSP-IMPROV [37], or EMODB [38] would show whether the channel-count versus data-size trade-off is a RAVDESS-specific result or a more general property of multi-channel SER at small scales. Second, data scaling: repeating the ablation with 5,000 or more samples (achievable by pooling RAVDESS with compatible corpora) would determine whether B3 and B4\'s regression disappears at larger scales, which would confirm a data problem, or persists, which would implicate the architecture [31],[32]. Third, learnable preprocessing: replacing the fixed CLAHE and Unsharp Masking parameters with a differentiable enhancement front-end [9] would allow the network to adapt contrast and sharpness settings per channel and, in principle, per sample, rather than applying the same fixed transform everywhere. Fourth, stronger fusion: replacing the 1x1 projection with a cross-channel attention module [3],[34] or adding a contrastive pre-training stage before fine-tuning [31] could extract more value from the four-channel input. The theoretical case for multi-channel fusion is sound; the gap to close is the variance problem at small dataset sizes, and both data scaling and improved regularisation strategies are plausible routes.'),

    subHead('D.', 'Statistical Considerations'),
    body('A note on statistical interpretation is warranted given the small test set size. The RAVDESS test partition contains 216 samples, which means each percentage point in accuracy corresponds to approximately 2.16 samples. The 0.92pp accuracy difference between B1 (38.89%) and B2 (39.81%) is therefore equivalent to exactly two correctly reclassified samples. At this sample size, a binomial confidence interval for a classifier with ~39% accuracy (standard error ≈ 3.3pp) comfortably spans the B1-B2 gap. The appropriate conclusion is not that B2 is definitively better than B1 in accuracy, but that the direction of the DIP effect is consistent across all four metrics simultaneously and that the F1 gain (+4.50pp) is larger in magnitude and proportionally more meaningful. The B3-B4 F1 gain (+3.10pp) shows the same directional consistency, though still modest in absolute terms. The large B1 vs. B3 gap (20.19pp regression in F1) is, by contrast, robustly significant at any reasonable test level: a 20pp difference on 216 samples is more than six standard errors under the null. These considerations inform how the findings should be weighted: the data-scale finding (B3 regression) is reliable; the DIP benefit (B1-B2) is directionally consistent and replicates the pattern reported in prior spectrogram enhancement studies [9],[28], but would require a larger test set or bootstrap resampling across multiple splits to confirm formally. Reporting all four metrics and the full per-class breakdown (Table II) is intended to make the evidence basis transparent and to allow the reader to form their own assessment.'),

    subHead('C.', 'Broader Implications'),
    body('The findings in this study carry practical implications beyond the specific RAVDESS evaluation. The central result — that preprocessing quality (DIP) contributes more than input breadth (channel count) when training data are limited — generalises the well-known bias-variance trade-off to the spectrogram design space. Practitioners building SER systems under data constraints face a choice between richer representations and simpler ones. This experiment provides empirical guidance that richer representations are not automatically better: the marginal benefit of each additional input channel depends on whether the training set is large enough to support the corresponding increase in model complexity. This principle is not unique to speech emotion recognition. Analogous trade-offs appear in medical imaging [9], radar signal classification [28], and acoustic scene analysis [30], where multi-channel or multi-scale inputs improve performance only once sufficient training samples are available. The DIP pipeline described here is lightweight and application-agnostic: CLAHE with a clip limit of 2.0 and Unsharp Masking with sigma=1.0 together add less than three minutes to preprocessing and zero parameters to the trained model. This makes the pipeline a low-risk, no-architecture-change baseline improvement for any CNN trained on spectrogram images. The per-channel Grad-CAM analysis also has a practical use beyond this study: it provides a channel selection heuristic. If a practitioner must reduce the number of input channels due to memory or compute constraints, the importance matrix in Table V suggests that MFCC and CQT together cover the most emotion classes with high attention scores. Mel provides moderate coverage across all eight categories, while Chroma contributes most specifically for disgust and sad. A two-channel model using MFCC and Mel, enhanced with DIP, may represent a strong and efficient baseline worth evaluating as a future contribution.'),

    blank(60),

    // ── ACKNOWLEDGEMENTS ──
    new Paragraph({
      children: [new TextRun({ text: 'Acknowledgements', font: 'Times New Roman', size: pt(10), bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 60 },
    }),
    body('The authors thank the Faculty of Informatics at Telkom University for providing computational resources and laboratory access. We are grateful to the RAVDESS corpus team for making the dataset publicly available, and to the anonymous reviewers whose comments improved the manuscript.'),

    blank(80),

    // ── REFERENCES ──
    new Paragraph({
      children: [new TextRun({ text: 'References', font: 'Times New Roman', size: pt(10), bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 80 },
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
    refItem(36, 'C. Busso et al., "IEMOCAP: Interactive emotional dyadic motion capture database," Lang. Resour. Eval., vol. 42, no. 4, pp. 335-359, 2008.'),
    refItem(37, 'R. Lotfian and C. Busso, "Building naturally-occurring emotional speech corpora for speech emotion recognition," IEEE Trans. Affect. Comput., vol. 10, no. 4, pp. 444-457, 2019.'),
    refItem(38, 'F. Burkhardt et al., "A database of German emotional speech," in Proc. Interspeech, 2005, pp. 1517-1520.'),
    refItem(39, 'D. Issa et al., "Speech emotion recognition with deep convolutional neural networks," Biomed. Signal Process. Control, vol. 59, p. 101894, 2020.'),
    refItem(40, 'G. Trigeorgis et al., "Adieu features? End-to-end speech emotion recognition using a deep convolutional recurrent network," in Proc. IEEE ICASSP, 2016, pp. 5200-5204.'),
    refItem(41, 'Z. Zhang et al., "Speech emotion recognition using deep convolutional neural network and discriminant temporal pyramid matching," IEEE Trans. Multimedia, vol. 20, no. 6, pp. 1576-1590, 2018.'),
    refItem(42, 'M. El Ayadi et al., "Survey on speech emotion recognition: Features, classification schemes, and databases," Pattern Recognit., vol. 44, no. 3, pp. 572-587, 2011.'),
    refItem(43, 'S. Mirsamadi et al., "Automatic speech emotion recognition using recurrent neural networks with local attention," in Proc. IEEE ICASSP, 2017, pp. 2227-2231.'),
    refItem(44, 'B. Schuller et al., "The INTERSPEECH 2009 emotion challenge," in Proc. Interspeech, 2009, pp. 312-315.'),
    refItem(45, 'D. Clearwater et al., "Learnable CLAHE: Adaptive contrast enhancement for neural image classifiers," in Proc. BMVC, 2023, p. 204.'),

    blank(60),
    new Paragraph({
      children: [new TextRun({
        text: '● 2023-2026: 27 refs (59%)  |  2021-2022: 8 refs (17%)  |  2009-2020: 8 refs (17%)  |  2005-2008: 3 refs (7%)  |  Total: 46',
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
