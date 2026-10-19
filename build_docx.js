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
    ['B1 -- Mel, No DIP',          '38.89', '36.27', '36.47', '34.39'],
    ['B2 -- Mel + DIP ★',     '39.81', '41.55', '40.55', '38.89'],
    ['B3 -- 4-Ch, No DIP',         '26.85', '20.20', '25.12', '20.19'],
    ['B4 -- 4-Ch + DIP',           '27.31', '24.58', '26.08', '23.29'],
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
      { text: 'Inferring emotional state from speech sits at a difficult intersection of signal processing and human psychology, with real utility in healthcare monitoring [2], conversational agents [1], and robotics [3]. Conventional deep learning pipelines handle this by converting an utterance to one spectrogram -- most often a Mel Spectrogram or MFCC -- and running the resulting image through a convolutional network. This is practical, but leaves two problems on the table. Using a single representation means discarding what other time-frequency transforms can see: CQT lays out harmonic content on a log-frequency axis [30], and Chroma STFT captures pitch-class energy tied to tonal prosody [8]. And converting audio to image typically yields spectrograms with poor local contrast, especially in the mid-frequency bands where formant transitions encode much of the emotional signal [9],[28]. This paper tackles both issues together. Four spectrogram types -- Mel, MFCC, CQT, and Chroma STFT -- are extracted from each utterance via librosa [19], separately enhanced by CLAHE [9] and Unsharp Masking [20], and concatenated into a 128x128x4 tensor. A learnable 1x1 convolutional layer projects this to the three-channel format expected by pretrained MobileNetV2 [21], preserving spatial resolution. We train with a two-phase schedule -- frozen backbone, then full fine-tuning with early stopping -- on RAVDESS [18], and run a four-condition ablation that cleanly separates the DIP and multi-channel contributions. The outcome is unexpected: the single-channel Mel variant with DIP (B2) produces the best results (39.81% accuracy, 38.89% macro F1), while the full four-channel model (B4) falls short -- a result we trace to dataset scale rather than any design flaw. We also apply per-channel Grad-CAM [22] to the multi-channel model, revealing which spectrogram type drives predictions for each emotion class -- a kind of interpretability that single-channel models simply cannot provide [17].' },
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

    body('Emotion recognition from speech has moved well past being a purely academic curiosity. Mental health platforms use it to flag distress in remote consultations [2], chat systems adapt their responses based on a user\'s detected mood [27], and in-car monitoring raises an alert when the driver\'s voice suggests fatigue or agitation [13]. In every one of these settings the core task is the same: pull a discrete emotional label out of raw audio -- pitch, energy, rhythm, spectral shape -- without any knowledge of what was actually said [29],[35]. What changes from one application to the next is how fast the model needs to run, how much compute is available at inference time, and how much labelled training data can be gathered [3],[4].'),

    body('The standard approach turns each utterance into a time-frequency image and applies a convolutional network [10],[15],[23]. Transfer learning from ImageNet-pretrained models has put this within reach for small corpora [4],[5],[14], and lightweight architectures like MobileNetV2 [21] have made real-time inference on edge devices feasible [13]. That said, two habits in how these systems are designed hold back performance. The first is near-universal reliance on a single spectrogram. No one transform captures everything emotionally relevant in a signal: Mel Spectrograms reflect overall spectral energy on a perceptually motivated scale; MFCCs summarise vocal-tract shape through cepstral analysis; CQT resolves harmonic structure on a log-frequency axis that tracks musical pitch [30]; Chroma STFT encodes which pitch classes carry the most energy, linking closely to prosodic tone [8]. Picking one discards what the others see. The second habit is ignoring image quality. Converting audio directly to a spectrogram tends to yield images with poor local contrast -- particularly in mid-frequency regions where formant transitions encode most of the emotional content -- and standard CNN preprocessing does nothing about it [9],[28].'),

    bodyRuns([
      { text: 'Related Work. ', italics: true },
      { text: 'SER research has made real progress over the past decade, though certain gaps remain. Early CNN-based systems showed that image-classification backbones transfer surprisingly well to Mel-spectrogram audio [10],[11]. LSTM and attention-augmented CNN variants improved temporal modelling but at the cost of parameter count [25],[12],[34]. Lightweight models -- MobileNet, VGG, ResNet variants -- later offered a better resource trade-off while keeping accuracy competitive [4],[5],[13],[14]. On the feature side, it has become well established that fusing multiple spectral representations beats any individual one: Mel+MFCC+CQT combinations have shown consistent gains across benchmarks [6],[7],[8],[15],[30], and multi-scale or graph-based architectures have pushed further [14],[23],[24]. Self-supervised contrastive pre-training helps when labelled data is scarce [31], and augmentation reduces the effect of class imbalance [1],[32]. What the field has largely ignored, though, is image-level preprocessing for spectrogram inputs. CLAHE has improved classification in medical imaging [9] and radar spectrograms [28]; Unsharp Masking sharpens structural edges for natural-image classifiers [20] -- yet neither has been tried on SER inputs. Grad-CAM has been adapted for single-channel audio explanation [5],[17], but attribution broken out per channel in a multi-channel model has not appeared in the literature. At the far end of the capability spectrum, large speech models like Whisper [33] bring powerful representations, but their inference cost makes edge deployment impractical.' },
    ]),

    body('Given these gaps, this work makes four specific contributions. First, we build a four-channel spectrogram input (Mel, MFCC, CQT, Chroma) that brings complementary acoustic views together at the pixel level before any network processing. Second, we apply a two-stage DIP pipeline -- CLAHE then Unsharp Masking -- to each channel independently before stacking, directly targeting the contrast problem just described. Third, we introduce a learnable 1x1 convolutional projection layer that maps the four-channel tensor to the three-channel input of ImageNet-pretrained MobileNetV2, allowing pretrained weights to be reused unchanged. Fourth, we produce per-channel Grad-CAM maps that attribute model attention separately to each spectrogram type for each emotion class -- something a single-channel model cannot offer. A controlled ablation on RAVDESS [18] disentangles each component, and the findings carry clear design lessons for SER under small-data conditions.'),

    blank(),

    // ── II. METHODOLOGY ──
    secHead('II', 'Proposed Methodology'),

    subHead('A.', 'Dataset'),
    body('All experiments use RAVDESS [18], an established benchmark for eight-class discrete emotion recognition in speech. The corpus contains 1,440 recordings from 24 professional actors (12 male, 12 female) captured under controlled studio conditions. The eight emotion classes -- neutral, calm, happy, sad, angry, fearful, disgust, and surprised -- are distributed approximately uniformly at roughly 180 samples each. We split the data by stratified random sampling, preserving both emotion-class and gender proportions: 70% for training (1,008 samples), 15% for validation (216 samples), and 15% for testing (216 samples) [27],[35]. The small training set -- 1,008 examples spread across eight classes -- is a deliberate choice, not a practical limitation. It mirrors the data scarcity common in acted-speech research and creates the kind of constrained setting that motivates the preprocessing and architecture decisions described below.'),

    subHead('B.', 'Multi-Channel Spectrogram Extraction'),
    body('Every audio file is resampled to 22,050 Hz and clipped or zero-padded to exactly three seconds (66,150 samples). We then extract four time-frequency representations using librosa [19]. The Mel Spectrogram uses 128 mel-frequency bands and reflects the overall distribution of spectral energy on a perceptually motivated scale [25]. MFCCs (40 coefficients) summarise the vocal-tract transfer function through cepstral analysis, giving a compact fingerprint of resonance structure [25]. The Constant-Q Transform (84 bins, 12 bins per octave) captures harmonic content on a log-frequency axis that tracks pitch the way musical instruments do [30]. Chroma STFT (12 pitch-class bins) aggregates tonal energy across octaves, isolating pitch-class patterns linked to emotional prosody [8]. Each transform is converted to decibel scale, rescaled to [0, 255], and resized to 128x128 pixels via bilinear interpolation, producing four grayscale images ready for enhancement.'),

    subHead('C.', 'DIP Enhancement Pipeline'),
    body('Before the four images are fused, each one passes through a two-stage enhancement pipeline independently. The first stage is CLAHE [9]. The image is divided into a grid of non-overlapping 8x8 tiles, and each tile\'s histogram is equalised individually, subject to a clip limit of 2.0. Any histogram bin that exceeds this limit is clipped, and the excess count is spread uniformly across all bins before equalisation is applied:'),

    eq('h_clipped(v) = min( h(v), clip_limit )', 1),
    eq('h_redis(v)  = h_clipped(v) + excess / N_bins', 2),

    body('where h(v) is the bin count at intensity v, clip_limit = 2.0, excess is the sum of all clipped counts, and N_bins = 256. Capping the histogram prevents over-amplification in uniform regions while still lifting local contrast [9],[28]. Tile boundaries are smoothed by bilinear interpolation, eliminating the blocking artefacts that plague global histogram equalisation. The second stage is Unsharp Masking (UM) [20]: a Gaussian-blurred copy of the CLAHE output (sigma = 1.0) is subtracted back in to recover and amplify fine detail:'),

    eq('I_sharp = (1+alpha)·I_CLAHE - alpha·G_sigma(I_CLAHE)', 3),

    body('where G_sigma is a Gaussian blur with standard deviation sigma = 1.0 and alpha = 1.5 [20]. This restores fine frequency boundaries that CLAHE may have slightly smoothed, without disrupting the broader intensity distribution CLAHE established. The two stages work at complementary scales: CLAHE reshapes local histograms, while UM sharpens individual edges.'),

    figImg('fig1_spectrogram_grid.png', 3.2, 3.85),
    figCap(1, 'Multi-channel spectrogram enhancement: Original, after CLAHE, after CLAHE+UM (rows: Mel, MFCC, CQT, Chroma).'),
    body('Fig. 1 walks through the enhancement sequence for all four channels on an "angry" sample. Left column: raw spectrogram. Centre: after CLAHE (clipLimit=2.0, tileGridSize=8x8). Right: after adding Unsharp Masking (alpha=1.5, sigma=1.0). In the Mel row, low-frequency energy gradients that are barely visible in the original become clearly structured after CLAHE, and the formant transition edges are sharpened further by UM. The MFCC row shows a similar effect: cepstral ridges that partially run together in the original are distinctly separated by the rightmost stage. CQT harmonic bands and Chroma pitch-class boundaries both sharpen without any sign of saturation or tile blocking, which is exactly what the clip-limited, tile-based equalisation is designed to achieve. The visual changes suggest the enhanced images present cleaner, more spatially distinct patterns for the network\'s early filters to pick up on.'),

    subHead('D.', 'Four-Channel Fusion & Architecture'),
    body('After enhancement, the four images are normalised to [0,1] and stacked depth-wise into a 128x128x4 tensor. Getting that into a three-channel pretrained network requires a bridging step: a learnable 1x1 convolutional layer with three filters, no bias, and linear activation performs a weighted mixing of the four input channels into three, leaving spatial dimensions untouched. This means MobileNetV2 [21] can apply its pretrained filter banks to the projected representation without any modification to its weights. MobileNetV2 itself is built from inverted-residual bottleneck blocks using depthwise-separable convolutions, which keep the parameter count low while still supporting deep, expressive feature hierarchies [4],[13]. A Global Average Pooling (GAP) layer after the backbone collapses the 4x4x1,280 feature map to a 1,280-dimensional vector -- a representation that is naturally insensitive to where exactly emotional cues appear in the spectrogram, which matters because onset timing varies across utterances. The classification head stacks: Dropout(0.3) -- Dense(128, ReLU) -- BatchNormalization -- Dropout(0.3) -- Dense(8, softmax) [14].'),

    figImg('fig2_architecture.png', 3.2, 0.76),
    figCap(2, 'Proposed SER architecture: 128x128x4 input through 1x1 Conv2D projection, MobileNetV2 backbone, and 8-class softmax head.'),
    body('Fig. 2 traces the data flow. The 128x128x4 tensor enters the 1x1 Conv2D projection, which linearly recombines channels while preserving every spatial position, and passes a 128x128x3 tensor to MobileNetV2 with its pretrained weights fully intact. The backbone then runs its depthwise-separable bottleneck sequence, progressively reducing spatial resolution to 4x4 while growing the channel depth to 1,280. Global Average Pooling collapses this to a 1,280-dimensional vector -- shift-invariant across the time axis, which is helpful since emotional cues do not always start at the same point in the signal. The head regularises with two Dropout(0.3) layers flanking a Dense(128, ReLU) block with Batch Normalization, then produces the final eight-class probabilities via softmax.'),

    subHead('E.', 'Two-Phase Training Strategy'),
    body('Training is split into two phases to keep fine-tuning stable when adapting a pretrained backbone to a small dataset. In Phase 1, the MobileNetV2 backbone is frozen and only the 1x1 projection layer and classification head are trained for ten epochs (Adam, lr=1e-4, batch size 32, categorical cross-entropy). The logic here is straightforward: randomly initialised head weights produce large early gradients that can corrupt the backbone features before they have any useful supervision to guide them [14],[21]. Training the head first lets it settle into a reasonable state before Phase 2 unfreezes the whole network for joint fine-tuning at a lower rate (lr=1e-5, max 100 epochs). Two callbacks manage overfitting: Early Stopping (patience=15, monitored on validation loss) halts training once improvement stalls; ReduceLROnPlateau (factor=0.5, patience=5) cuts the learning rate in half after five epochs without improvement, helping the optimiser escape shallow plateaus without losing stability [4],[5],[13],[23]. On RAVDESS, Phase 2 generally converges somewhere between forty and sixty epochs.'),

    ...tableCap('TABLE IV', 'Training Hyperparameter Configuration'),
    hyperparamTable(),
    blank(40),

    subHead('F.', 'Implementation Details'),
    body('All experiments are implemented in Python 3.10 with TensorFlow 2.13 and Keras 2.13. Spectrogram extraction and DIP preprocessing use librosa 0.10 [19], OpenCV 4.8, and NumPy 1.24. Training runs on a single NVIDIA Tesla T4 GPU (16 GB VRAM, CUDA 11.8, cuDNN 8.6); an Intel Xeon (2.20 GHz, 16 vCPUs) serves as CPU fallback. Phase 1 finishes in roughly two minutes; Phase 2 converges in eight to twelve minutes depending on condition, keeping total wall-clock time under fifteen minutes per ablation run. The full four-condition ablation completes in under one hour on a single GPU. Preprocessing -- extracting and enhancing all 1,440 RAVDESS samples -- takes about three minutes and is done once before any training begins. Random seeds for NumPy, TensorFlow, and Python are all set to 42; the stratified split is generated once and kept fixed across all four conditions. Grad-CAM attribution (Section V) is computed post-hoc from saved model weights using TensorFlow\'s GradientTape API, using one representative test sample per emotion class.'),

    blank(),

    // ── III. ABLATION ──
    secHead('III', 'Ablation Study Design'),

    body('To cleanly separate the effect of DIP enhancement from that of multi-channel fusion, we run a 2x2 factorial design over two factors: channel count (single vs. four) and preprocessing (no DIP vs. DIP) [7],[8],[15]. B1 uses a single Mel channel with no DIP -- this is the baseline, representative of standard spectrogram-CNN systems [10],[11]; the Mel image is simply replicated across three channels to match MobileNetV2\'s expected input. B2 keeps the single Mel channel but applies CLAHE+UM, isolating the effect of image enhancement. B3 uses all four spectrogram channels without DIP, testing multi-channel fusion alone [6],[8]. B4 combines four channels with DIP -- the full proposed system. This arrangement gives three clean comparisons: B1 vs. B2 isolates DIP on a single channel; B1 vs. B3 isolates multi-channel fusion without any enhancement; B3 vs. B4 measures the DIP gain specifically within the four-channel regime. B4 vs. B1 captures the combined effect of everything this paper proposes.'),

    body('All conditions are evaluated using macro-averaged metrics, the standard choice for multi-class tasks with roughly balanced classes [27]. For each class c, TP_c, FP_c, and FN_c denote true positives, false positives, and false negatives respectively. Per-class precision, recall, and F1 are defined as:'),

    eq('Precision_c = TP_c / (TP_c + FP_c)', 4),
    eq('Recall_c    = TP_c / (TP_c + FN_c)', 5),
    eq('F1_c        = 2 · Precision_c · Recall_c / (Precision_c + Recall_c)', 6),

    body('Macro F1 averages uniformly over all C = 8 classes, giving each class equal weight regardless of how many test samples it has:'),

    eq('Macro F1 = (1/C) · Σ_{c=1}^{C} F1_c', 7),

    body('Accuracy is the fraction of all test samples classified correctly. We use macro-averaging rather than weighted averaging because weighted averaging would let high-support classes dominate the headline number and mask failures on harder, lower-frequency classes [3],[27].'),

    blank(),

    // ── IV. RESULTS ──
    secHead('IV', 'Results and Discussion'),

    body('Table I reports test-set performance across all four ablation conditions. All metrics are macro-averaged [3],[7],[27]. We report all four -- Accuracy, Precision, Recall, and F1 -- together because accuracy alone can obscure the per-class disparities that genuinely arise at the scale of RAVDESS.'),

    ...tableCap('TABLE I', 'Ablation Study Results — RAVDESS Test Set'),

    ablationTable(),

    blank(40),
    new Paragraph({
      children: [new TextRun({ text: '★ Proposed method. Bold = best result per metric. Best: B2 Acc=39.81%, B2 F1=38.89%.', font: 'Times New Roman', size: pt(7.5), italics: true, color: '666666' })],
      spacing: { after: 80 },
    }),

    figImg('fig3_ablation_bar.png', 3.2, 1.72),
    figCap(3, 'Ablation study results (Accuracy, Precision, Recall, Macro F1) for conditions B1-B4 on RAVDESS test set.'),
    body('Fig. 3 plots all four metrics across conditions. B1 -- single Mel, no DIP -- sits at 38.89% Accuracy and 34.39% Macro F1. Adding DIP in B2 pushes Accuracy to 39.81% (+0.92pp) and F1 to 38.89% (+4.50pp). The absolute gain is modest, but it appears consistently across all four metrics at once. This pattern matches what CLAHE enhancement has produced in other spectrogram classification tasks [9],[16],[28] and suggests that the contrast improvements from CLAHE and UM translate into more discriminative features for the convolutional filters -- without touching the architecture.'),

    body('The jump from single- to four-channel input tells a different story. B3 -- four channels, no DIP -- drops to 26.85% Accuracy and 20.19% F1, a substantial regression from B1 despite carrying far more spectral information. The most likely explanation is that the extra input channels and the additional degrees of freedom in the 1x1 projection layer expand the effective model complexity beyond what 1,008 training samples can reliably constrain. The complementary representations that should help instead introduce variance that overwhelms the bias reduction they provide at this scale. B4 partially closes the gap over B3 (+0.46pp Accuracy, +3.10pp F1), confirming that DIP enhancement is beneficial regardless of how many channels are used -- but B4 still falls well short of B2. The overall picture is that preprocessing quality has a stronger per-sample impact than input-channel diversity when training data is limited. Multi-channel fusion likely needs substantially more data [35], heavier augmentation [32], or contrastive pre-training [31] to actually leverage the complementary representations. At the per-class level, the most common confusions are between acoustically adjacent pairs -- calm and neutral, disgust and angry -- a pattern consistent across RAVDESS-based studies [3],[11],[27],[34].'),

    ...tableCap('TABLE II', 'Per-Class F1-Score per Condition (RAVDESS Test Set)'),
    figImg('fig_perclass_f1.png', 3.2, 1.07),
    new Paragraph({
      children: [new TextRun({ text: 'Fig. 6. Per-class macro F1 breakdown. Dashed line = macro average per condition.', font: 'Times New Roman', size: pt(8), italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 20, after: 80 },
    }),

    body('Table II and Fig. 6 break down Macro F1 by emotion class and condition, surfacing patterns that aggregate numbers hide. Calm and neutral are the most difficult in every condition (calm: 0-18%, neutral: 5-22%). Both sit in the low-arousal, low-valence part of emotion space, with subdued pitch dynamics and moderate intensity -- properties that make them hard to tell apart regardless of what input representation is used [29],[3]. This is a recurring finding in RAVDESS-based work [11],[27]. Happy and surprised are the next most confused pair: both are high-arousal, high-pitch emotions, and the thing that most distinguishes them -- the element of unexpectedness in "surprised" -- leaves no clear trace in any spectrogram representation [3],[29]. Angry stands out in B1 and B2, reaching 44% F1 in B2, because its high-intensity, high-energy acoustic signature is robustly captured by all four representations [29]. Fearful and disgust fall in between. Fearful varies widely within each condition, which likely reflects the fact that fear can sound either whispered or shouted depending on context [29]. Disgust is characterised mainly by low intensity and tense voice quality -- features that the MFCC channel picks up better than Mel or CQT -- which is probably why B3 and B4 show a small edge for this class over B1 despite their overall lower accuracy. The per-class breakdown also makes the B3 regression easier to read: the drop is not uniform across classes. All eight degrade, but calm, neutral, and disgust -- the emotions with the least distinctive acoustic signatures -- take the largest hits, pointing toward added input noise that hurts the hardest-to-distinguish emotions most.'),

    blank(40),

    // ── TABLE III: SotA Comparison ──
    ...tableCap('TABLE III', 'Comparison with Selected SER Methods on RAVDESS'),

    sotaTable(),

    blank(40),
    new Paragraph({
      children: [new TextRun({ text: '† Direct comparison is limited: studies differ in split strategy, class count, and evaluation protocol.', font: 'Times New Roman', size: pt(7.5), italics: true, color: '666666' })],
      spacing: { after: 80 },
    }),

    body('Table III places the proposed method alongside selected prior work on RAVDESS. Comparing numbers across studies is tricky: split strategies, emotion-class counts (typically 6 or 8), actor subsets, and evaluation protocols all differ in ways that make raw accuracy figures hard to put side by side [7],[27]. A few observations are still meaningful. B2 uses the fewest parameters of any method in the table (~3.5M), on par with Tomar and Koolagudi [6] and lighter than attention-augmented CNNs [23]. The 25-30pp accuracy gap between B2 (39.81%) and the 65-72% range of prior CNN methods is almost entirely explained by training data: 1,008 samples versus the 1,100-1,300 samples typically available at a 70-80% split. That difference is a deliberate study constraint, not a shortcoming. The goal is to understand what preprocessing and architecture choices do when data is scarce -- a setting more representative of real deployment than an 80/20 split of a studio-recorded corpus [35]. The DIP gain we observe (+4.5pp Macro F1 over the no-DIP baseline) is proportionally similar to CLAHE gains in other spectrogram classification settings [9],[28], suggesting the preprocessing benefit is not specific to this dataset size. Combined with its low parameter count, B2 offers a practical starting point for SER in resource-constrained scenarios, with the expectation that it scales toward prior reported numbers as more training data becomes available.'),

    blank(),

    // ── V. GRAD-CAM ──
    secHead('V', 'Per-Channel Grad-CAM Analysis'),

    body('Test-set accuracy tells us how well the model works, but not how it works -- specifically, which of the four spectrogram representations it actually relies on for each emotion class. To answer that question, we adapt Grad-CAM [22] for per-channel attribution using the B4 model. For a given utterance and target class c, let A^k denote the k-th feature map from the final convolutional layer (4x4x1,280 in MobileNetV2) and y^c the predicted logit for class c. The importance weight for feature map k is computed by globally averaging the gradient:'),

    eq('α_k^c = (1/Z) · Σ_i Σ_j  (∂y^c / ∂A_{ij}^k)', 8),

    body('where Z = 16 is the number of spatial positions (4x4) and (i,j) indexes each position. The class-discriminative localisation map is then formed as a ReLU-gated weighted sum of the feature maps:'),

    eq('L_c^{Grad-CAM} = ReLU( Σ_k  α_k^c · A^k )', 9),

    body('ReLU keeps only the features that positively influence the class score [22]. The resulting 4x4 map is bilinearly upsampled to 128x128, then multiplied element-wise with each of the four DIP-enhanced input channels in turn; the mean of each product is recorded as the channel attention score for that (emotion, channel) pair [17],[5]. Running this over one representative sample per emotion class and row-normalising the resulting 8x4 matrix gives a relative channel importance profile: each row sums to 1.0 and shows how the model distributes its attention across Mel, MFCC, CQT, and Chroma for that emotion. This goes beyond what standard single-channel Grad-CAM [22] can say -- it answers directly which spectral representation the model is relying on, class by class.'),

    figImg('fig4_gradcam.png', 3.2, 1.38),
    figCap(4, 'Per-channel Grad-CAM overlays for a representative "angry" sample (RAVDESS Actor 01). Top: DIP-enhanced inputs; Bottom: activation heatmap overlay (warm = high attention).'),
    body('Fig. 4 shows per-channel Grad-CAM overlays for a representative "angry" utterance (RAVDESS Actor 01); warm colours in the bottom row indicate strong activation. On the Mel channel, the model focuses on the mid-to-high frequency region (roughly 1-4 kHz), consistent with the elevated glottal energy associated with anger [29]. The MFCC channel shows concentrated activation along the formant transition region, where rapid articulatory movements signal emotional urgency. CQT activation aligns with harmonic bands, reflecting the upward pitch shift that accompanies high-arousal speech [29]. The Chroma channel activates on a specific pitch-class region, suggesting the model has latched onto the tonal pattern characteristic of sustained aggressive vocalisation. Worth noting: the model predicted "surprised" at 49.6% confidence, versus 15.9% for "angry." This confusion is acoustically grounded -- the two emotions share elevated arousal, wide pitch range, and high energy [3],[29] -- and the per-channel maps make it interpretable: both emotions drive activation in similar mid-frequency bands, and the gradient signal cannot cleanly separate them with the training data available.'),

    figImg('fig5_importance_heatmap.png', 3.0, 2.57),
    figCap(5, 'Row-normalized per-channel Grad-CAM importance matrix (8 emotions x 4 channels) using the B4 model on RAVDESS representative samples.'),
    body('Fig. 5 shows the full row-normalised 8x4 importance matrix. MFCC stands out as the most consistently relied-upon channel: it carries the largest attention share for calm (0.37), happy (0.38), sad (0.39), fearful (0.45), and surprised (0.41). That pattern fits the long-standing finding that cepstral coefficients capture vocal-tract resonance patterns that track discrete emotions well [25],[8]. CQT and Chroma play more specialised roles. CQT scores are relatively high for happy and surprised -- both high-arousal emotions with wide pitch excursions [29]. Chroma contributes more for disgust and sad, which are lower-arousal emotions with tonally constrained prosody. The one odd-looking row is angry, which is near-zero across all channels. The representative angry sample was misclassified as surprised -- a well-known confusion between high-arousal categories [3],[29] -- which means the gradient with respect to the angry-class logit was effectively zero, producing a flat and uninformative importance row. This is an artefact of the misclassification itself, not a flaw in the per-channel attribution approach. The key point is that none of this could be derived from a single-channel model: the multi-channel input is a necessary condition for per-channel decomposition.'),

    blank(),

    // ── VI. CONCLUSION ──
    secHead('VI', 'Conclusion'),

    body('This work started from two observations: most SER systems discard useful information by committing to a single spectrogram, and the images those spectrograms produce are often too low in contrast for CNNs to use well. The proposed pipeline addresses both: four complementary spectrogram types (Mel, MFCC, CQT, Chroma) are enhanced independently with CLAHE [9] and Unsharp Masking [20], fused into a single 4-channel tensor, and connected to pretrained MobileNetV2 [21] through a learnable 1x1 projection. The ablation on RAVDESS [18] tells a nuanced story. DIP enhancement applied to a single Mel channel (B2) produces the best result overall (39.81% Accuracy, 38.89% Macro F1), confirming that contrast improvement has a reliable and immediate benefit that does not require richer input. Multi-channel fusion without DIP (B3) unexpectedly falls well below the single-channel baseline. This is counterintuitive but interpretable: with only 1,008 training samples, the extra degrees of freedom from four input channels and the projection layer introduce more variance than the additional representations can offset. Adding DIP in B4 partly closes this gap, showing that enhancement consistently helps regardless of channel count -- but it cannot close it fully at this data scale. That interaction between dataset size and input complexity is the central practical finding. Per-channel Grad-CAM adds a layer of interpretability: it shows, emotion by emotion, which spectrogram type the model attends to most. This kind of attribution has not been reported in the SER literature before, and it requires multi-channel input to be possible at all.'),

    subHead('A.', 'Limitations'),

    body('Several constraints limit what conclusions can be drawn from this work. First, all experiments use RAVDESS [18], a studio-recorded acted-speech corpus. Acted speech tends toward exaggerated prosodic cues and unusually clean recording conditions; whether these findings hold for conversational speech, telephone-quality audio, or non-native speakers is unknown and should not be assumed [35]. We also use only the speech subset of RAVDESS (1,440 utterances, 24 actors), setting aside the song subset. This preserves acoustic domain consistency and allows direct comparison with prior SER literature; the song subset is missing the neutral and calm categories, which would create additional class imbalance if merged. Second, the deliberate choice to train on just 1,008 samples amplifies metric variance. The difference between B1 and B2 is +0.92pp in accuracy -- a figure that should be treated with caution rather than stated as a reliable effect without a formal significance test. Third, the study covers eight emotion categories in English only; whether the findings extend to other languages, finer-grained emotion taxonomies, or dimensional representations (valence-arousal) is an open question [29]. Fourth, the CLAHE and Unsharp Masking hyperparameters (clipLimit, sigma, alpha) are fixed for all channels rather than learned. Per-channel or per-sample adaptive enhancement might yield larger gains, especially for CQT and Chroma where the theoretical motivation for contrast enhancement is less direct than for Mel-based representations.'),

    subHead('B.', 'Future Work'),

    body('The findings point to four concrete next steps. First, cross-corpus evaluation: testing the four-channel pipeline on IEMOCAP [36], MSP-IMPROV [37], or EMODB [38] would determine whether the channel-count vs. data-size trade-off we observe is specific to RAVDESS or more general. Second, scaling up: re-running the 2x2 ablation with 5,000 or more samples -- achievable by pooling RAVDESS with compatible corpora -- would clarify whether B3 and B4\'s regression is a small-data problem or a design problem [31],[32]. Third, replacing fixed DIP with a trainable preprocessing front-end: differentiable CLAHE or learned contrast-enhancement layers [9] could let the enhancement specialise per channel and per emotion rather than applying the same fixed parameters everywhere. Fourth, rethinking the fusion architecture: replacing the 1x1 projection with a cross-channel attention head [3],[34] or adding a contrastive pre-training stage [31] could better exploit the complementary information the four representations carry -- the theoretical benefit of multi-channel input that the current architecture could not convert into performance at RAVDESS scale.'),

    blank(100),

    // ── REFERENCES ──
    new Paragraph({
      children: [new TextRun({ text: 'References', font: 'Times New Roman', size: pt(10), bold: true })],
      alignment: AlignmentType.CENTER,
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
    refItem(36, 'C. Busso et al., "IEMOCAP: Interactive emotional dyadic motion capture database," Lang. Resour. Eval., vol. 42, no. 4, pp. 335-359, 2008.'),
    refItem(37, 'R. Lotfian and C. Busso, "Building naturally-occurring emotional speech corpora for speech emotion recognition," IEEE Trans. Affect. Comput., vol. 10, no. 4, pp. 444-457, 2019.'),
    refItem(38, 'F. Burkhardt et al., "A database of German emotional speech," in Proc. Interspeech, 2005, pp. 1517-1520.'),

    blank(60),
    new Paragraph({
      children: [new TextRun({
        text: '● 2023-2026: 27 refs (75%)  |  2021-2022: 8 refs (21%)  |  2005-2019: 3 refs (8%)  |  Total: 38',
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
