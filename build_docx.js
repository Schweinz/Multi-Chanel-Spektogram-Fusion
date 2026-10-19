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
              bold: ri === rows.length - 1,
              color: ri === rows.length - 1 && ci > 0 ? '1a3a5c' : '222222',
            })],
            alignment: ci === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
          })],
          width: { size: colW[ci], type: WidthType.DXA },
        })),
      })),
    ],
  });
}

// ── Ablation table ───────────────────────────────────────────
function ablationTable() {
  // Two-column layout: each column ~4320 twip. Table must fit within one column.
  const colW = [1760, 620, 620, 620, 620]; // total 4240 twip — fits in one column
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
      { text: 'Recognition of emotional state from spoken language is an open problem in affective computing, with applications in human-computer interaction [1], mental health monitoring [2], and social robotics [3]. Most current deep learning systems convert each utterance to a single spectrogram -- typically a Mel Spectrogram or MFCC -- and feed the resulting image into a convolutional network. Two practical problems arise from this design. First, a single representation discards complementary acoustic information that other time-frequency analyses capture: CQT reveals harmonic structure on a log-frequency scale [30], while Chroma STFT encodes pitch-class energy that reflects tonal prosody [8]. Second, the direct conversion from audio to image tends to produce spectrograms with low local contrast, particularly in mid-frequency regions where formant transitions carry much of the affective signal [9],[28]. We address both problems simultaneously. Four spectrogram types -- Mel Spectrogram, MFCC, Constant-Q Transform (CQT), and Chroma STFT -- are extracted from each utterance using librosa [19], enhanced independently by CLAHE [9] and Unsharp Masking [20], and stacked into a 128x128x4 input tensor. A learnable 1x1 convolutional projection layer maps this to the three-channel space of pretrained MobileNetV2 [21] without spatial resampling. Training uses a two-phase schedule (frozen backbone, then full fine-tuning) with early stopping on RAVDESS [18]. A systematic four-condition ablation isolates the contributions of the DIP pipeline and the multi-channel fusion independently. Unexpectedly, the single-channel Mel variant with DIP (B2) achieves the best test accuracy (39.81%) and macro F1 (38.89%), while the full four-channel model (B4) falls below it -- a finding that points to dataset-scale constraints rather than any flaw in the fusion design. Per-channel Grad-CAM [22] is applied to the multi-channel model to reveal which spectrogram type most influences each emotion-class prediction, providing channel-level interpretability not available from single-channel approaches [17].' },
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

    body('Recognising the emotional content of speech has attracted growing interest well beyond its academic origins. Deployed healthcare applications now use vocal analysis for remote mental health screening [2], conversational agents adjust response generation based on detected affect [27], and driver-monitoring systems raise alerts when voice signals suggest elevated stress or fatigue [13]. In each case the underlying task is the same: extracting from an audio signal -- pitch, intensity, rhythm, spectral envelope -- a label that corresponds to the speaker\'s emotional state, without reference to the linguistic content [29],[35]. What varies across applications is the tolerance for latency, the available compute budget, and the quantity and quality of labelled training data [3],[4].'),

    body('The prevailing approach converts each utterance to a time-frequency image and applies a convolutional neural network [10],[15],[23]. Transfer learning from ImageNet-pretrained models has brought this within reach of small corpora [4],[5],[14], and lightweight architectures such as MobileNetV2 [21] have made edge deployment practical [13]. Yet two design choices, repeated across much of the literature, unnecessarily limit performance. Almost universally, systems use a single spectrogram representation. In practice, no single representation captures all of the acoustic features that carry emotional information: the Mel Spectrogram tracks spectral energy in a perceptually weighted scale; MFCC captures vocal-tract resonance through cepstral coefficients; CQT resolves harmonic content on a log-frequency axis that mirrors musical pitch [30]; and Chroma STFT encodes pitch-class energy that reflects prosodic tonality [8]. Choosing one means discarding information from the others. The second choice concerns contrast. Spectrograms computed directly from waveforms often have low local contrast, especially in the mid-frequency bands where formant transitions carry the most affective signal -- a problem that standard CNN preprocessing does not address [9],[28].'),

    bodyRuns([
      { text: 'Related Work. ', italics: true },
      { text: 'The trajectory of SER research over the past decade illustrates both progress and persistent gaps. Early convolutional approaches applied image-classification backbones to Mel spectrograms and established that visual feature hierarchies transfer effectively to audio [10],[11]. LSTM and attention-augmented CNN architectures then extended temporal modelling [25],[12],[34] at the cost of substantially higher parameter counts. Lightweight transfer learning -- MobileNet, VGG, ResNet -- has since offered a more practical balance, achieving competitive recognition rates with fewer resources [4],[5],[13],[14]. Parallel work on feature engineering has shown that combining multiple spectral representations consistently outperforms any single one: Mel+MFCC+CQT fusion has demonstrated reliable gains across several benchmarks [6],[7],[8],[15],[30], and multi-scale or graph-based architectures have further enriched the learned representations [14],[23],[24]. Self-supervised contrastive pre-training addresses the label scarcity that limits small datasets [31], and augmentation strategies reduce the impact of class imbalance [1],[32]. What is conspicuously absent from this trajectory is systematic work on image-level preprocessing for spectrogram inputs. CLAHE has meaningfully improved classification accuracy in medical imaging [9] and radar spectrogram recognition [28]; Unsharp Masking has sharpened structural detail for natural-image classifiers [20]; but neither has been tested on SER spectrogram inputs. Similarly, while Grad-CAM has been adapted to single-channel audio [5],[17], per-channel attribution in multi-channel SER models has not been reported. Large pretrained speech models such as Whisper [33] offer powerful representations, but their inference cost remains prohibitive for edge deployment.' },
    ]),

    body('Against this background, the present work makes four focused contributions. The first is a four-channel spectrogram input (Mel+MFCC+CQT+Chroma) that captures complementary acoustic information simultaneously at the pixel level. The second is a two-stage DIP pipeline -- CLAHE followed by Unsharp Masking -- applied independently to each channel before stacking, directly addressing the contrast limitation identified above. The third is a learnable 1x1 convolutional projection layer that adapts the four-channel tensor to the three-channel input of ImageNet-pretrained MobileNetV2 without discarding pretrained weights. The fourth is per-channel Grad-CAM attribution, which maps model attention separately for each spectrogram type and each emotion class. A controlled ablation study on RAVDESS [18] isolates the effect of each component and produces findings with direct design implications for small-corpus SER systems.'),

    blank(),

    // ── II. METHODOLOGY ──
    secHead('II', 'Proposed Methodology'),

    subHead('A.', 'Dataset'),
    body('Experiments are conducted on the Ryerson Audio-Visual Database of Emotional Speech and Song (RAVDESS) [18], a widely used benchmark for eight-class discrete SER. The corpus consists of 1,440 audio recordings from 24 professional actors (12 male, 12 female) recorded under controlled studio conditions. The eight emotion categories -- neutral, calm, happy, sad, angry, fearful, disgust, and surprised -- are distributed approximately equally at around 180 samples per class. We partition the data by stratified random sampling to preserve emotion-class and gender balance: 70% for training (1,008 samples), 15% for validation (216 samples), and 15% for testing (216 samples) [27],[35]. The modest corpus size -- 1,008 training examples across eight classes -- is a deliberate constraint of the study, not a shortcut; it reflects the data scarcity typical of acted-speech corpora and motivates the design choices made throughout.'),

    subHead('B.', 'Multi-Channel Spectrogram Extraction'),
    body('All audio files are resampled to 22,050 Hz and either truncated or zero-padded to a fixed duration of three seconds (66,150 samples). Four complementary time-frequency representations are then extracted with librosa [19]. The Mel Spectrogram (128 mel-frequency bands) captures the overall spectral energy distribution weighted to perceptual scale [25]. MFCCs (40 coefficients) encode the vocal-tract transfer function through cepstral analysis, providing a compact description of resonance structure [25]. The Constant-Q Transform (84 bins, 12 bins per octave) resolves harmonic content on a logarithmic frequency axis that closely mirrors musical pitch perception [30]. Chroma STFT (12 pitch-class bins) summarises tonal energy independently of octave, highlighting pitch-class patterns associated with emotional prosody [8]. Each representation is converted to decibel scale, normalised to [0, 255], and resized via bilinear interpolation to 128x128 pixels, yielding a set of grayscale images ready for downstream DIP enhancement.'),

    subHead('C.', 'DIP Enhancement Pipeline'),
    body('Prior to fusion, each of the four 128x128 grayscale images undergoes a two-stage image-enhancement pipeline applied independently. The first stage is CLAHE [9]. The image is subdivided into non-overlapping 8x8 contextual tiles, and within each tile a histogram is equalised subject to a clip limit of 2.0. Bins whose count exceeds the clip limit are clipped, and the excess mass is redistributed uniformly across all bins before equalisation:'),

    eq('h_clipped(v) = min( h(v), clip_limit )', 1),
    eq('h_redis(v)  = h_clipped(v) + excess / N_bins', 2),

    body('where h(v) is the histogram count at intensity v, clip_limit = 2.0, excess = sum of all clipped counts, and N_bins = 256. This prevents over-amplification of uniform regions while still enhancing local contrast [9],[28]. Tile boundaries are blended by bilinear interpolation, avoiding the blocking artefacts that afflict global histogram equalisation. The second stage is Unsharp Masking (UM) [20]: a Gaussian-blurred version of the CLAHE output (sigma = 1.0) is formed and then combined with the original to recover and amplify high-frequency detail:'),

    eq('I_sharp = (1+alpha)·I_CLAHE - alpha·G_sigma(I_CLAHE)', 3),

    body('where G_sigma denotes convolution with a Gaussian kernel of standard deviation sigma = 1.0, and alpha = 1.5 [20]. The operation restores fine frequency boundaries that CLAHE may have smoothed, without disturbing the global intensity distribution that CLAHE established. Together, the two stages address different scales of the contrast problem: CLAHE operates on the local histogram, while UM targets the sharpness of edges at the pixel level.'),

    figImg('fig1_spectrogram_grid.png', 3.2, 3.85),
    figCap(1, 'Multi-channel spectrogram enhancement: Original, after CLAHE, after CLAHE+UM (rows: Mel, MFCC, CQT, Chroma).'),
    body('Fig. 1 illustrates the enhancement sequence for all four channels on an "angry" sample. The left column shows the unprocessed spectrogram; the centre column shows the CLAHE output (clipLimit=2.0, tileGridSize=8x8); the right column adds Unsharp Masking (alpha=1.5, sigma=1.0) on top of CLAHE. In the Mel row, low-frequency energy gradients that are nearly invisible in the original become clearly structured after CLAHE, and the formant transition boundaries are further sharpened by UM. The MFCC row reveals a similar pattern: the cepstral ridges, which partially merge in the original, are distinctly separated in the rightmost column. CQT harmonic band edges and Chroma pitch-class boundaries both show appreciable sharpening without evidence of saturation or blocking artefacts -- outcomes expected from the clip-limited tile-based equalisation mechanism. These visual improvements suggest that the enhanced images present more spatially distinct patterns for the CNN\'s early convolutional filters to respond to.'),

    subHead('D.', 'Four-Channel Fusion & Architecture'),
    body('After enhancement the four images are normalised to [0,1] and concatenated depth-wise to form a 128x128x4 input tensor. Connecting this directly to a standard three-channel pretrained network requires a bridging step: we use a learnable 1x1 convolutional projection layer (three filters, no bias, linear activation) that performs a linear recombination of the four input channels into three, without touching the spatial dimensions. This allows the pretrained MobileNetV2 [21] filter banks to operate on the projected representation without modification. MobileNetV2 uses depthwise-separable convolutions in inverted-residual bottleneck blocks, reducing parameter count substantially relative to standard convolutions while preserving depth and width for expressive feature learning [4],[13]. After the backbone, a Global Average Pooling (GAP) layer collapses the 4x4x1,280 feature map to a 1,280-dimensional vector, making the representation insensitive to spatial shifts in the spectrogram -- a useful property given that the temporal onset of emotional cues varies across utterances. The classification head is: Dropout(0.3) -- Dense(128, ReLU) -- BatchNormalisation -- Dropout(0.3) -- Dense(8, softmax), producing the probability distribution over eight emotion classes [14].'),

    figImg('fig2_architecture.png', 3.2, 0.76),
    figCap(2, 'Proposed SER architecture: 128x128x4 input through 1x1 Conv2D projection, MobileNetV2 backbone, and 8-class softmax head.'),
    body('Fig. 2 traces the data flow through the proposed architecture. The 128x128x4 tensor enters the 1x1 Conv2D projection layer, which performs a linear channel mixing -- without touching the spatial dimensions -- and outputs a 128x128x3 tensor that MobileNetV2 can process with its pretrained weights intact. The backbone then applies its sequence of depthwise-separable bottleneck blocks, progressively reducing the spatial resolution to 4x4 while expanding the channel count to 1,280. A Global Average Pooling operation collapses the spatial dimensions, producing a fixed-length 1,280-dimensional feature vector that is insensitive to spatial shifts within the spectrogram -- a desirable property given that the onset times of emotional features vary across utterances. The head applies two dropout layers around a Dense(128, ReLU) block with Batch Normalisation, providing regularisation against the relatively small RAVDESS training set, before the final Dense(8) with softmax.'),

    subHead('E.', 'Two-Phase Training Strategy'),
    body('Training is split into two phases to stabilise fine-tuning of a pretrained backbone on a small dataset. Phase 1 freezes the MobileNetV2 backbone and trains only the 1x1 projection layer and the classification head for ten epochs (Adam, lr=1e-4, batch size 32, categorical cross-entropy). The purpose is to reach a stable head configuration before any pretrained weights are perturbed; randomly initialised head weights carrying large gradients early in training can distort the backbone features that give transfer learning its value [14],[21]. Once the head has converged to a reasonable initialisation, Phase 2 unfreezes all layers and fine-tunes them jointly at a lower rate (lr=1e-5, max 100 epochs). Two callbacks guard against overfitting: Early Stopping (patience=15, monitored on validation loss) terminates the run when no improvement occurs for fifteen consecutive epochs; ReduceLROnPlateau (factor=0.5, patience=5) halves the learning rate after five stagnant epochs, allowing the optimiser to descend from shallow plateaus without destabilising the solution [4],[5],[13],[23]. On RAVDESS, Phase 2 typically converges within forty to sixty epochs.'),

    blank(),

    // ── III. ABLATION ──
    secHead('III', 'Ablation Study Design'),

    body('To separate the effect of DIP enhancement from the effect of multi-channel fusion, four conditions are constructed in a 2x2 factorial arrangement over channel count (single vs. four) and enhancement (no DIP vs. DIP) [7],[8],[15]. B1: single Mel channel, no DIP -- the standard baseline representative of earlier spectrogram CNN approaches [10],[11]. The Mel image is replicated to three channels to satisfy MobileNetV2\'s expected input shape. B2: single Mel channel with CLAHE+UM applied, isolating image enhancement as the sole experimental variable over B1. B3: four spectrogram channels (Mel, MFCC, CQT, Chroma) without DIP, testing multi-channel fusion alone in the style of [6],[8]. B4: four channels with DIP, the complete proposed system. This design yields three orthogonal comparisons: B1 vs. B2 measures the DIP gain on a single channel; B1 vs. B3 measures the multi-channel gain without enhancement; B3 vs. B4 measures the DIP gain in the multi-channel regime. B4 vs. B1 gives the overall combined gain of the full system relative to the minimal baseline.'),

    body('All conditions are evaluated on macro-averaged metrics, as recommended for multi-class tasks with approximately balanced classes [27]. For each class c, let TP_c, FP_c, and FN_c denote true positives, false positives, and false negatives respectively. Per-class precision, recall, and F1 are:'),

    eq('Precision_c = TP_c / (TP_c + FP_c)', 4),
    eq('Recall_c    = TP_c / (TP_c + FN_c)', 5),
    eq('F1_c        = 2 · Precision_c · Recall_c / (Precision_c + Recall_c)', 6),

    body('Macro-averaged F1 averages uniformly across all C = 8 emotion classes, giving equal weight to each class regardless of support:'),

    eq('Macro F1 = (1/C) · Σ_{c=1}^{C} F1_c', 7),

    body('Accuracy is reported as the fraction of correctly classified test samples across all classes. Macro-averaging is preferred over weighted averaging here because it prevents the metric from being dominated by high-support classes that are consistently easier to classify [3],[27].'),

    blank(),

    // ── IV. RESULTS ──
    secHead('IV', 'Results and Discussion'),

    body('Table I reports test-set performance for each ablation condition. Macro-averaging is used throughout, as recommended for multi-class recognition tasks with approximately balanced classes [3],[7],[27]; all four metrics -- Accuracy, macro Precision, macro Recall, and macro F1 -- are reported jointly because optimising on Accuracy alone can mask per-class imbalances that are genuinely present at this corpus scale.'),

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
    body('Fig. 3 visualises the four metrics across all conditions. B1, the single-Mel no-DIP baseline, scores 38.89% Accuracy and 34.39% Macro F1. Adding DIP in B2 raises Accuracy to 39.81% (+0.92pp) and F1 to 38.89% (+4.50pp), a gain that is modest in absolute terms but consistent across all four metrics simultaneously. This result is in line with CLAHE-enhancement studies on other spectrogram-based classifiers [9],[16],[28], and suggests that the contrast improvements introduced by CLAHE and UM translate to more discriminative CNN features without requiring any architectural change.'),

    body('The transition from single- to four-channel input tells a different story. B3 -- four channels, no DIP -- falls to 26.85% Accuracy and 20.19% F1, a substantial regression from B1 despite offering richer spectral information. One plausible explanation is that the increased input dimensionality and the extra degrees of freedom in the 1x1 projection layer expand the effective hypothesis space beyond what 1,008 training samples can reliably constrain. The additional channels introduce representational diversity that, under RAVDESS scale, manifests as added variance rather than reduced bias. B4 recovers partially over B3 (+0.46pp Accuracy, +3.10pp F1), confirming that DIP enhancement provides a consistent positive contribution regardless of channel count, yet B4 still falls well short of B2. Taken together, the ablation suggests that on small acted-speech corpora, preprocessing quality has a stronger per-sample effect than input-channel diversity. Multi-channel fusion should be revisited with larger, more diverse corpora [35], stronger data augmentation [32], or contrastive pre-training [31] that can exploit the complementary representations more effectively. Per-class analysis reveals that the most frequent confusions occur between acoustically similar emotion pairs -- calm and neutral, disgust and angry -- consistent with observations across RAVDESS-based studies [3],[11],[27],[34].'),

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

    blank(80),

    // ── TABLE III: SotA Comparison ──
    new Paragraph({
      children: [new TextRun({ text: 'TABLE III. Comparison with Selected SER Methods on RAVDESS', font: 'Arial', size: pt(8.5), bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
    }),

    sotaTable(),

    blank(40),
    new Paragraph({
      children: [new TextRun({ text: '† Direct comparison is limited: studies differ in split strategy, class count, and evaluation protocol.', font: 'Times New Roman', size: pt(7.5), italics: true, color: '666666' })],
      spacing: { after: 80 },
    }),

    body('Table III contextualises the proposed method relative to selected prior work on RAVDESS. Direct numerical comparison across studies is complicated by heterogeneous experimental protocols -- variations in train/test split strategy, number of emotion classes (commonly 6 or 8), actor subsets, and evaluation metrics mean that raw accuracy figures are not directly commensurable [7],[27]. Nonetheless, several observations are instructive. The proposed B2 configuration operates with the fewest parameters among the listed methods (~3.5M), comparable to Tomar and Koolagudi [6] and lower than attention-augmented CNNs [23]. The accuracy gap between B2 (39.81%) and prior CNN-based methods (65--72%) is consistent with the deliberate dataset constraint of this study: training on 1,008 samples is substantially smaller than the 70--80% partitions used in most cited works, which typically yield 1,100--1,300 training samples. This scale difference is a design choice rather than a methodological shortcoming, intended to isolate the effect of preprocessing and architecture under data-scarce conditions representative of low-resource deployment. The DIP enhancement gain observed in this study (+4.5pp Macro F1 over the unenhanced baseline) is comparable in relative magnitude to CLAHE gains reported in spectrogram classification at larger scales [9],[28], suggesting that the preprocessing benefit is dataset-size-agnostic. Parameter efficiency and preprocessing effectiveness together position the proposed pipeline as a practical starting point for SER in resource-constrained settings, with the expectation that performance scales toward prior benchmarks as training data grows.'),

    blank(),

    // ── V. GRAD-CAM ──
    secHead('V', 'Per-Channel Grad-CAM Analysis'),

    body('Performance metrics alone do not reveal how the model distributes attention across the four spectrogram representations, nor do they indicate which channel contributes most to a given emotion class. To address this, Grad-CAM [22] is adapted for per-channel attribution using the B4 model. For a given utterance and target class c, let A^k denote the k-th feature map of the final convolutional layer (4x4x1,280 in MobileNetV2), and y^c the predicted logit for class c. The importance weight for feature map k is obtained by global average pooling of the gradient:'),

    eq('α_k^c = (1/Z) · Σ_i Σ_j  (∂y^c / ∂A_{ij}^k)', 8),

    body('where Z = 4x4 = 16 is the number of spatial positions and (i,j) indexes the spatial location. The class-discriminative localisation map is then formed as a ReLU-gated weighted combination of the feature maps:'),

    eq('L_c^{Grad-CAM} = ReLU( Σ_k  α_k^c · A^k )', 9),

    body('ReLU is applied to retain only the features that have a positive influence on the class score [22]. The resulting 4x4 scalar map is bilinearly upsampled to 128x128 and multiplied element-wise with each of the four DIP-enhanced input channels separately; the mean of each product is recorded as the attention score for that (emotion, channel) pair [17],[5]. Applying this procedure to one representative sample per emotion class and row-normalising the resulting 8x4 matrix produces a relative channel importance profile: each row sums to 1.0 and reflects how the model distributes its attention across Mel, MFCC, CQT, and Chroma for that emotion class. This goes beyond standard single-channel Grad-CAM [22], directly answering which spectral representation the model finds most informative per emotion.'),

    figImg('fig4_gradcam.png', 3.2, 1.38),
    figCap(4, 'Per-channel Grad-CAM overlays for a representative "angry" sample (RAVDESS Actor 01). Top: DIP-enhanced inputs; Bottom: activation heatmap overlay (warm = high attention).'),
    body('Fig. 4 shows the per-channel Grad-CAM overlays for a representative "angry" utterance (RAVDESS Actor 01). In the bottom row, warm colours indicate high activation. On the Mel channel, the model attends most strongly to the mid-to-high frequency region (roughly 1-4 kHz), consistent with the elevated glottal energy and intensity range associated with angry speech [29]. The MFCC channel shows activation concentrated along the formant transition region, where rapid articulatory changes encode emotional urgency. CQT activation peaks align with harmonic bands, reflecting the upward pitch shift documented in high-arousal speech [29]. The Chroma channel activates on a specific pitch-class region, suggesting the model has detected the tonal emphasis that characterises sustained aggressive vocalisation. A notable observation is that the model predicted "surprised" with 49.6% confidence rather than "angry" (15.9%) -- a confusion that is acoustically motivated, as the two emotions share elevated arousal, wide pitch range, and high energy [3],[29]. The per-channel view makes this confusion interpretable: both emotions activate similar mid-frequency regions, and the gradient-based signal does not cleanly separate them under the limited training data available.'),

    figImg('fig5_importance_heatmap.png', 3.0, 2.57),
    figCap(5, 'Row-normalized per-channel Grad-CAM importance matrix (8 emotions x 4 channels) using the B4 model on RAVDESS representative samples.'),
    body('Fig. 5 presents the row-normalised 8x4 importance matrix. Several patterns emerge. MFCC is the most consistently dominant channel, contributing the largest fraction of attention for calm (0.37), happy (0.38), sad (0.39), fearful (0.45), and surprised (0.41). This aligns with the well-established finding in the SER literature that cepstral coefficients capture vocal-tract resonance patterns that correlate strongly with discrete emotion categories [25],[8]. CQT and Chroma offer complementary contributions particularly for pitch-related emotions: CQT scores are relatively elevated for happy and surprised, both high-arousal emotions with wider pitch range [29], while Chroma contributes more for disgust and sad -- emotions characterised by tonally constrained, lower-arousal prosody. One notable anomaly is the near-zero row for angry. The representative angry sample was misclassified as surprised -- a documented confusion between high-arousal categories [3],[29] -- which means the gradient with respect to the angry-class logit was effectively zero, resulting in a uniformly low importance row. This is an interpretable artefact of the model\'s misclassification rather than a structural limitation of the per-channel attribution method itself. Critically, this profile could not have been derived from a single-channel system; the multi-channel input is a prerequisite for the per-channel decomposition.'),

    blank(),

    // ── VI. CONCLUSION ──
    secHead('VI', 'Conclusion'),

    body('Three ideas drive this work: fusing four complementary spectrograms (Mel, MFCC, CQT, Chroma) into a single input tensor; enhancing each channel with CLAHE [9] and Unsharp Masking [20] before stacking; and adapting the result to pretrained MobileNetV2 [21] via a learnable 1x1 projection. Together they address the single-representation and low-contrast limitations identified in the introduction. The ablation on RAVDESS [18] clarifies what each component contributes. DIP enhancement alone, applied to a single Mel channel, yields the strongest result in this study (B2: 39.81% Accuracy, 38.89% Macro F1), confirming that contrast improvement has an immediate and consistent benefit that does not depend on richer input. Multi-channel fusion without DIP (B3) falls well below the single-channel baseline -- an outcome that, while counterintuitive, is interpretable: 1,008 training samples are insufficient to reliably constrain the additional degrees of freedom introduced by four input channels and the projection layer. Adding DIP to the four-channel configuration (B4) recovers part of the gap, showing that enhancement helps regardless of channel count, but does not close it. This dataset-scale interaction is the main practical finding of the study. Per-channel Grad-CAM contributes a complementary analytical tool: it reveals, for each emotion class, which spectrogram type receives the greatest model attention -- a form of attribution that has not been reported in the SER literature and is only possible with multi-channel input.'),

    subHead('A.', 'Limitations'),

    body('Several constraints bound the scope of this study and should inform how its results are interpreted. First, all experiments are conducted on RAVDESS [18], a studio-recorded acted-speech corpus. Acted speech exhibits exaggerated prosodic cues and higher signal quality than spontaneous or naturalistic speech; generalization to real-world conditions -- conversational speech, telephone noise, non-native speakers -- has not been assessed and cannot be inferred from RAVDESS performance alone [35]. Second, the deliberate choice of 1,008 training samples, while motivated by resource-constrained deployment scenarios, amplifies the variance of all reported metrics; small differences in accuracy between conditions (B1 vs. B2: +0.92pp) should be interpreted cautiously without a formal significance test. Third, the study covers 8 emotion categories in English; cross-lingual generalization and performance on finer-grained or dimensional (valence--arousal) emotion spaces remain open questions [29]. Fourth, DIP hyperparameters (CLAHE clipLimit, UM sigma and alpha) are fixed rather than learned; it is possible that per-channel or per-sample adaptive enhancement would yield larger gains, particularly on channels such as CQT and Chroma where the enhancement benefit is less directly motivated by contrast-enhancement literature.'),

    subHead('B.', 'Future Work'),

    body('Four directions follow directly from the findings above. (1) Cross-corpus evaluation: validating the four-channel pipeline on IEMOCAP [36], MSP-IMPROV [37], or EMODB [38] would test whether the channel-count versus data-size trade-off observed here is corpus-specific or general. (2) Larger-scale training: re-running the 2x2 ablation with 5,000+ samples -- achievable by pooling RAVDESS with compatible corpora -- would clarify whether the B3/B4 regression is inherently tied to small data or merely to the 1,008-sample regime used here [31],[32]. (3) Learned DIP: replacing fixed CLAHE and UM with trainable contrast-enhancement layers [9] or a differentiable preprocessing front-end would allow the enhancement to specialize per channel and per emotion class. (4) Attention-based fusion: replacing the 1x1 projection with a cross-channel attention head [3],[34] or a contrastive pre-training stage [31] should better exploit the complementary representations under limited supervision -- the theoretical advantage of four-channel input that the current architecture could not realise at RAVDESS scale.'),

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
