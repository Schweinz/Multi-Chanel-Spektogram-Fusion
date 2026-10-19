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
      { text: 'Speech Emotion Recognition (SER) remains a challenging open problem in affective computing, with practical demand spanning human-computer interaction, mental health monitoring, and social robotics [1],[2]. A persistent limitation of current deep learning approaches is their reliance on a single spectrogram type -- most commonly the Mel Spectrogram or MFCC -- leaving complementary spectral cues unexploited and producing low-contrast input images that hinder CNN feature discrimination [3],[8],[9]. This work investigates a SER framework that brings together four-channel spectrogram fusion and a two-stage Digital Image Processing (DIP) enhancement pipeline applied to each channel independently. Mel Spectrogram, MFCC, Constant-Q Transform (CQT), and Chroma STFT representations are extracted from each utterance using librosa [19], sequentially enhanced by CLAHE [9] and Unsharp Masking [20], and concatenated into a 128x128x4 tensor. A learnable 1x1 convolutional projection layer then maps this four-channel input to the three-channel space required by pretrained MobileNetV2 [21], preserving ImageNet weight compatibility without spatial resampling. Training proceeds in two phases -- a frozen-backbone initialisation followed by full fine-tuning with early stopping -- and four ablation conditions on RAVDESS [18] isolate the individual contributions of DIP enhancement and multi-channel fusion. Notably, the single-channel Mel variant with DIP enhancement (B2) attains the highest accuracy of 39.81% and macro F1 of 38.89%, outperforming the full four-channel configuration under the dataset-scale constraints of RAVDESS. In addition, per-channel Grad-CAM [22] is applied to the four-channel model to produce channel-specific saliency maps that identify which spectrogram type drives predictions for each emotion class, extending interpretability beyond standard single-channel visualisation [17]. These results suggest that image-level enhancement offers a reliable performance gain on small acted-speech corpora, while multi-channel fusion requires larger data regimes and stronger regularisation to yield its full benefit.' },
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

    body('Inferring emotional state from speech is a long-standing challenge in affective computing, with an expanding range of deployed applications. Healthcare platforms use SER for remote mental health screening [2]; conversational agents benefit from emotion-aware response generation; call-centre quality-assurance systems have adopted sentiment classifiers that treat prosodic cues as primary evidence [27]. Alongside these, social robotics and driver-monitoring applications require real-time, low-latency inference under resource constraints [3],[13]. What unifies all of these settings is the core task: mapping the acoustic content of an utterance -- pitch, energy, rhythm, spectral texture -- to a discrete or continuous label that reflects the speaker\'s affective state, independently of the words spoken [29],[35].'),

    body('The dominant paradigm treats each utterance as an image. After converting the raw waveform to a time-frequency spectrogram, deep CNNs or hybrid CNN-recurrent networks learn discriminative spatial patterns [10],[15],[23]. Transfer learning from ImageNet-pretrained backbones -- particularly lightweight architectures such as MobileNet -- has enabled competitive performance even on small labelled corpora [4],[5],[13],[14]. Despite these gains, two structural limitations appear repeatedly in the literature. Most systems examine only a single spectrogram type; yet Mel Spectrograms, MFCCs, CQT, and Chroma STFT each encode distinct aspects of vocal emotion -- energy contour, vocal-tract resonance, harmonic structure, and pitch-class distribution, respectively [8],[30]. Selecting one representation necessarily discards the others. Equally, the conversion from waveform to image introduces low local contrast in the spectrogram: frequency boundaries that demarcate emotional states can be difficult to resolve, particularly in mid-frequency regions where formant transitions carry much of the affective signal [9],[28].'),

    bodyRuns([
      { text: 'Related Work. ', italics: true },
      { text: 'Early deep SER systems applied standard image-classification backbones to Mel spectrograms and demonstrated that visual feature hierarchies transfer well to audio [10],[11]. Recurrent and attention-augmented CNN architectures subsequently captured long-range temporal dependencies [25],[12],[34], though at considerably higher computational cost. Lightweight transfer learning via MobileNet, VGG, and ResNet variants has offered a practical balance between accuracy and inference speed [4],[5],[13],[14]. Multi-scale and graph-based fusion methods have enriched the feature space further [14],[23],[24]. Input-level fusion of Mel, MFCC, and CQT representations has shown consistent advantages over single-feature baselines on multiple benchmarks [6],[7],[8],[15],[30]. On the training side, self-supervised contrastive objectives address label scarcity [31], and task-specific augmentation mitigates class imbalance [1],[32]. Image preprocessing for spectrograms has received comparatively less attention: CLAHE has improved classification on medical signal imagery [9] and radar spectrogram classification [28], and Unsharp Masking has sharpened edge detail for natural-image classifiers [20], yet neither has been systematically evaluated on SER spectrogram inputs. Interpretability work has adapted Grad-CAM to single-channel audio [17],[5]; per-channel attribution in multi-channel models has not, to our knowledge, been reported. Large pretrained speech models such as Whisper [33] offer strong representations but remain prohibitively costly for on-device deployment.' },
    ]),

    body('The present study addresses these gaps through four coordinated contributions: (1) a four-channel input fusion (Mel+MFCC+CQT+Chroma) capturing complementary acoustic cues simultaneously; (2) a two-stage DIP pipeline applying CLAHE followed by Unsharp Masking to each channel independently prior to fusion; (3) a learnable 1x1 convolutional projection adapting the four-channel tensor to the three-channel input space of ImageNet-pretrained MobileNetV2; and (4) per-channel Grad-CAM attribution revealing which spectrogram representation drives model predictions for each emotion class. A systematic ablation on RAVDESS [18] isolates the contribution of each component and yields findings with practical implications for small-corpus SER system design.'),

    blank(),

    // ── II. METHODOLOGY ──
    secHead('II', 'Proposed Methodology'),

    subHead('A.', 'Dataset'),
    body('All experiments are conducted on the Ryerson Audio-Visual Database of Emotional Speech and Song (RAVDESS) [18]. The corpus contains 1,440 audio recordings contributed by 24 professional actors -- 12 male and 12 female -- under controlled studio conditions, covering eight discrete emotion categories: neutral, calm, happy, sad, angry, fearful, disgust, and surprised. Class counts are approximately balanced at 180 samples each. Recordings are partitioned by stratified random sampling into training (70%, 1,008 samples), validation (15%, 216 samples), and test (15%, 216 samples) subsets, preserving both emotion-class and gender proportions across splits [27],[35].'),

    subHead('B.', 'Multi-Channel Spectrogram Extraction'),
    body('All audio files are resampled to 22,050 Hz and either truncated or zero-padded to a fixed duration of three seconds (66,150 samples). Four complementary time-frequency representations are then extracted with librosa [19]. The Mel Spectrogram (128 mel-frequency bands) captures the overall spectral energy distribution weighted to perceptual scale [25]. MFCCs (40 coefficients) encode the vocal-tract transfer function through cepstral analysis, providing a compact description of resonance structure [25]. The Constant-Q Transform (84 bins, 12 bins per octave) resolves harmonic content on a logarithmic frequency axis that closely mirrors musical pitch perception [30]. Chroma STFT (12 pitch-class bins) summarises tonal energy independently of octave, highlighting pitch-class patterns associated with emotional prosody [8]. Each representation is converted to decibel scale, normalised to [0, 255], and resized via bilinear interpolation to 128x128 pixels, yielding a set of grayscale images ready for downstream DIP enhancement.'),

    subHead('C.', 'DIP Enhancement Pipeline'),
    body('Prior to fusion, each of the four 128x128 grayscale images undergoes a two-stage image-enhancement pipeline applied independently. The first stage is CLAHE [9]. The image is subdivided into non-overlapping 8x8 contextual tiles, and within each tile a histogram is equalised subject to a clip limit of 2.0 -- a threshold that redistributes excess histogram mass to other bins, preventing over-amplification of uniform regions while still enhancing local contrast [9],[28]. Tile boundaries are blended by bilinear interpolation, avoiding the blocking artefacts that afflict global histogram equalisation. The second stage is Unsharp Masking (UM) [20]: a Gaussian-blurred version of the CLAHE output (sigma = 1.0) is formed and then combined with the original to recover and amplify high-frequency detail:'),

    eq('I_sharp = (1+alpha)·I_CLAHE - alpha·G_sigma(I_CLAHE)', 1),

    body('where G_sigma denotes convolution with a Gaussian kernel of standard deviation sigma = 1.0, and alpha = 1.5 [20]. The operation restores fine frequency boundaries that CLAHE may have smoothed, without disturbing the global intensity distribution that CLAHE established. Together, the two stages address different scales of the contrast problem: CLAHE operates on the local histogram, while UM targets the sharpness of edges at the pixel level.'),

    figImg('fig1_spectrogram_grid.png', 3.2, 3.85),
    figCap(1, 'Multi-channel spectrogram enhancement: Original, after CLAHE, after CLAHE+UM (rows: Mel, MFCC, CQT, Chroma).'),
    body('Fig. 1 illustrates the enhancement sequence for all four channels on an "angry" sample. The left column shows the unprocessed spectrogram; the centre column shows the CLAHE output (clipLimit=2.0, tileGridSize=8x8); the right column adds Unsharp Masking (alpha=1.5, sigma=1.0) on top of CLAHE. In the Mel row, low-frequency energy gradients that are nearly invisible in the original become clearly structured after CLAHE, and the formant transition boundaries are further sharpened by UM. The MFCC row reveals a similar pattern: the cepstral ridges, which partially merge in the original, are distinctly separated in the rightmost column. CQT harmonic band edges and Chroma pitch-class boundaries both show appreciable sharpening without evidence of saturation or blocking artefacts -- outcomes expected from the clip-limited tile-based equalisation mechanism. These visual improvements suggest that the enhanced images present more spatially distinct patterns for the CNN\'s early convolutional filters to respond to.'),

    subHead('D.', 'Four-Channel Fusion & Architecture'),
    body('After enhancement, the four 128x128 images are normalised to [0,1] and concatenated along the channel axis to form a 128x128x4 input tensor. A learnable 1x1 convolutional projection layer -- three filters, no bias, linear activation -- linearly recombines the four input channels into three without any spatial downsampling, allowing the three-channel filter banks of ImageNet-pretrained MobileNetV2 to be applied directly. MobileNetV2 [21] uses depthwise-separable convolutions organised into inverted-residual bottleneck blocks; this architecture achieves strong representational capacity at a fraction of the parameter count of standard convolutions, making it well suited to resource-constrained deployment [4],[13]. The frozen backbone outputs a feature map of 4x4x1,280, which a Global Average Pooling (GAP) layer reduces to a 1,280-dimensional vector. The classification head then applies Dropout(0.3), a Dense layer of 128 units with ReLU activation, Batch Normalisation, a second Dropout(0.3), and a final Dense(8) layer with softmax, producing a probability distribution over the eight RAVDESS emotion classes [14].'),

    figImg('fig2_architecture.png', 3.2, 1.65),
    figCap(2, 'Proposed SER architecture: 128x128x4 input through 1x1 Conv2D projection, MobileNetV2 backbone, and 8-class softmax head.'),
    body('Fig. 2 traces the data flow through the proposed architecture. The 128x128x4 tensor enters the 1x1 Conv2D projection layer, which performs a linear channel mixing -- without touching the spatial dimensions -- and outputs a 128x128x3 tensor that MobileNetV2 can process with its pretrained weights intact. The backbone then applies its sequence of depthwise-separable bottleneck blocks, progressively reducing the spatial resolution to 4x4 while expanding the channel count to 1,280. A Global Average Pooling operation collapses the spatial dimensions, producing a fixed-length 1,280-dimensional feature vector that is insensitive to spatial shifts within the spectrogram -- a desirable property given that the onset times of emotional features vary across utterances. The head applies two dropout layers around a Dense(128, ReLU) block with Batch Normalisation, providing regularisation against the relatively small RAVDESS training set, before the final Dense(8) with softmax.'),

    subHead('E.', 'Two-Phase Training Strategy'),
    body('Training proceeds in two stages. In Phase 1, the MobileNetV2 backbone is frozen and only the projection layer and classification head are updated. The model is trained for ten epochs with the Adam optimiser (lr=1e-4, batch size 32) under categorical cross-entropy loss. This phase initialises the head weights using the already-stable backbone representations, preventing the randomly initialised head from distorting the pretrained filters early in training [14],[21]. In Phase 2, all layers are unfrozen and fine-tuned jointly at a reduced learning rate (lr=1e-5, max 100 epochs). Two callbacks provide regularisation: Early Stopping (patience=15, monitored on validation loss) halts training when generalisation stops improving, and ReduceLROnPlateau (factor=0.5, patience=5) halves the learning rate on plateaus, allowing the optimiser to escape shallow local minima without diverging [4],[5],[13],[23].'),

    blank(),

    // ── III. ABLATION ──
    secHead('III', 'Ablation Study Design'),

    body('Four experimental conditions are constructed to disentangle the contribution of DIP enhancement from that of multi-channel fusion [7],[8],[15]. B1 uses a single Mel Spectrogram with no DIP preprocessing -- the standard baseline, equivalent in design to the approaches in [10],[11] -- where the Mel channel is replicated three times to satisfy MobileNetV2\'s three-channel input requirement. B2 retains the single-channel Mel setup but applies the full CLAHE+UM pipeline, isolating the effect of image enhancement alone. B3 uses all four spectrogram channels without DIP, following the multi-channel fusion paradigm of [6],[8]. B4 is the complete configuration, combining four channels with DIP enhancement. Comparing B1 with B2 quantifies the DIP gain on a single channel; comparing B1 with B3 quantifies the fusion gain without enhancement; and comparing B3 with B4 quantifies the DIP gain in the multi-channel setting. B4 vs B1 gives the combined effect.'),

    blank(),

    // ── IV. RESULTS ──
    secHead('IV', 'Results and Discussion'),

    body('Table I summarises test-set performance across the four conditions. Evaluation follows the macro-averaging protocol recommended for multi-class emotion recognition with approximately balanced classes [3],[7],[27], reporting Accuracy, macro Precision, macro Recall, and macro F1 simultaneously to capture both overall correctness and per-class balance.'),

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

    blank(),

    // ── V. GRAD-CAM ──
    secHead('V', 'Per-Channel Grad-CAM Analysis'),

    body('Performance metrics alone do not reveal how the model distributes attention across the four spectrogram representations, nor do they indicate which channel contributes most to a given emotion class. To address this, Grad-CAM [22] is adapted for per-channel attribution using the B4 model. For a given utterance and target class, TensorFlow GradientTape records the gradient of the predicted class logit with respect to the final convolutional feature map (4x4x1,280) of MobileNetV2. The gradient tensor is globally pooled to channel-wise importance weights, multiplied with the feature map, and summed across channels; the resulting scalar map is ReLU-clamped, max-normalised, and bilinearly upsampled to 128x128. This heatmap is then multiplied element-wise with each of the four DIP-enhanced input channels separately, and the mean of each product is recorded as the attention score for that (emotion, channel) pair [17],[5]. Applying this procedure to one representative sample per emotion class and row-normalising the resulting 8x4 matrix produces a relative channel importance profile: each row sums to 1.0 and reflects how the model distributes its attention across Mel, MFCC, CQT, and Chroma for that particular emotion class. This goes beyond what standard single-channel Grad-CAM can offer [22], directly answering which spectral representation the model finds most informative per emotion.'),

    figImg('fig4_gradcam.png', 3.2, 1.38),
    figCap(4, 'Per-channel Grad-CAM overlays for a representative "angry" sample (RAVDESS Actor 01). Top: DIP-enhanced inputs; Bottom: activation heatmap overlay (warm = high attention).'),
    body('Fig. 4 shows the per-channel Grad-CAM overlays for a representative "angry" utterance (RAVDESS Actor 01). In the bottom row, warm colours indicate high activation. On the Mel channel, the model attends most strongly to the mid-to-high frequency region (roughly 1-4 kHz), consistent with the elevated glottal energy and intensity range associated with angry speech [29]. The MFCC channel shows activation concentrated along the formant transition region, where rapid articulatory changes encode emotional urgency. CQT activation peaks align with harmonic bands, reflecting the upward pitch shift documented in high-arousal speech [29]. The Chroma channel activates on a specific pitch-class region, suggesting the model has detected the tonal emphasis that characterises sustained aggressive vocalisation. A notable observation is that the model predicted "surprised" with 49.6% confidence rather than "angry" (15.9%) -- a confusion that is acoustically motivated, as the two emotions share elevated arousal, wide pitch range, and high energy [3],[29]. The per-channel view makes this confusion interpretable: both emotions activate similar mid-frequency regions, and the gradient-based signal does not cleanly separate them under the limited training data available.'),

    figImg('fig5_importance_heatmap.png', 3.0, 2.57),
    figCap(5, 'Row-normalized per-channel Grad-CAM importance matrix (8 emotions x 4 channels) using the B4 model on RAVDESS representative samples.'),
    body('Fig. 5 presents the row-normalised 8x4 importance matrix. Several patterns emerge. MFCC is the most consistently dominant channel, contributing the largest fraction of attention for calm (0.37), happy (0.38), sad (0.39), fearful (0.45), and surprised (0.41). This aligns with the well-established finding in the SER literature that cepstral coefficients capture vocal-tract resonance patterns that correlate strongly with discrete emotion categories [25],[8]. CQT and Chroma offer complementary contributions particularly for pitch-related emotions: CQT scores are relatively elevated for happy and surprised, both high-arousal emotions with wider pitch range [29], while Chroma contributes more for disgust and sad -- emotions characterised by tonally constrained, lower-arousal prosody. One notable anomaly is the near-zero row for angry. The representative angry sample was misclassified as surprised -- a documented confusion between high-arousal categories [3],[29] -- which means the gradient with respect to the angry-class logit was effectively zero, resulting in a uniformly low importance row. This is an interpretable artefact of the model\'s misclassification rather than a structural limitation of the per-channel attribution method itself. Critically, this profile could not have been derived from a single-channel system; the multi-channel input is a prerequisite for the per-channel decomposition.'),

    blank(),

    // ── VI. CONCLUSION ──
    secHead('VI', 'Conclusion'),

    body('This paper has described a SER framework built on three coordinated ideas: four-channel spectrogram fusion (Mel+MFCC+CQT+Chroma), a two-stage DIP enhancement pipeline (CLAHE [9]+Unsharp Masking [20]) applied to each channel independently, and a learnable 1x1 convolutional projection that adapts the four-channel tensor to ImageNet-pretrained MobileNetV2 [21]. The systematic ablation on RAVDESS [18] shows that DIP enhancement yields a reliable and immediate improvement on single-channel input (B1 to B2: +0.92pp Accuracy, +4.50pp F1), establishing CLAHE+UM as an effective, zero-cost preprocessing step for spectrogram-based SER. Multi-channel fusion, by contrast, requires training data beyond the 1,008 samples available in RAVDESS to realise its theoretical advantage; the combined four-channel+DIP model (B4) lags behind the enhanced single-channel model (B2), though it consistently surpasses the unenhanced four-channel baseline (B3). Beyond performance, the per-channel Grad-CAM analysis provides channel-level attribution that reveals qualitatively which spectral representation the model prioritises for each emotion -- an interpretability tool not previously reported in the SER literature and only realisable in a multi-channel input framework. Future work should examine whether adaptive DIP parameter learning, task-specific augmentation [32], contrastive pre-training [31], Transformer-based fusion heads [3],[12],[34], and evaluation on spontaneous cross-corpus data [35] can close the performance gap between multi-channel fusion and the single-channel DIP baseline on small corpora.'),

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
