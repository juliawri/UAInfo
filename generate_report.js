const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak, TableOfContents
} = require('/sessions/sweet-friendly-einstein/node_modules/docx');
const fs = require('fs');

// ── Colour palette ────────────────────────────────────────────────────────────
const BLUE      = "1F5C99";
const LIGHT_BLU = "D5E3F0";
const MID_BLU   = "2E75B6";
const DARK      = "1A1A2E";
const GREY_BG   = "F2F4F8";

// ── Border helpers ────────────────────────────────────────────────────────────
const cellBorder = { style: BorderStyle.SINGLE, size: 1, color: "BBCDE0" };
const cellBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── Helper: plain paragraph ───────────────────────────────────────────────────
function p(text, opts = {}) {
  const runs = Array.isArray(text)
    ? text
    : [new TextRun({ text, font: "Arial", size: 22, color: DARK, ...opts.runOpts })];
  return new Paragraph({
    children: runs,
    spacing: { before: 80, after: 120 },
    ...opts.paraOpts,
  });
}

// ── Helper: heading ───────────────────────────────────────────────────────────
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", size: 36, bold: true, color: BLUE })],
    spacing: { before: 480, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: MID_BLU, space: 4 } },
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color: MID_BLU })],
    spacing: { before: 320, after: 120 },
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: "44607A" })],
    spacing: { before: 200, after: 80 },
  });
}

// ── Helper: bullet ────────────────────────────────────────────────────────────
function bullet(text, level = 0, bold = false) {
  const runs = Array.isArray(text)
    ? text
    : [new TextRun({ text, font: "Arial", size: 22, color: DARK, bold })];
  return new Paragraph({
    numbering: { reference: "bullets", level },
    children: runs,
    spacing: { before: 40, after: 40 },
  });
}

// ── Helper: info box (shaded paragraph) ──────────────────────────────────────
function infoBox(paragraphs) {
  const contentWidth = 9360;
  return new Table({
    width: { size: contentWidth, type: WidthType.DXA },
    columnWidths: [contentWidth],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 4, color: MID_BLU },
                       bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLU },
                       left: { style: BorderStyle.THICK, size: 12, color: MID_BLU },
                       right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
            width: { size: contentWidth, type: WidthType.DXA },
            shading: { fill: GREY_BG, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 200, right: 120 },
            children: paragraphs,
          }),
        ],
      }),
    ],
  });
}

// ── Helper: 2-col table ───────────────────────────────────────────────────────
function twoColTable(rows, colWidths = [3000, 6360]) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: rows.map(([left, right], i) =>
      new TableRow({
        children: [
          new TableCell({
            borders: cellBorders,
            width: { size: colWidths[0], type: WidthType.DXA },
            shading: { fill: i === 0 ? LIGHT_BLU : "FFFFFF", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: left, font: "Arial", size: 22, bold: i === 0, color: DARK })] })],
          }),
          new TableCell({
            borders: cellBorders,
            width: { size: colWidths[1], type: WidthType.DXA },
            shading: { fill: i === 0 ? LIGHT_BLU : "FFFFFF", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: right, font: "Arial", size: 22, bold: i === 0, color: DARK })] })],
          }),
        ],
      })
    ),
  });
}

// ── Helper: metrics table ─────────────────────────────────────────────────────
function metricsTable(headers, dataRows) {
  const n = headers.length;
  const totalW = 9360;
  const colW = Math.floor(totalW / n);
  const colWidths = Array(n).fill(colW);
  colWidths[n - 1] += totalW - colW * n;

  const allRows = [
    new TableRow({
      children: headers.map((h, i) =>
        new TableCell({
          borders: cellBorders,
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { fill: MID_BLU, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: h, font: "Arial", size: 22, bold: true, color: "FFFFFF" })] })],
        })
      ),
    }),
    ...dataRows.map((row, ri) =>
      new TableRow({
        children: row.map((cell, ci) =>
          new TableCell({
            borders: cellBorders,
            width: { size: colWidths[ci], type: WidthType.DXA },
            shading: { fill: ri % 2 === 0 ? "FFFFFF" : GREY_BG, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 22, color: DARK })] })],
          })
        ),
      })
    ),
  ];

  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: allRows,
  });
}

function spacer(n = 1) {
  return Array(n).fill(null).map(() => new Paragraph({ children: [new TextRun("")], spacing: { before: 0, after: 0 } }));
}

// ═════════════════════════════════════════════════════════════════════════════
// DOCUMENT CONTENT
// ═════════════════════════════════════════════════════════════════════════════

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
        ],
      },
    ],
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: DARK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: BLUE },
        paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: MID_BLU },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "44607A" },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ],
  },
  sections: [
    // ── TITLE PAGE ────────────────────────────────────────────────────────────
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children: [
        ...spacer(6),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "URBAN ACCESS", font: "Arial", size: 72, bold: true, color: BLUE })],
          spacing: { before: 0, after: 120 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "AI-Powered Accessible Routing for Montreal", font: "Arial", size: 36, color: MID_BLU, italics: true })],
          spacing: { before: 0, after: 480 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: MID_BLU, space: 4 } },
          children: [new TextRun("")],
          spacing: { before: 0, after: 480 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "Comprehensive Project Report", font: "Arial", size: 28, color: "44607A" })],
          spacing: { before: 0, after: 120 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "AI4Good Lab — Demo Day 2026", font: "Arial", size: 26, color: "44607A" })],
          spacing: { before: 0, after: 120 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "June 2026", font: "Arial", size: 24, color: "888888" })],
          spacing: { before: 0, after: 0 },
        }),
        ...spacer(4),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            text: "Covers: Dataset Creation • ML Pipeline • Application Architecture • Evaluation Criteria • Panel Q&A",
            font: "Arial", size: 20, color: "888888", italics: true
          })],
        }),
        new Paragraph({ children: [new PageBreak()] }),
      ],
    },
    // ── MAIN BODY ────────────────────────────────────────────────────────────
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: "Urban Access — Project Report", font: "Arial", size: 18, color: "888888", italics: true }),
                new TextRun({ text: "\t", font: "Arial", size: 18 }),
                new TextRun({ text: "AI4Good Lab 2026", font: "Arial", size: 18, color: "888888" }),
              ],
              tabStops: [{ type: "right", position: 9360 }],
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: "Confidential — AI4Good Lab Demo Day", font: "Arial", size: 18, color: "AAAAAA" }),
                new TextRun({ text: "\t", font: "Arial", size: 18 }),
                new TextRun({ text: "Page ", font: "Arial", size: 18, color: "888888" }),
                new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "888888" }),
              ],
              tabStops: [{ type: "right", position: 9360 }],
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 4 } },
            }),
          ],
        }),
      },
      children: [

        // ── TABLE OF CONTENTS ─────────────────────────────────────────────────
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun({ text: "Table of Contents", font: "Arial", size: 36, bold: true, color: BLUE })],
          spacing: { before: 0, after: 200 },
        }),
        new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
        new Paragraph({ children: [new PageBreak()] }),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 1: PROJECT OVERVIEW
        // ══════════════════════════════════════════════════════════════════════
        h1("1. Project Overview"),

        p("Urban Access is an AI-powered accessible routing application designed to help individuals with mobility impairments navigate the streets of Montreal safely and effectively. The project was built as part of the AI4Good Lab program and presented at Demo Day 2026."),

        p("Montreal is home to over 200,000 residents with mobility impairments. While the city has made strides in accessible infrastructure, existing navigation tools — such as Google Maps — do not account for the real, ground-level conditions that determine whether a sidewalk, ramp, or path is actually usable. A government building may be accessible, but the route to it often is not. Urban Access addresses this gap by integrating computer vision, street-level imagery, and large language models to generate route recommendations optimized for a user's specific mobility aid."),

        h2("1.1 Elevator Pitch"),
        infoBox([
          new Paragraph({
            children: [new TextRun({ text: "\"What if your navigation app actually understood what it’s like to use a wheelchair?\"", font: "Arial", size: 22, italics: true, color: MID_BLU })],
            spacing: { before: 60, after: 60 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "Urban Access gives people with mobility impairments a routing tool that reflects real sidewalk conditions — not just distance and traffic. Using AI-scored street-level images, it recommends the most accessible route for your specific mobility aid, with plain-language explanations of why.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 60, after: 0 },
          }),
        ]),
        ...spacer(1),

        h2("1.2 Problem Statement"),
        p("People with mobility impairments face daily challenges navigating urban environments that standard mapping tools are not designed to address. Key issues include:"),
        bullet("Sidewalk surface quality (cracked, uneven, cobblestone)"),
        bullet("Ramp availability and grade at curb cuts and entrances"),
        bullet("Construction zones temporarily blocking accessible paths"),
        bullet("Snow and debris accumulation (seasonal in Montreal)"),
        bullet("Steepness and elevation changes along routes"),
        bullet("Pedestrian safety concerns at intersections"),
        ...spacer(1),
        p("These factors affect different mobility aid users differently. A route passable with a walking cane may be impassable for an electric wheelchair. Urban Access models these distinctions explicitly, supporting six mobility aid types: manual wheelchair, electric wheelchair, walker, walking cane, mobility scooter, and mobility impairment without aid."),

        h2("1.3 Core Technical Approach"),
        p("Urban Access combines three AI layers:"),
        bullet([new TextRun({ text: "Fine-tuned Vision Transformer (ViT): ", font: "Arial", size: 22, bold: true, color: DARK }), new TextRun({ text: "Classifies sidewalk accessibility from street-level images for each of the six mobility aid types.", font: "Arial", size: 22, color: DARK })]),
        bullet([new TextRun({ text: "Vision-Language Model (VLM) secondary scoring: ", font: "Arial", size: 22, bold: true, color: DARK }), new TextRun({ text: "Gemini 2.5 Flash provides detailed, contextual ratings at the most uncertain or critical route points.", font: "Arial", size: 22, color: DARK })]),
        bullet([new TextRun({ text: "Open-source Large Language Model (LLM): ", font: "Arial", size: 22, bold: true, color: DARK }), new TextRun({ text: "Synthesises ViT scores, VLM annotations, pedestrian safety, construction data, inclination, and user feedback into a plain-language route recommendation with two alternatives.", font: "Arial", size: 22, color: DARK })]),
        ...spacer(1),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 2: COMMUNITY OUTREACH
        // ══════════════════════════════════════════════════════════════════════
        h1("2. Community Outreach and Stakeholder Engagement"),

        p("Meaningful community consultation was central to Urban Access from the outset. The team engaged directly with disability advocates, urban planning researchers, and individuals with lived experience to ensure the system addressed genuine needs rather than assumed ones."),

        h2("2.1 Key Consultations"),

        h3("Rashid Mushkani — Urban Planning and AI Researcher"),
        p("Rashid Mushkani provided the conceptual grounding for the project’s routing focus. Two key insights from his consultation shaped our direction:"),
        bullet([new TextRun({ text: "Universal accessibility and the last-mile gap: ", font: "Arial", size: 22, bold: true, color: DARK }), new TextRun({ text: "Rashid introduced the concept of universal accessibility — that accessibility must be continuous, not just at destinations. A government building may be fully accessible, but if the sidewalk leading to it is not, it is effectively inaccessible. This framing shifted our focus from point-level scoring to route-level recommendation.", font: "Arial", size: 22, color: DARK })]),
        bullet([new TextRun({ text: "Connection to Promo-Accès and survey outreach: ", font: "Arial", size: 22, bold: true, color: DARK }), new TextRun({ text: "Based on his guidance, the team initiated an additional consultation with Promo-Accès (a Montreal-based accessibility organisation) and expanded our survey outreach strategy.", font: "Arial", size: 22, color: DARK })]),
        ...spacer(1),

        h3("Laurence Parent — Plateau-Mont-Royal City Councillor"),
        p("Laurence Parent provided ongoing feedback throughout the project’s development. As a public official with direct engagement with accessibility issues in one of Montreal’s most walkable but challenging neighbourhoods, she helped the team verify that the system’s outputs accurately reflected the needs and experiences of disabled individuals. Her involvement constituted a form of continuous, iterative validation from a policy and community perspective."),

        h2("2.2 Survey Campaign"),
        p("To build ground-truth accessibility labels, the team conducted a structured survey campaign. Street-level images from across Montreal were presented to community members and disability advocates for manual rating. Key details:"),

        ...spacer(1),
        metricsTable(
          ["Dimension", "Details"],
          [
            ["Distribution channels", "30+ disability-focused community groups across Montreal"],
            ["Survey responses collected", "154 rated image-aid pairs across 76 unique images"],
            ["Aid types covered", "Manual wheelchair, electric wheelchair, walker, walking cane, mobility scooter, mobility impairment no aid"],
            ["Rating scale", "0–4 (0 = completely inaccessible, 4 = fully accessible)"],
            ["Data handling", "Fully anonymised; no personally identifiable information collected"],
          ]
        ),
        ...spacer(1),

        p("Survey responses were not used directly as training labels (the sample was too small for that purpose), but served two critical functions: (1) calibrating the statistical distribution of AI-generated labels to match human perception, and (2) providing partial external validation of the VLM scoring approach."),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 3: DATASET CREATION
        // ══════════════════════════════════════════════════════════════════════
        h1("3. Dataset Creation Pipeline"),

        p("One of the most substantial contributions of this project is a purpose-built, large-scale accessibility-labelled street-level image dataset for Montreal. No existing public dataset combined street-level imagery with per-mobility-aid accessibility labels at this scale."),

        h2("3.1 Image Collection via Mapillary"),
        p("Mapillary is an open-source, crowd-sourced street-level imagery platform similar to Google Street View. It provides a public API for retrieving images by geographic coordinate, making it suitable for structured dataset collection without the licensing restrictions or cost of Google Street View."),
        p("The collection strategy focused on areas with high relevance to mobility-impaired users:"),
        bullet("Coordinates sampled near Montreal transit stations (bus stops, metro stations)"),
        bullet("Coverage extended across Montreal Island"),
        bullet("Approximately 14,000 unique coordinate locations collected"),
        bullet("Multiple view angles captured per location where available"),
        bullet("Images stored with filename format: {latitude}_{longitude}_{view_angle}.jpg"),
        ...spacer(1),
        p("Three separate data collection repositories contributed images to the final dataset: the primary gemini_dataset repository, extra_gemini for extended coverage, and ps_data_extraction for additional locations. Images from all three sources were also used in the Masked Autoencoder pre-training phase (see Section 4)."),

        h2("3.2 AI Labelling with Gemini 2.5 Flash (VLM Scoring)"),
        p("Manual labelling of 14,000+ locations across six mobility aid types would require approximately 500,000 individual human ratings — infeasible at project scale. Instead, Gemini 2.5 Flash (Google’s vision-language model) was used to generate accessibility scores programmatically."),
        p("For each image, Gemini was prompted to rate accessibility for each of the six mobility aid types on a 0–4 integer scale, along with a text justification. The structured prompt format ensured consistent output. Total labelled records: 31,887 rows in the final dataset."),
        ...spacer(1),
        metricsTable(
          ["Mobility Aid Type", "Gemini Mean Score", "Gemini Std Dev"],
          [
            ["Manual wheelchair", "2.235", "1.236"],
            ["Electric wheelchair", "(calibrated to survey)", "—"],
            ["Walker", "2.611", "1.448"],
            ["Walking cane", "(calibrated to survey)", "—"],
            ["Mobility scooter", "(global fallback used)", "—"],
            ["Mobility impairment (no aid)", "(calibrated to survey)", "—"],
          ]
        ),
        ...spacer(1),

        h2("3.3 Label Calibration and Binarization"),
        p("A key methodological contribution is the recalibration procedure that bridges AI-generated scores and human perception. Raw Gemini scores have a different statistical distribution than human survey ratings — a bias arising from the model’s different interpretation of the scale. Left uncorrected, this would cause training labels to be systematically offset from human judgement."),

        h3("Calibration Formula"),
        infoBox([
          new Paragraph({
            children: [new TextRun({ text: "z = (gemini_score − μ_gemini) / (σ_gemini + ε)", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 60, after: 40 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "recalibrated = z × σ_survey + μ_survey", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 0, after: 40 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "binary_label = 1 if recalibrated ≥ 2.5 else 0", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 0, after: 60 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "This z-score normalisation transforms Gemini scores into the same mean and standard deviation as the human survey responses for each aid type, before applying the binarization threshold.", font: "Arial", size: 20, color: "555555", italics: true })],
            spacing: { before: 0, after: 0 },
          }),
        ]),
        ...spacer(1),

        h3("Survey Statistics Used for Calibration"),
        metricsTable(
          ["Mobility Aid Type", "Survey n", "Survey Mean", "Survey Std Dev"],
          [
            ["Manual wheelchair", "20", "2.200", "1.005"],
            ["Electric wheelchair", "54", "2.704", "1.039"],
            ["Walker", "38", "2.211", "1.069 (raw scores used; no recal.)"],
            ["Walking cane", "40", "2.400", "1.033"],
            ["Mobility scooter", "n/a", "2.727 (global fallback)", "1.122"],
            ["Mobility impairment (no aid)", "40", "3.050", "1.260"],
          ]
        ),
        ...spacer(1),
        p("Note: Walker labels used raw Gemini scores without recalibration due to dataset characteristics. Mobility scooter lacked sufficient survey responses for a per-aid-type calibration, so global statistics were used as a fallback."),

        h3("Class Balance After Binarization"),
        metricsTable(
          ["Mobility Aid Type", "Accessible (1)", "Inaccessible (0)"],
          [
            ["Manual wheelchair", "63.2%", "36.8%"],
            ["Electric wheelchair", "76.8%", "23.2%"],
            ["Walker", "30.3%", "69.7%"],
            ["Walking cane", "74.8%", "25.2%"],
            ["Mobility scooter", "76.8%", "23.2%"],
            ["Mobility impairment (no aid)", "74.8%", "25.2%"],
          ]
        ),
        ...spacer(1),
        p("The walker class is notably more imbalanced, reflecting that the raw Gemini scores for walkers skewed higher (mean 2.611) but were not recalibrated, resulting in more locations being labelled inaccessible after binarization at 2.5. This class imbalance was one of the factors explored during the ViT fine-tuning phase."),

        h2("3.4 Final Dataset"),
        p("The final labelled dataset (gemini_labels_binary.csv) contains 31,887 rows, each with:"),
        bullet("image_path: filename in {lat}_{lon}_{view}.jpg format"),
        bullet("Raw Gemini scores for all six aid types (0–4 integer)"),
        bullet("Recalibrated continuous scores per aid type"),
        bullet("Binary accessibility labels per aid type (0 or 1)"),
        ...spacer(1),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 4: MAE PRE-TRAINING
        // ══════════════════════════════════════════════════════════════════════
        h1("4. Masked Autoencoder (MAE) Pre-Training"),

        p("Before fine-tuning a Vision Transformer for accessibility classification, the model was pre-trained using Masked Autoencoder (MAE) methodology — a self-supervised learning technique that teaches the model to reconstruct masked patches of an image. This domain adaptation step is critical: a ViT pre-trained on ImageNet learns general visual features but has no specific knowledge of street-level images, sidewalk surfaces, kerbs, ramps, or urban geometry. MAE pre-training on our accessibility dataset gives the model a strong domain-specific initialisation before any labelled examples are shown."),

        h2("4.1 Training Configuration"),
        metricsTable(
          ["Parameter", "Value"],
          [
            ["Training data", "Street-level images from gemini_dataset, extra_gemini, and ps_data_extraction"],
            ["Target epochs", "200"],
            ["Completed epochs", "95 (training stopped early)"],
            ["Steps per epoch", "975"],
            ["Approximate time per epoch", "536–553 seconds (≈9 minutes)"],
            ["Checkpoint frequency", "Every 5 epochs to Google Drive"],
            ["Optimisation target", "Masked patch reconstruction loss (MSE)"],
          ]
        ),
        ...spacer(1),

        h2("4.2 Training Loss Curve"),
        p("The MAE reconstruction loss decreased substantially over the 95 completed epochs, indicating effective learning of domain-relevant visual representations:"),
        ...spacer(1),
        metricsTable(
          ["Epoch", "Reconstruction Loss"],
          [
            ["1", "1.8415"],
            ["10", "0.4671"],
            ["20", "0.4159"],
            ["30", "0.4043"],
            ["40", "0.4000"],
            ["50", "0.3956"],
            ["60", "0.3926"],
            ["70", "0.3896"],
            ["80", "0.3858"],
            ["90", "0.3824"],
            ["95 (final)", "0.3813"],
          ]
        ),
        ...spacer(1),
        p("The rapid initial drop from 1.84 to 0.47 in the first 10 epochs reflects fast acquisition of low-level visual structure. The continued slow decrease from epoch 10 through 95 indicates progressive refinement of higher-level representations. Training was halted at 95 epochs due to time constraints; the loss trajectory suggests additional epochs would have yielded further improvement, and this is noted as a limitation."),

        h2("4.3 MAE Architecture Role"),
        p("The MAE encoder (approximately 85M parameters of a ViT-Base architecture) was used as the backbone initialisation for all subsequent ViT fine-tuning passes. Rather than starting from ImageNet weights, fine-tuning began from these street-view-adapted weights, giving the classifier a head start on recognising the visual features that matter for accessibility assessment."),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 5: VIT FINE-TUNING
        // ══════════════════════════════════════════════════════════════════════
        h1("5. Vision Transformer (ViT) Fine-Tuning"),

        p("With MAE-pre-trained encoder weights in hand, the team conducted six rounds (passes) of supervised fine-tuning to build the accessibility classifier. Each pass explored different hyperparameter configurations and data strategies. Pass 1 was ultimately selected for deployment in the production application."),

        h2("5.1 Fine-Tuning Architecture"),
        p("The ViT fine-tuning used a two-phase approach for all passes:"),
        ...spacer(1),
        metricsTable(
          ["Phase", "Trainable Parameters", "Description"],
          [
            ["Phase 1", "~1.187M", "Classification heads only. Encoder weights frozen. Allows the new classification layers to align with the pre-trained representations before any disruptive full-model updates."],
            ["Phase 2", "~87.0M (full model)", "Entire model (encoder + heads) trained end-to-end. Allows the encoder to adapt its representations specifically for the accessibility classification task."],
          ]
        ),
        ...spacer(1),
        p("The classification head outputs a binary prediction (accessible / inaccessible) for each of the six mobility aid types independently. The model therefore performs six parallel binary classifications from a single image forward pass."),

        h2("5.2 Evaluation Metrics"),
        p("Model performance was evaluated using two primary metrics:"),
        bullet([new TextRun({ text: "Validation Accuracy: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Overall fraction of correct binary predictions across all classes and images in the validation set.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Mean F1 Score: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Macro-averaged F1 score across all six mobility aid classes. This metric is preferred over accuracy given class imbalances, as it balances precision and recall for both accessible and inaccessible predictions.", font: "Arial", size: 22 })]),
        ...spacer(1),

        h2("5.3 Pass-by-Pass Summary"),
        ...spacer(1),
        metricsTable(
          ["Pass", "Best Val Mean F1", "Best Val Accuracy", "Key Changes from Previous Pass"],
          [
            ["Pass 1 ★ (USED)", "0.726 (epoch 21, Ph.2)", "0.761", "Baseline: MAE encoder weights, standard dropout (0.2), lr=1e-5 Ph.2, 6 output heads"],
            ["Pass 2", "0.669", "—", "Added label smoothing (0.1), increased dropout (0.2→0.35), reduced lr (1e-5→5e-6) in Ph.2"],
            ["Pass 3", "0.641", "—", "Added weighted random sampling to address walker class imbalance"],
            ["Pass 4", "0.685", "~0.795", "Removed walker class head, removed weighted sampling, dropout back to 0.2"],
            ["Pass 5", "0.651", "—", "Downsampled majority classes in training data"],
            ["Pass 6", "0.608", "—", "Added more data, changed walker threshold, no downsampling; longer epochs (~126s) suggesting different hardware"],
          ]
        ),
        ...spacer(1),

        h2("5.4 ViT Pass 1 — Production Model (Detailed)"),
        p("ViT Pass 1 was selected for the final application. Below are the detailed training results:"),

        h3("Phase 1: Classification Heads Only"),
        metricsTable(
          ["Metric", "Value"],
          [
            ["Trainable parameters", "1,187,334 (~1.19M)"],
            ["Epochs run", "10"],
            ["Best epoch", "3"],
            ["Best validation mean F1", "0.673"],
          ]
        ),
        ...spacer(1),

        h3("Phase 2: Full Model Fine-Tuning"),
        metricsTable(
          ["Metric", "Value"],
          [
            ["Trainable parameters", "87,012,870 (~87.0M)"],
            ["Epochs run", "50"],
            ["Best epoch", "21"],
            ["Best validation mean F1", "0.726"],
            ["Best validation accuracy", "0.761"],
            ["Per-class F1 range (epoch 21)", "~0.70–0.75 across all six heads"],
          ]
        ),
        ...spacer(1),

        h3("Per-Class F1 Scores (Pass 1, Phase 2, Best Epoch)"),
        metricsTable(
          ["Mobility Aid Type", "Approx. F1 Score"],
          [
            ["Manual wheelchair (manu)", "~0.72"],
            ["Electric wheelchair (elec)", "~0.73"],
            ["Walker (walk)", "~0.70"],
            ["Walking cane (walk2)", "~0.73"],
            ["Mobility scooter (mobi)", "~0.74"],
            ["Mobility impairment no aid (mobi2)", "~0.75"],
          ]
        ),
        ...spacer(1),

        h2("5.5 Why Pass 1 Over Later Passes?"),
        p("Although Pass 4 achieved a higher validation accuracy (~0.795), its mean F1 (0.685) was lower than Pass 1’s (0.726). Pass 4 also removed the walker class entirely. Since mean F1 better captures balanced performance across all classes (especially the more imbalanced walker class), Pass 1 was selected as the most comprehensively performant model. Later passes showed that aggressive regularisation (higher dropout, label smoothing) and data manipulation (downsampling, weighted sampling) tended to hurt rather than help on this dataset, suggesting the base MAE encoder already provided strong generalisation."),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 6: VLM VALIDATION
        // ══════════════════════════════════════════════════════════════════════
        h1("6. VLM Validation Study"),

        p("A key scientific question in this project is: how well do Vision-Language Models (VLMs) agree with human accessibility judgements? This matters both as validation of our Gemini-generated labels and as evidence for or against the reliability of VLM-based accessibility assessment in general."),

        p("The team conducted a formal correlation study comparing three VLMs against human survey scores from 18 real-world Montreal locations."),

        h2("6.1 Study Design"),
        metricsTable(
          ["Parameter", "Details"],
          [
            ["Reference scores", "Human survey ratings for electric wheelchair accessibility"],
            ["Locations tested", "18 Montreal locations (KR_01–KR_10, MH_02–MH_19)"],
            ["VLMs evaluated", "Gemma 3 4B, Qwen2-VL-2B, Gemini 2.5 Flash"],
            ["Statistical tests", "Spearman rank correlation, Pearson r, Kendall’s τ"],
            ["Significance threshold", "α = 0.05"],
          ]
        ),
        ...spacer(1),

        h2("6.2 Results"),
        metricsTable(
          ["Model", "Spearman ρ", "Pearson r", "Kendall τ", "Any Significant?"],
          [
            ["Gemma 3 4B", "−0.055 (p=0.828)", "−0.080 (p=0.753)", "−0.049 (p=0.820)", "No"],
            ["Qwen2-VL-2B", "+0.331 (p=0.180)", "+0.277 (p=0.266)", "+0.295 (p=0.173)", "No"],
            ["Gemini 2.5 Flash", "+0.190 (p=0.451)", "+0.173 (p=0.493)", "+0.160 (p=0.450)", "No"],
          ]
        ),
        ...spacer(1),

        h2("6.3 Interpretation"),
        p("None of the three VLMs showed statistically significant correlation with human survey scores for electric wheelchair accessibility at the α=0.05 level. Key observations:"),
        bullet([new TextRun({ text: "Qwen2-VL-2B showed the strongest positive correlation", font: "Arial", size: 22, bold: true }), new TextRun({ text: " (Spearman ρ=0.331), suggesting it may have the best alignment with human accessibility perception among those tested, though this did not reach significance.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Gemini 2.5 Flash showed moderate positive correlation", font: "Arial", size: 22, bold: true }), new TextRun({ text: " (ρ=0.190), which is notable given that Gemini was used to generate our training labels. The modest positive correlation is consistent with partial validity of the labels.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Gemma 3 4B showed near-zero and slightly negative correlations", font: "Arial", size: 22, bold: true }), new TextRun({ text: ", suggesting poor alignment with human accessibility judgements.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Sample size (n=18) is the primary limitation. ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "With 18 data points, the study is underpowered for detecting moderate correlations (ρ≈0.3 requires approximately n=55 for 80% power). The non-significant results do not imply zero correlation; they imply insufficient evidence to confirm or deny it.", font: "Arial", size: 22 })]),
        ...spacer(1),

        h2("6.4 Role in Label Validation"),
        p("Despite the non-significant correlations, the VLM validation study played an important role in quality assurance. The recalibration procedure (Section 3.3) was motivated in part by the observation that raw VLM scores diverge from human distributions. By rescaling Gemini scores to match the human survey’s mean and standard deviation before binarizing, the team ensured that the binary labels reflect human-calibrated thresholds rather than raw model outputs. This is an important methodological safeguard."),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 7: APPLICATION ARCHITECTURE
        // ══════════════════════════════════════════════════════════════════════
        h1("7. Application Architecture"),

        p("Urban Access is a full-stack application that integrates the ML pipeline into a user-facing routing tool. The following describes the end-to-end flow from user input to route recommendation."),

        h2("7.1 User Interface and Input"),
        p("The user specifies three inputs:"),
        bullet("Their mobility aid type (from the six supported categories)"),
        bullet("Origin location (address or map pin)"),
        bullet("Destination location (address or map pin)"),
        ...spacer(1),

        h2("7.2 Route Generation"),
        p("Multiple candidate routes between origin and destination are generated using pedestrian routing (based on OpenStreetMap/OSRM or equivalent). These candidate routes serve as the pool for accessibility evaluation and ranking."),

        h2("7.3 Image Retrieval"),
        p("For each candidate route, hundreds of street-level images are retrieved from Mapillary at regular intervals along the path. Images are fetched via the Mapillary API using the coordinate-based query approach developed during dataset creation."),

        h2("7.4 ViT Accessibility Scoring"),
        p("The fine-tuned ViT model (Pass 1) runs inference on each retrieved image, producing a binary accessibility prediction for the user’s selected mobility aid type. This gives a per-image accessibility signal that is aggregated to score each route segment."),

        h2("7.5 VLM Secondary Scoring"),
        p("The most inaccessible or uncertain route points are passed to Gemini 2.5 Flash for a higher-quality, contextual assessment. The VLM provides a rating and a natural-language justification for why a particular location is or is not accessible. This secondary scoring layer improves reliability at critical decision points (intersections, ramp transitions, construction zones) without the computational cost of running the VLM on every image."),

        h2("7.6 Auxiliary Data Integration"),
        p("The LLM route recommendation incorporates several additional data signals beyond image-based accessibility scores:"),
        bullet([new TextRun({ text: "Pedestrian safety data: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Intersection and crossing hazard information.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Construction data: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Active construction zones that may block or degrade sidewalk access.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Inclination/gradient data: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Steepness of route segments, particularly relevant for manual wheelchair and mobility scooter users.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "User feedback: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Historical ratings from previous users of similar routes, incorporated into the LLM prompt (see Section 8).", font: "Arial", size: 22 })]),
        ...spacer(1),

        h2("7.7 LLM Route Recommendation"),
        p("An open-source LLM synthesises all of the above signals and generates a ranked route recommendation. The output includes:"),
        bullet("The recommended primary route with a plain-language explanation"),
        bullet("Two alternative routes with comparative reasoning"),
        bullet("Specific accessibility warnings for notable barriers along each route"),
        ...spacer(1),

        h2("7.8 Infrastructure"),
        metricsTable(
          ["Component", "Technology"],
          [
            ["Backend", "Python (FastAPI or equivalent)"],
            ["ML inference", "PyTorch (ViT), Gemini API (VLM)"],
            ["Database", "Supabase (user accounts, feedback storage)"],
            ["Image source", "Mapillary API"],
            ["Hosting", "Google Cloud Platform (free tier, 3-month trial)"],
            ["Routing", "OpenStreetMap / OSRM"],
          ]
        ),
        ...spacer(1),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 8: FEEDBACK SYSTEM
        // ══════════════════════════════════════════════════════════════════════
        h1("8. User Feedback System"),

        h2("8.1 Feedback Collection"),
        p("After completing a journey, users can rate the accessibility of the route they used and provide qualitative comments. This feedback is stored securely in a Supabase database linked to the user’s account."),

        h2("8.2 Feedback Incorporation into Recommendations"),
        p("User feedback is incorporated into the LLM prompt that generates route recommendations. The weight given to a piece of feedback is proportional to how closely the previous user’s origin and destination match the current user’s query — feedback from a highly similar journey is weighted more heavily than feedback from a tangentially related one. This allows the system to personalise recommendations based on real-world user experiences without requiring a full model retraining cycle."),

        h2("8.3 Future Direction: Reinforcement Learning"),
        p("The current feedback integration is a prompt-engineering approach — effective but limited. The team has identified Reinforcement Learning (RL) as the appropriate long-term mechanism for incorporating user feedback at scale. RL would allow the route scoring and recommendation system to learn directly from user satisfaction signals, improving over time as more feedback accumulates. This remains a future development priority."),

        h2("8.4 Privacy and Data Handling"),
        p("User privacy is treated as a core design constraint, not an afterthought:"),
        bullet("All survey data collected during the label creation phase is fully anonymised"),
        bullet("User accounts and their associated feedback data are stored in a secure Supabase database"),
        bullet("Users can delete their account (and all associated feedback data) at any time through the application"),
        bullet("No personal movement data is retained beyond what is necessary for the feedback weighting calculation"),
        ...spacer(1),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 9: EVALUATION CRITERIA ASSESSMENT
        // ══════════════════════════════════════════════════════════════════════
        h1("9. Evaluation Criteria Assessment"),

        p("This section addresses the five Demo Day evaluation criteria explicitly, structured for panel review."),

        h2("9.1 Three-Minute Pitch Quality"),
        p("The pitch narrative was structured to move from lived experience to technical solution to social impact in under three minutes. Key rhetorical elements:"),
        bullet("Opening with the 200,000+ Montrealers statistic grounds the problem in concrete human scale"),
        bullet("The Rashid Mushkani insight (“a government building may be accessible, but the path to it often is not”) frames the routing focus intuitively"),
        bullet("The live demo walk-through of the app’s route recommendation provides a tangible, graspable output for a non-technical audience"),
        bullet("Future directions close with Canada-wide scale (4M+ potential users), demonstrating the project’s ambition beyond Montreal"),
        ...spacer(1),

        h2("9.2 Innovation"),
        p("Urban Access is innovative across several dimensions:"),
        bullet([new TextRun({ text: "Novel dataset: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "No publicly available dataset combines Mapillary street-level imagery with per-mobility-aid accessibility labels at 31,887 scale for Canadian cities.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "VLM-calibrated labelling: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "The recalibration methodology that aligns AI-generated scores to human survey distributions before binarization is a principled, novel contribution to AI-generated label quality control.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "MAE domain adaptation: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Self-supervised pre-training on domain-specific street-view images before supervised fine-tuning is state-of-the-art for sample-efficient visual classification.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Multi-signal LLM routing: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Combining ViT scores, VLM annotations, pedestrian safety, construction, inclination, and user feedback into a single LLM-generated route recommendation is a novel integration approach.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Feedback-weighted personalisation: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Weighting user feedback by origin/destination proximity for prompt construction is a practical, low-latency personalisation mechanism without model retraining.", font: "Arial", size: 22 })]),
        ...spacer(1),

        h2("9.3 Impact"),
        p("The potential impact of Urban Access is significant and measurable:"),
        bullet([new TextRun({ text: "Immediate beneficiaries: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "200,000+ Montrealers with mobility impairments who currently lack accessible routing tools.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Scalable to Canada: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "4M+ Canadians live with mobility impairments. The pipeline is city-agnostic and replicable wherever Mapillary coverage exists.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Quality of life and autonomy: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Accessible routing directly enables independence. The ability to plan a safe route without assistance is a meaningful quality-of-life improvement.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Policy evidence: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "The accessibility scores generated by Urban Access are a potential data source for city planners identifying infrastructure gaps — a secondary impact pathway.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Future scope: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Expansion to broader disability categories (visual impairment, hearing impairment) was identified during VLM labelling, where these categories were also scored.", font: "Arial", size: 22 })]),
        ...spacer(1),

        h2("9.4 Feasibility"),
        p("Urban Access is technically feasible and partially deployed. Evidence:"),
        bullet("A working end-to-end application was demonstrated at Demo Day"),
        bullet("31,887 labelled training samples generated; ViT trained and deployed"),
        bullet("Mapillary API provides free image access with reasonable coverage in urban cores"),
        bullet("Google Cloud hosts the backend; Supabase handles secure data storage"),
        bullet("All ML components are open-source or accessible via free API tiers"),
        ...spacer(1),
        p("Key feasibility constraints and mitigations:"),
        bullet([new TextRun({ text: "Mapillary coverage gaps: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Mapillary has excellent coverage in downtown Montreal but sparser coverage in suburban areas. Mitigation: seek funding to access Google Street View for production scale.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Hosting costs: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Google Cloud free tier expires after 3 months. Mitigation: identify stable hosting (grant funding, university partnership, or non-profit cloud credits).", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "MAE training incomplete: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "95 of 200 planned epochs completed. The deployed model is functional but would benefit from continued pre-training.", font: "Arial", size: 22 })]),
        ...spacer(1),

        h2("9.5 Ethics"),
        p("Ethical considerations were embedded throughout the project design:"),
        bullet([new TextRun({ text: "Community-led design: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "The routing focus, survey design, and label categories were all shaped by consultation with Rashid Mushkani and ongoing feedback from Laurence Parent. The system reflects genuine disability needs, not assumptions.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Label bias mitigation: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "The recalibration procedure reduces the risk that AI-generated labels systematically misrepresent human accessibility judgements. This is an active, methodological response to AI bias.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Anonymised survey data: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "No personally identifiable information was collected in the survey. Participants rated images only.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "User data control: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Users can delete their accounts and all associated feedback data at any time, ensuring meaningful data agency.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Secure storage: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "All user data is stored in Supabase, which provides enterprise-grade security including row-level access controls.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Representation: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Supporting six distinct mobility aid types ensures the system does not treat mobility impairment as monolithic. Different conditions have different requirements, and the model captures this.", font: "Arial", size: 22 })]),
        ...spacer(1),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 10: PANEL Q&A
        // ══════════════════════════════════════════════════════════════════════
        h1("10. Anticipated Panel Questions and Responses"),

        p("The following addresses specific questions likely to arise during panel evaluation."),

        h2("Q1: What did you learn from Rashid and Laurence specifically?"),
        infoBox([
          new Paragraph({
            children: [new TextRun({ text: "From Rashid Mushkani: ", font: "Arial", size: 22, bold: true, color: MID_BLU })],
            spacing: { before: 60, after: 40 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "The central insight was the concept of universal accessibility and the last-mile problem: even when destination buildings are accessible, the paths leading to them often are not. This directly shaped our decision to build a route recommendation system rather than a static accessibility map. Rashid also facilitated connections to Promo-Accès, which improved our survey distribution strategy.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 0, after: 120 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "From Laurence Parent: ", font: "Arial", size: 22, bold: true, color: MID_BLU })],
            spacing: { before: 0, after: 40 },
          }),
          new Paragraph({
            children: [new TextRun({ text: "Laurence provided iterative feedback throughout development, ensuring that our chosen mobility aid categories, the granularity of our scoring, and the nature of the route explanations accurately reflected what disabled individuals actually need from a navigation tool. Her engagement was ongoing — not a one-time consultation.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 0, after: 60 },
          }),
        ]),
        ...spacer(1),

        h2("Q2: How was the VLM validated against human judgement?"),
        infoBox([
          new Paragraph({
            children: [new TextRun({ text: "We conducted a formal correlation study comparing Gemma 3 4B, Qwen2-VL-2B, and Gemini 2.5 Flash against human survey scores for electric wheelchair accessibility across 18 Montreal locations. None reached statistical significance at n=18 (underpowered study), but Qwen2 showed the strongest positive correlation (ρ=0.331). More importantly, we mitigated VLM-human divergence through label recalibration: Gemini scores were z-score normalised to match the human survey distribution (same mean and standard deviation per aid type) before binarization. This ensures our binary labels reflect human-calibrated thresholds.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 60, after: 60 },
          }),
        ]),
        ...spacer(1),

        h2("Q3: How is user feedback incorporated?"),
        infoBox([
          new Paragraph({
            children: [new TextRun({ text: "User feedback is included in the prompt sent to the LLM when generating route recommendations. Feedback is weighted by how closely the previous user’s origin and destination match the current query — similar journeys receive more weight. This is a prompt-engineering approach that allows feedback to influence recommendations without requiring model retraining. In the future, we plan to implement Reinforcement Learning to incorporate feedback more systematically and at greater scale.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 60, after: 60 },
          }),
        ]),
        ...spacer(1),

        h2("Q4: How does the system handle user privacy?"),
        infoBox([
          new Paragraph({
            children: [new TextRun({ text: "Privacy is a first-class design constraint: (1) All survey data used for label calibration is fully anonymised. (2) User accounts and feedback are stored in Supabase with enterprise-grade security and row-level access controls. (3) Users can delete their account and all associated data at any time from within the application. No movement data is retained beyond what is needed for feedback weighting.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 60, after: 60 },
          }),
        ]),
        ...spacer(1),

        h2("Q5: Is this scalable beyond Montreal?"),
        infoBox([
          new Paragraph({
            children: [new TextRun({ text: "Yes, with the right resources. The pipeline is city-agnostic: Mapillary has global coverage and the image collection, labelling, and training pipeline can be replicated for any city. Mapillary provides excellent coverage in the downtown cores of many major cities, which is where the greatest density of mobility-impaired users tend to be. Scaling to the full Canadian scope (4M+ potential users) would likely require: (1) funding to access Google Street View for broader and higher-quality image coverage, (2) a stable hosting platform (Google Cloud free tier is limited to 3 months), and (3) access to city-specific auxiliary data such as snow removal schedules (Montreal’s snow removal API was under maintenance during our development period).", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 60, after: 60 },
          }),
        ]),
        ...spacer(1),

        h2("Q6: What would you do with more time and resources?"),
        infoBox([
          new Paragraph({
            children: [new TextRun({ text: "Priority investments with additional resources:", font: "Arial", size: 22, bold: true, color: MID_BLU })],
            spacing: { before: 60, after: 60 },
          }),
          new Paragraph({
            numbering: { reference: "bullets", level: 0 },
            children: [new TextRun({ text: "Google Street View access: Mapillary is excellent but Google Street View offers better coverage in lower-density areas. Licensing costs are the barrier.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 40, after: 40 },
          }),
          new Paragraph({
            numbering: { reference: "bullets", level: 0 },
            children: [new TextRun({ text: "Stable production hosting: Google Cloud free tier expires in 3 months. A grant or institutional partnership is needed for sustained deployment.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 40, after: 40 },
          }),
          new Paragraph({
            numbering: { reference: "bullets", level: 0 },
            children: [new TextRun({ text: "Snow removal data integration: Montreal’s snow removal API (currently under maintenance) would dramatically improve winter routing accuracy.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 40, after: 40 },
          }),
          new Paragraph({
            numbering: { reference: "bullets", level: 0 },
            children: [new TextRun({ text: "Complete MAE training: 95 of 200 planned epochs were run. Full training would likely improve ViT classification performance.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 40, after: 40 },
          }),
          new Paragraph({
            numbering: { reference: "bullets", level: 0 },
            children: [new TextRun({ text: "Reinforcement Learning for feedback: Replace prompt-weighting with a proper RL feedback loop.", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 40, after: 40 },
          }),
          new Paragraph({
            numbering: { reference: "bullets", level: 0 },
            children: [new TextRun({ text: "Real-time transit integration: Incorporate GTFS transit data so users can plan multi-modal accessible journeys (walk + metro).", font: "Arial", size: 22, color: DARK })],
            spacing: { before: 40, after: 60 },
          }),
        ]),
        ...spacer(1),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 11: LIMITATIONS AND FUTURE WORK
        // ══════════════════════════════════════════════════════════════════════
        h1("11. Limitations and Future Work"),

        h2("11.1 Current Limitations"),
        bullet([new TextRun({ text: "MAE training incomplete: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Pre-training ran 95/200 planned epochs. Further pre-training could improve ViT performance.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "VLM label quality: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Gemini-generated labels have limited correlation with human ground truth (n=18 validation). Larger-scale human labelling studies would strengthen confidence in label quality.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Mapillary coverage gaps: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Suburban and lower-density areas of Montreal have sparse Mapillary coverage.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Seasonal data: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Images collected primarily in summer/fall conditions. Winter snow and ice dramatically change accessibility and are not well-represented in current training data.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Hosting stability: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Google Cloud free tier has a 3-month expiry.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Walker class performance: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "The walker class consistently showed lower F1 performance, partly due to class imbalance (only 30.3% accessible) and the decision not to recalibrate its labels.", font: "Arial", size: 22 })]),
        ...spacer(1),

        h2("11.2 Future Development Roadmap"),
        bullet([new TextRun({ text: "Phase 1 — Stability: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Secure hosting funding; deploy stable production server; extend MAE training.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Phase 2 — Data Quality: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Integrate Google Street View for broader coverage; acquire snow removal data API; expand survey to increase ground truth.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Phase 3 — Intelligence: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Implement RL feedback loop; add GTFS transit integration; expand to visual and hearing impairment categories.", font: "Arial", size: 22 })]),
        bullet([new TextRun({ text: "Phase 4 — Scale: ", font: "Arial", size: 22, bold: true }), new TextRun({ text: "Replicate pipeline for additional Canadian cities; explore data-sharing partnerships with municipalities.", font: "Arial", size: 22 })]),
        ...spacer(1),

        // ══════════════════════════════════════════════════════════════════════
        // SECTION 12: SUMMARY
        // ══════════════════════════════════════════════════════════════════════
        h1("12. Summary"),

        p("Urban Access is a complete, end-to-end AI system for accessible pedestrian routing in Montreal. In the space of a single AI4Good Lab cohort, the team:"),
        bullet("Collected and labelled a dataset of 31,887 street-level images across 14,000 Montreal locations, annotated for six mobility aid types using a novel VLM-calibrated labelling pipeline"),
        bullet("Conducted community outreach with 30+ disability groups and key advisors (Rashid Mushkani, Laurence Parent) to ensure the system addresses real accessibility needs"),
        bullet("Pre-trained a Vision Transformer using Masked Autoencoder methodology on 95 epochs of domain-specific street-view imagery, achieving reconstruction loss of 0.38"),
        bullet("Fine-tuned the ViT through six experimental passes, achieving a best validation mean F1 of 0.726 with the deployed model (Pass 1)"),
        bullet("Built and deployed a full-stack application integrating ViT scoring, VLM secondary annotation, multi-signal LLM route recommendation, and user feedback weighting"),
        bullet("Implemented a privacy-first user feedback system with deletable accounts and secure Supabase storage"),
        bullet("Conducted a formal VLM validation study comparing three models against human ground truth"),
        ...spacer(1),

        p("The project demonstrates that AI can be a powerful tool for social inclusion — but only when technical design is grounded in genuine community need. Urban Access is a working system with a clear path to city-scale deployment."),

        ...spacer(2),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "— End of Report —", font: "Arial", size: 22, color: "888888", italics: true })],
          spacing: { before: 480, after: 0 },
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/sessions/sweet-friendly-einstein/mnt/outputs/Urban_Access_Project_Report.docx", buffer);
  console.log("Done");
});
