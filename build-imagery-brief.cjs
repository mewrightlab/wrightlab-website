// Generates WrightLab_ImageryBrief.docx
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, PageOrientation, LevelFormat, ExternalHyperlink,
  TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageBreak,
} = require('docx');

// ---------- helpers ----------
const PAGE = { W: 12240, H: 15840, MARGIN: 1440 };
const CONTENT_W = PAGE.W - PAGE.MARGIN * 2; // 9360

const ACCENT = '0AA5A4';
const ACCENT_DARK = '088987';
const SLATE = '5A6770';
const TEXT = '1A2329';
const ACCENT_LIGHT = 'E0F4F3';

const border = { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    ...opts,
    children: opts.children || [new TextRun({ text })],
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text })],
  });
}

function bullet(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: opts.level || 0 },
    spacing: { after: 60 },
    children: text instanceof Array ? text : [new TextRun({ text })],
  });
}

function number(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: 'numbers', level: opts.level || 0 },
    spacing: { after: 60 },
    children: text instanceof Array ? text : [new TextRun({ text })],
  });
}

function cell(content, opts = {}) {
  const width = opts.width;
  const paragraphs = Array.isArray(content)
    ? content
    : [new Paragraph({ children: [new TextRun({ text: String(content), bold: opts.bold || false })] })];
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR } : undefined,
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: paragraphs,
  });
}

function table(columnWidths, rows, headerRow = true) {
  return new Table({
    width: { size: columnWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths,
    rows: rows.map((row, idx) =>
      new TableRow({
        tableHeader: headerRow && idx === 0,
        children: row.map((c, i) =>
          cell(c, {
            width: columnWidths[i],
            bold: headerRow && idx === 0,
            shade: headerRow && idx === 0 ? ACCENT_LIGHT : undefined,
          })
        ),
      })
    ),
  });
}

// ---------- document content ----------

const children = [];

// Title block
children.push(
  new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 120 },
    children: [
      new TextRun({
        text: 'Wright Lab Website',
        bold: true,
        size: 56,
        color: TEXT,
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({
        text: 'Imagery Direction & Briefs',
        bold: true,
        size: 40,
        color: ACCENT_DARK,
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 80 },
    children: [
      new TextRun({
        text: 'Visual identity, hero image brief, and follow-up image briefs',
        italics: true,
        size: 24,
        color: SLATE,
      }),
    ],
  }),
  new Paragraph({
    spacing: { after: 60 },
    children: [
      new TextRun({ text: 'Prepared: ', bold: true, size: 22 }),
      new TextRun({ text: 'May 18, 2026', size: 22 }),
    ],
  }),
  new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({ text: 'Project: ', bold: true, size: 22 }),
      new TextRun({ text: 'mewrightlab.org (Astro static site, Spectronaut-inspired design)', size: 22 }),
    ],
  }),
  new Paragraph({ children: [new PageBreak()] })
);

// TOC
children.push(
  new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text: 'Contents', bold: true, size: 32, color: TEXT })],
  }),
  new TableOfContents('Table of Contents', { hyperlink: true, headingStyleRange: '1-3' }),
  new Paragraph({ children: [new PageBreak()] })
);

// 1. Overall Visual Concept
children.push(
  h1('Overall Visual Concept — "Cellular Cartography"'),
  p(
    "The site's unifying visual metaphor is cellular protein cartography: the lab maps spatial relationships between proteins the way Google Maps maps geography. The metaphor is scientifically accurate (proximity labeling literally produces spatial maps of protein neighborhoods) and visually distinctive."
  ),
  h2('Why this metaphor was chosen'),
  bullet('It is scientifically literal, not just decorative.'),
  bullet('It scales as a unifying language across all site imagery (hero, focus areas, database page, brand voice).'),
  bullet('It implies a searchable atlas — which extends naturally to the Database page concept.'),
  bullet('Visitors remember "the lab that does protein cartography" — distinctive and ownable.'),
  h2('Why it was chosen over the alternative concept'),
  p(
    'An alternative direction was considered: a hand shooting electrical current to mass spectrometers from Thermo, Bruker, and SCIEX, then hitting a plate of cells, then producing PPI graphs of protein complexes. This concept was set aside for three reasons:'
  ),
  number('Hero images need one strong visual idea; that composition tried to convey five (hand, electricity, three instruments, cells, PPI graph), which fragments the eye.'),
  number('Vendor logos (Thermo, Bruker, SCIEX) read as advertising rather than science to non-specialists, and as an unintentional product lineup to specialists.'),
  number('The "researcher as superhero" tone (a hand shooting lightning) is comic-book imagery and undercuts the scientific seriousness of the lab.')
);

// 2. Site-Wide Visual Language
children.push(
  h1('Site-Wide Visual Language'),
  h2('Color Palette'),
  table(
    [2400, 1400, 5560],
    [
      ['Role', 'Hex', 'Notes'],
      ['Primary teal accent', '#0aa5a4', 'Matches the site CSS accent. Use for highlighted protein clusters, key landmarks, the "you are here" pin.'],
      ['Teal hover / dark', '#088987', 'Slightly darker teal for depth.'],
      ['Accent light', '#e0f4f3', 'Background tints and hover states.'],
      ['Text primary', '#1a2329', 'Dark slate for any labels.'],
      ['Text muted / secondary', '#5a6770', 'Soft slate for "streets" and connections.'],
      ['Background', '#ffffff or #fafbfc', 'Off-white, crisp, not dark.'],
      ['Optional warm accent', 'dusty coral or amber', 'Only for one or two attention points if needed.'],
    ]
  ),
  h2('Typography'),
  bullet('Use the Inter font family (free, available from Google Fonts) for any labels embedded in images.'),
  bullet("Match the website's typography for visual continuity."),
  h2('Style restraint — what to avoid'),
  bullet('Heavy gradients.'),
  bullet('Drop shadows on everything.'),
  bullet('Photorealistic rendering.'),
  bullet('Vibrant rainbows.'),
  bullet('Cluttered compositions.'),
  p('Target aesthetic: clean, scientific, Google Maps default light theme energy.', {
    spacing: { before: 80, after: 200 },
  })
);

// 3. Hero Image Brief
children.push(
  h1('Hero Image Brief — "Cellular Atlas"'),
  h2('Purpose'),
  p('The first image visitors see at the top of the homepage. Sets the entire visual tone for the site.'),
  h2('Concept'),
  p(
    'A stylized top-down map view of a single cell rendered in the visual language of digital cartography. Organelles read as districts of a city; proximal proteins read as clusters of points of interest connected by faint "streets" representing proximity relationships.'
  ),
  h2('Technical specs'),
  table(
    [3000, 6360],
    [
      ['Spec', 'Value'],
      ['Dimensions', '1600 × 900 px (16:9 aspect ratio)'],
      ['Format', 'PNG (preferred for vector-style rendering) or JPG'],
      ['Color mode', 'sRGB'],
      ['File name', 'hero-atlas.png'],
      ['Save to folder', 'C:\\Users\\mkwrg\\Documents\\wrightlab-website\\public\\images\\'],
    ]
  ),
  h2('Composition suggestions'),
  number('The cell as the canvas — outline like a coastline, soft and organic.'),
  number('Organelles as labeled districts:'),
  bullet('Nucleus (large, centered or upper-left).', { level: 1 }),
  bullet('ER + Golgi (manufacturing district).', { level: 1 }),
  bullet('Mitochondria (scattered "power plants").', { level: 1 }),
  bullet('Cytoskeleton as faint "transit lines" running across the map.', { level: 1 }),
  bullet('Plasma membrane = the coastline.', { level: 1 }),
  number('Protein complexes as clusters of dots — small filled circles, with teal connecting lines between proximal ones (this is the "proximity" idea made visual).'),
  number('A single "highlight" cluster in bold teal with a pin or callout — this is the rhetorical focal point ("this is what we map").'),
  number('Subtle map-style touches: a tiny compass rose in the corner, a tiny scale bar, OR a faint grid — but use ONE of these, not all. Restraint.'),
  h2('What to skip on v1'),
  bullet('Do not try to render actual specific proteins (AR, FOXA1, etc.) — abstract clusters are fine.'),
  bullet('Do not add a key/legend — that overcomplicates a hero.'),
  bullet('Do not try to depict experimental workflow (labeling step, MS, etc.) — that is for the focus area figures.')
);

// 4. Focus Area 1
children.push(
  h1('Focus Area 1 Brief — "Protein Neighborhood"'),
  h2('Purpose'),
  p('Lives on the homepage and the Research page next to Focus Area 1 text (proximity interaction networks).'),
  h2('Concept'),
  p('Zoom-in view of one protein complex in the cellular atlas style. Like zooming in on a single neighborhood on Google Maps.'),
  h2('Technical specs'),
  table(
    [3000, 6360],
    [
      ['Spec', 'Value'],
      ['Dimensions', '1200 × 900 px (4:3 aspect ratio)'],
      ['Format', 'PNG'],
      ['File name', 'focus-1-neighborhood.png'],
      ['Save to', 'C:\\Users\\mkwrg\\Documents\\wrightlab-website\\public\\images\\'],
    ]
  ),
  h2('Composition'),
  bullet('Central highlighted protein (teal), surrounded by proximal partners shown as dots.'),
  bullet('Faint dashed lines showing proximity-labeling "reach" radius.'),
  bullet('Maintain same color palette as hero.')
);

// 5. Focus Area 2
children.push(
  h1('Focus Area 2 Brief — "Cartographer\'s Toolkit"'),
  h2('Purpose'),
  p('Lives on the homepage and the Research page next to Focus Area 2 text (novel proximity enzymes).'),
  h2('Concept'),
  p('Side-by-side comparison of different proximity enzymes (BioID, TurboID, APEX2) rendered as different "surveying tools" with different reach radii.'),
  h2('Technical specs'),
  table(
    [3000, 6360],
    [
      ['Spec', 'Value'],
      ['Dimensions', '1200 × 900 px (4:3 aspect ratio)'],
      ['Format', 'PNG'],
      ['File name', 'focus-2-toolkit.png'],
      ['Save to', 'C:\\Users\\mkwrg\\Documents\\wrightlab-website\\public\\images\\'],
    ]
  ),
  h2('Composition'),
  bullet('Each enzyme rendered as a distinct "tool" or instrument with characteristic visual treatment.'),
  bullet('Show the labeling radius as a faint circle around each.'),
  bullet('Same color palette and restraint as hero.')
);

// 6. Focus Area 3
children.push(
  h1('Focus Area 3 Brief — "Resolution Levels"'),
  h2('Purpose'),
  p('Lives on the homepage and the Research page next to Focus Area 3 text (reagents and methods).'),
  h2('Concept'),
  p('A diagonal split showing the same molecular landscape at different "zoom levels" — population to cell to complex to protein. Like Google Earth zoom but for biology.'),
  h2('Technical specs'),
  table(
    [3000, 6360],
    [
      ['Spec', 'Value'],
      ['Dimensions', '1200 × 900 px (4:3 aspect ratio)'],
      ['Format', 'PNG'],
      ['File name', 'focus-3-resolution.png'],
      ['Save to', 'C:\\Users\\mkwrg\\Documents\\wrightlab-website\\public\\images\\'],
    ]
  ),
  h2('Composition'),
  bullet('Progressive zoom levels — possibly arranged in panels or as a single continuous zoom.'),
  bullet('Same cartography visual language.'),
  bullet('Same color palette.')
);

// 7. Future Imagery Ideas
children.push(
  h1('Future Imagery Ideas'),
  bullet('Database page: could feature an interactive map-like interface for browsing data — long-term project.'),
  bullet('Team page: square headshots of lab members (not part of the cartography language, but should use clean white backgrounds and the same teal accent for consistency).'),
  bullet('Brand extension: the lab could be branded as "cellular cartographers" — long-term identity opportunity.')
);

// 8. Workflow Notes
children.push(
  h1('Workflow Notes'),
  bullet('Tool: Affinity Designer (vector) and Affinity Photo (raster) — both owned by Michael.'),
  bullet('Process: draft each image in Affinity, export to spec, drop file into public/images/ folder, then notify Claude to wire it into the site.'),
  bullet('Iteration: do not try to nail v1 — get a draft, see it in context, iterate.'),
  bullet('Recommended order: Hero first (highest impact, sets tone), then Focus Areas 1–3 in sequence.'),
  bullet('File naming convention: lowercase, hyphenated, descriptive (e.g., hero-atlas.png, focus-1-neighborhood.png).')
);

// 9. Appendix
children.push(
  h1('Appendix — Where Each Image Lives on the Site'),
  table(
    [3000, 3000, 3360],
    [
      ['Image', 'Page(s)', 'Section'],
      ['hero-atlas.png', '/ (Home)', 'Top hero, right column'],
      ['focus-1-neighborhood.png', '/ (Home) and /research/', 'Focus Area 1'],
      ['focus-2-toolkit.png', '/ (Home) and /research/', 'Focus Area 2'],
      ['focus-3-resolution.png', '/ (Home) and /research/', 'Focus Area 3'],
      ['michael-wright.jpg', '/about/', 'PI portrait (already in place)'],
      ['Team headshots (TBD)', '/team/', 'Member cards'],
    ]
  )
);

// ---------- document ----------
const doc = new Document({
  creator: 'Wright Lab Project',
  title: 'Wright Lab Website — Imagery Direction & Briefs',
  features: { updateFields: true },
  styles: {
    default: {
      document: {
        run: { font: 'Arial', size: 22, color: TEXT },
        paragraph: { spacing: { line: 320 } },
      },
    },
    paragraphStyles: [
      {
        id: 'Heading1',
        name: 'Heading 1',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { size: 36, bold: true, font: 'Arial', color: TEXT },
        paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 },
      },
      {
        id: 'Heading2',
        name: 'Heading 2',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: ACCENT_DARK },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 },
      },
      {
        id: 'Heading3',
        name: 'Heading 3',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: TEXT },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: '•',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
          {
            level: 1,
            format: LevelFormat.BULLET,
            text: '◦',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1440, hanging: 360 } } },
          },
        ],
      },
      {
        reference: 'numbers',
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: '%1.',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: PAGE.W, height: PAGE.H },
          margin: { top: PAGE.MARGIN, right: PAGE.MARGIN, bottom: PAGE.MARGIN, left: PAGE.MARGIN },
        },
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  const outPath = 'C:\\Users\\mkwrg\\Documents\\wrightlab-website\\WrightLab_ImageryBrief.docx';
  fs.writeFileSync(outPath, buffer);
  console.log('Wrote:', outPath, '(' + buffer.length + ' bytes)');
});
