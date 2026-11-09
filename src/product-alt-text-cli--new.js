#!/usr/bin/env node

import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import ExcelJS from 'exceljs';

const DEFAULT_MODEL = 'qwen3-vl:4b-instruct';
const DEFAULT_MAX_CHARS = 1000;
const IMAGE_EXTENSIONS = new Set([
  '.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.tif', '.tiff'
]);

const DEFAULT_PROMPT = `You are writing accessibility alt text for an e-commerce product image.

Product name: {{productName}}
Image filename: {{filename}}
Maximum length: {{maxChars}} characters.

Context hint:
{{context}}

Write concise, useful alt text that identifies the product and describes only visually meaningful details that help a person who cannot see the image understand it.

Rules:
- Treat the supplied product name and context hint as authoritative metadata, but do not invent visual details.
- Describe the product type, visible color/shade, packaging/container, orientation, and other distinguishing visible characteristics when useful.
- Include visible text only when it helps identify or understand the product.
- Do not infer ingredients, materials, benefits, quality, demographic suitability, or other claims unless they are explicitly provided in the context hint.
- Do not begin with "image of", "photo of", "picture of", or similar wording.
- Do not use promotional or subjective language.
- Do not mention the filename or cell reference.
- Stay at or below {{maxChars}} characters.
- Return only the alt text, with no quotation marks, labels, markdown, explanation, or commentary.`;

const IMAGE_TEXT_INFO_PROMPT = `You are analyzing an image for accessibility purposes.

Your task is to determine whether the image contains meaningful text that should
be provided separately as an accessible text equivalent.

IMPORTANT DEFINITIONS AND RULES

1. TEXT THAT IS PART OF A PRODUCT OR PHYSICAL OBJECT

Do NOT classify text as an "image of text" merely because words are visible on
a photographed or rendered product, package, label, bottle, box, sign, shirt,
book cover, device, or other physical object.

Examples that should normally be ignored for this field:
- a brand name printed on a lipstick tube
- ingredients printed on product packaging
- a shade name printed on a bottle
- text printed on a box shown as part of a product photograph
- a logo printed on the product

This text may still be useful when generating normal alt text, but it does not
belong in the separate image-text field merely because it is visible.

2. IMAGES WHOSE PRIMARY PURPOSE IS TO COMMUNICATE WRITTEN TEXT

If written text is itself an important part of the content being communicated,
classify the image as "image-of-text".

Examples include:
- promotional graphics containing headings and explanatory copy
- quotations rendered as graphics
- instructions rendered as an image
- announcements
- advertising copy presented graphically
- screenshots where the textual content itself is important

For these images, DO NOT DESCRIBE OR SUMMARIZE THE TEXT.

TRANSCRIBE THE MEANINGFUL TEXT VERBATIM.

Requirements:
- Preserve the original wording.
- Preserve punctuation where visible.
- Preserve numbers, percentages, units, symbols, and footnote markers.
- Preserve meaningful capitalization when practical.
- Preserve the intended reading order.
- Combine line breaks into normal sentences or paragraphs when the original
  text is clearly meant to be read continuously.
- Do not correct grammar or spelling.
- Do not paraphrase.
- Do not summarize.
- Do not explain what the text says.
- Do not describe the typography, layout, background, colors, or decorative
  imagery.
- Do not describe accompanying product swatches or decorative graphics unless
  they communicate information that cannot be understood from the text alone.

For example, if the image contains:

HYALURONIC ACID
and ARGAN OIL
provide 6-hour hydration*

return:

HYALURONIC ACID and ARGAN OIL provide 6-hour hydration*

Do NOT return:

"The image contains text explaining that hyaluronic acid and argan oil provide
6-hour hydration."

3. CHARTS, GRAPHS, TABLES, DIAGRAMS, AND DATA INFOGRAPHICS

If the image is primarily a chart, graph, data visualization, table, diagram,
process graphic, or infographic where text and visual relationships work
together to communicate information, classify it as "complex-image".

For a complex image, DO NOT simply transcribe all visible words.

Instead, provide a complete textual equivalent of the information communicated
by the graphic.

Include all meaningful information necessary to understand it, including when
applicable:
- title
- subject
- categories
- axes
- axis labels
- units
- legends
- series
- values
- percentages
- dates
- quantities
- rankings
- comparisons
- trends
- relationships
- stages or steps
- annotations
- exceptions
- notable highs and lows
- conclusions directly supported by the graphic

Do not omit meaningful data merely to make the response shorter.

4. MIXED CONTENT

If an image contains both:
- substantial standalone written text that should be transcribed, AND
- a chart, graph, diagram, or other visualization that conveys additional
  information,

classify it as "mixed".

For "mixed":
- transcribe the standalone written content accurately, and
- provide a complete textual equivalent of the visualized information.

5. NO MEANINGFUL IMAGE TEXT

If none of the above applies, classify the image as "none" and return an empty
content string.

6. NEVER USE INTRODUCTORY IMAGE PHRASES

The content must NEVER begin with phrases such as:

- "Image of"
- "An image of"
- "Photo of"
- "A photo of"
- "Photograph of"
- "Picture of"
- "A picture of"
- "Graphic of"
- "Illustration of"
- "Screenshot of"
- "The image contains"
- "The image shows"
- "The picture shows"
- "This image contains"
- "This image shows"

Start directly with the text or information the user needs.

Return exactly one JSON object and nothing else, using this schema:
{
  "hasRelevantInformation": true,
  "kind": "image-of-text|complex-image|mixed|none",
  "content": "accessible text equivalent"
}

If there is no qualifying image of text and no chart/graph/table/diagram/infographic that needs a textual equivalent, return:
{
  "hasRelevantInformation": false,
  "kind": "none",
  "content": ""
}`;

const IMAGE_TEXT_INFO_SCHEMA = {
  type: 'object',
  properties: {
    hasRelevantInformation: { type: 'boolean' },
    kind: {
      type: 'string',
      enum: ['image-of-text', 'complex-image', 'mixed', 'none'],
    },
    content: { type: 'string' },
  },
  required: ['hasRelevantInformation', 'kind', 'content'],
};

function printHelp() {
  console.log(`envision

Generate accessibility alt text for product images using a local Ollama vision model.

Usage:
  envision <image-root> [options]

Expected input layout:
  image-root/
    Product A/
      A1.jpg
      A2.png
    Product B/
      B1.jpg

The immediate parent directory of each image is used as the product name.

Options:
  -o, --output <file>        CSV output path (default: <image-root>/alt-text.csv)
  -m, --model <name>         Ollama model (default: ${DEFAULT_MODEL})
      --max-chars <number>   Maximum alt-text length (default: ${DEFAULT_MAX_CHARS})
      --context <text>       Global context hint supplied on the command line
      --context-file <file>  Global context hint file
      --prompt <text>        Override the default prompt template
      --prompt-file <file>   Read the prompt template from a file
      --temperature <n>      Ollama temperature (default: 0.2)
      --host <url>           Ollama API URL (default: http://127.0.0.1:11434)
      --retries <number>     Retry attempts if output is too long (default: 2)
      --dry-run              Discover images and show configuration without Ollama calls
      --verbose              Print the rendered prompt for each image
      --excel                Export as Excel (.xlsx) instead of CSV
      --image-text-info      Add WCAG image-of-text / complex-image information
  -h, --help                 Show this help

Context precedence/composition:
  1. <image-root>/context.md, if present
  2. --context-file, if supplied
  3. --context, if supplied
  4. <product-directory>/context.md, if present (product-specific)

All available context blocks are appended together.

Prompt placeholders:
  {{productName}}  Product directory name
  {{filename}}     Image filename
  {{maxChars}}     Configured character limit
  {{context}}      Combined context hint
`);
}

function parseArgs(argv) {
  const options = {
    input: null,
    output: null,
    model: DEFAULT_MODEL,
    maxChars: DEFAULT_MAX_CHARS,
    context: '',
    contextFile: null,
    prompt: null,
    promptFile: null,
    temperature: 0.2,
    retries: 2,
    dryRun: false,
    verbose: false,
    host: 'http://127.0.0.1:11434',
    excel: false,
    imageTextInfo: false,
  };

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];

    if (arg === '-h' || arg === '--help') {
      printHelp();
      process.exit(0);
    }

    if (!arg.startsWith('-') && !options.input) {
      options.input = arg;
      continue;
    }

    const requireValue = () => {
      if (i + 1 >= argv.length) {
        throw new Error(`Missing value for ${arg}`);
      }
      return argv[++i];
    };

    switch (arg) {
      case '-o':
      case '--output':
        options.output = requireValue();
        break;
      case '-m':
      case '--model':
        options.model = requireValue();
        break;
      case '--max-chars':
        options.maxChars = Number.parseInt(requireValue(), 10);
        break;
      case '--context':
        options.context = requireValue();
        break;
      case '--context-file':
        options.contextFile = requireValue();
        break;
      case '--prompt':
        options.prompt = requireValue();
        break;
      case '--prompt-file':
        options.promptFile = requireValue();
        break;
      case '--temperature':
        options.temperature = Number.parseFloat(requireValue());
        break;
      case '--host':
        options.host = requireValue().replace(/\/$/, '');
        break;
      case '--retries':
        options.retries = Number.parseInt(requireValue(), 10);
        break;
      case '--dry-run':
        options.dryRun = true;
        break;
      case '--verbose':
        options.verbose = true;
        break;
      case '--excel':
        options.excel = true;
        break;
      case '--image-text-info':
        options.imageTextInfo = true;
        break;
      default:
        throw new Error(`Unknown argument: ${arg}`);
    }
  }

  if (!options.input) {
    throw new Error('Missing image-root directory. Use --help for usage.');
  }

  if (!Number.isInteger(options.maxChars) || options.maxChars < 1) {
    throw new Error('--max-chars must be a positive integer.');
  }

  if (!Number.isInteger(options.retries) || options.retries < 0) {
    throw new Error('--retries must be zero or a positive integer.');
  }

  if (!Number.isFinite(options.temperature) || options.temperature < 0) {
    throw new Error('--temperature must be zero or a positive number.');
  }

  return options;
}

async function fileExists(filename) {
  try {
    await fs.access(filename);
    return true;
  } catch {
    return false;
  }
}

async function readTextIfExists(filename) {
  if (!(await fileExists(filename))) return '';
  return (await fs.readFile(filename, 'utf8')).trim();
}

function joinContextBlocks(blocks) {
  return blocks
    .map((block) => block?.trim())
    .filter(Boolean)
    .join('\n\n');
}

function renderTemplate(template, values) {
  return template.replace(/{{\s*(productName|filename|maxChars|context)\s*}}/g, (_, key) => {
    return String(values[key] ?? '');
  });
}

function cleanAltText(text) {
  let result = String(text ?? '').trim();

  // Remove common wrappers despite explicit prompt instructions.
  result = result.replace(/^```(?:text)?\s*/i, '').replace(/\s*```$/i, '').trim();
  result = result.replace(/^(?:alt\s*text|alt|description)\s*:\s*/i, '').trim();

  if (
    (result.startsWith('"') && result.endsWith('"')) ||
    (result.startsWith("'") && result.endsWith("'"))
  ) {
    result = result.slice(1, -1).trim();
  }

  // CSV is one record per image; normalize model line breaks to spaces.
  return result.replace(/\s+/g, ' ').trim();
}

function csvEscape(value) {
  const text = String(value ?? '');
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function imageSort(a, b) {
  return a.productName.localeCompare(b.productName, undefined, { numeric: true }) ||
    a.filename.localeCompare(b.filename, undefined, { numeric: true });
}

async function discoverImages(rootDirectory) {
  const rootEntries = await fs.readdir(rootDirectory, { withFileTypes: true });
  const images = [];

  for (const entry of rootEntries) {
    if (!entry.isDirectory()) continue;

    const productDirectory = path.join(rootDirectory, entry.name);
    const productEntries = await fs.readdir(productDirectory, { withFileTypes: true });

    for (const imageEntry of productEntries) {
      if (!imageEntry.isFile()) continue;

      const extension = path.extname(imageEntry.name).toLowerCase();
      if (!IMAGE_EXTENSIONS.has(extension)) continue;

      images.push({
        productName: entry.name,
        productDirectory,
        filename: imageEntry.name,
        absolutePath: path.join(productDirectory, imageEntry.name),
        relativePath: path.relative(rootDirectory, path.join(productDirectory, imageEntry.name)),
        cell: path.parse(imageEntry.name).name,
      });
    }
  }

  return images.sort(imageSort);
}

async function ollamaRequest(host, endpoint, init = {}) {
  let response;
  try {
    response = await fetch(`${host}${endpoint}`, init);
  } catch (error) {
    throw new Error(`Could not connect to Ollama at ${host}: ${error.message}`);
  }

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Ollama returned HTTP ${response.status}${text ? `: ${text}` : ''}`);
  }

  return response.json();
}

async function ollamaChat(host, request) {
  return ollamaRequest(host, '/api/chat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ...request, stream: false }),
  });
}

async function checkOllama(host, model) {
  const data = await ollamaRequest(host, '/api/tags');
  const names = new Set((data.models ?? []).map((item) => item.name));

  if (!names.has(model)) {
    const baseName = model.includes(':') ? model.split(':')[0] : model;
    const related = [...names].filter((name) => name.startsWith(`${baseName}:`) || name === baseName);
    const extra = related.length ? `\nRelated installed models: ${related.join(', ')}` : '';
    throw new Error(`Model "${model}" is not installed in Ollama. Run:\n  ollama pull ${model}${extra}`);
  }
}

async function generateOne({ image, prompt, model, maxChars, temperature, retries, host }) {
  const imageBytes = await fs.readFile(image.absolutePath);
  const base64Image = imageBytes.toString('base64');

  let response = await ollamaChat(host, {
    model,
    messages: [{ role: 'user', content: prompt, images: [base64Image] }],
    think: false,
    options: { temperature },
  });

  let altText = cleanAltText(response.message?.content);
  let retryCount = 0;

  while (altText.length > maxChars && retryCount < retries) {
    retryCount += 1;

    const shorteningPrompt = `Rewrite the following accessibility alt text so it is no more than ${maxChars} characters. Preserve the most useful identifying and visually meaningful information. Do not add new facts. Return only the revised alt text.\n\n${altText}`;

    response = await ollamaChat(host, {
      model,
      messages: [{ role: 'user', content: shorteningPrompt }],
      think: false,
      options: { temperature: 0.1 },
    });

    altText = cleanAltText(response.message?.content);
  }

  let status = 'ok';

  if (!altText) {
    status = 'empty';
  } else if (altText.length > maxChars) {
    altText = altText.slice(0, maxChars).trim();
    status = 'trimmed';
  }

  return { altText, status, retryCount };
}

function cleanImageTextInfo(text) {
  return String(text ?? '')
    .replace(/^```(?:json|text)?\s*/i, '')
    .replace(/\s*```$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseImageTextInfoResponse(rawResponse) {
  let cleaned = String(rawResponse ?? '').trim();
  cleaned = cleaned
    .replace(/^```(?:json|text)?\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();

  // Models occasionally add a short preface despite instructions. Extract the
  // outer JSON object so that prose around an otherwise valid response does not
  // end up in the spreadsheet.
  const firstBrace = cleaned.indexOf('{');
  const lastBrace = cleaned.lastIndexOf('}');
  const jsonText = firstBrace >= 0 && lastBrace > firstBrace
    ? cleaned.slice(firstBrace, lastBrace + 1)
    : cleaned;

  try {
    const parsed = JSON.parse(jsonText);
    const hasRelevantInformation = parsed?.hasRelevantInformation === true;
    const kind = String(parsed?.kind ?? '').toLowerCase();

    if (!hasRelevantInformation || kind === 'none') {
      return '';
    }

    return cleanImageTextInfo(parsed?.content);
  } catch {
    // Backward-compatible fallback for a model that ignores the JSON schema.
    // Treat explicit "none" / "no qualifying..." answers as empty; otherwise
    // keep the cleaned prose rather than losing potentially useful information.
    const fallback = cleanImageTextInfo(cleaned);
    if (
      !fallback ||
      /^(?:none|null|n\/?a|no qualifying image(?:s)? of text|no relevant information)\.?$/i.test(fallback)
    ) {
      return '';
    }
    return fallback;
  }
}

// --- Generate WCAG image-of-text / complex-image information using AI model ---
async function generateImageTextInfo({ image, model, host }) {
  const imageBytes = await fs.readFile(image.absolutePath);
  const base64Image = imageBytes.toString('base64');

  const response = await ollamaChat(host, {
    model,
    messages: [{ role: 'user', content: IMAGE_TEXT_INFO_PROMPT, images: [base64Image] }],
    format: IMAGE_TEXT_INFO_SCHEMA,
    think: false,
    options: { temperature: 0 },
  });

  return parseImageTextInfoResponse(response.message?.content);
}

// --- Excel Export Function ---
async function exportToExcel(images, outputFile, includeImageTextInfo) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'envision';
  workbook.created = new Date();

  // Group images by productName (directory)
  const grouped = images.reduce((acc, img) => {
    if (!acc[img.productName]) acc[img.productName] = [];
    acc[img.productName].push(img);
    return acc;
  }, {});

  for (const [productName, imgs] of Object.entries(grouped)) {
    const sheet = workbook.addWorksheet(productName, {
      properties: { tabColor: { argb: 'FFC0000' } },
      views: [{ state: 'frozen', ySplit: 1 }],
    });

    // Header row
    const columns = [
      { header: 'Filename', key: 'filename', width: 30 },
      { header: 'Image', key: 'image', width: 20 },
      { header: 'Alt Text', key: 'altText', width: 60, style: { alignment: { wrapText: true, vertical: 'top' } } },
    ];
    if (includeImageTextInfo) {
      columns.push({ header: 'Image Text Info', key: 'imageTextInfo', width: 60, style: { alignment: { wrapText: true, vertical: 'top' } } });
    }
    columns.push(
      { header: 'Character Count', key: 'characterCount', width: 18 },
      { header: 'Model', key: 'model', width: 18 },
      { header: 'Status', key: 'status', width: 18 }
    );
    sheet.columns = columns;
    sheet.getRow(1).font = { bold: true };

    for (const img of imgs) {
      const { filename, absolutePath, altText, characterCount, model, status, imageTextInfo } = img;

      // Use ExcelJS column keys rather than hard-coded/offset column letters.
      // This keeps the values aligned whether or not the optional
      // imageTextInfo column is present.
      const rowValues = {
        filename,
        altText,
        characterCount,
        model,
        status,
      };
      if (includeImageTextInfo) {
        rowValues.imageTextInfo = imageTextInfo || '';
      }

      const row = sheet.addRow(rowValues);
      const rowIdx = row.number;

      // Insert image if exists (as before)
      try {
        const imgBuffer = await fs.readFile(absolutePath);
        const ext = path.extname(filename).toLowerCase();
        let excelImgType = 'png';
        if (ext === '.jpg' || ext === '.jpeg') excelImgType = 'jpeg';
        const imgId = workbook.addImage({
          buffer: imgBuffer,
          extension: excelImgType,
        });
        // Place image in cell B{rowIdx}
        sheet.addImage(imgId, {
          tl: { col: 1, row: rowIdx - 1 }, // B column, current row
          ext: { width: 80, height: 80 },
          editAs: 'oneCell',
        });
        // Adjust row height for image
        row.height = 60;
      } catch (err) {
        // If image can't be read, leave cell blank but log error
        console.error(`Warning: Could not embed image for ${filename}: ${err.message}`);
      }
    }

    // Set alt text and image text info columns to wrap and auto-fit
    sheet.getColumn('C').alignment = { wrapText: true, vertical: 'top' };
    if (includeImageTextInfo) {
      sheet.getColumn('D').alignment = { wrapText: true, vertical: 'top' };
    }
    let maxAltLen = Math.max(...imgs.map(img => (img.altText || '').length), 10);
    sheet.getColumn('C').width = Math.min(Math.max(30, maxAltLen / 2), 100);
    if (includeImageTextInfo) {
      let maxInfoLen = Math.max(...imgs.map(img => (img.imageTextInfo || '').length), 10);
      sheet.getColumn('D').width = Math.min(Math.max(30, maxInfoLen / 2), 100);
    }
  }

  await fs.mkdir(path.dirname(outputFile), { recursive: true });
  await workbook.xlsx.writeFile(outputFile);
  console.log(`Excel file saved to ${outputFile}`);
}

// --- Main CLI ---
async function main() {
  const options = parseArgs(process.argv.slice(2));
  const inputRoot = path.resolve(options.input);
  const stat = await fs.stat(inputRoot).catch(() => null);

  if (!stat?.isDirectory()) {
    throw new Error(`Input directory does not exist or is not a directory: ${inputRoot}`);
  }

  const outputFile = path.resolve(
    options.output ?? path.join(inputRoot, options.excel ? 'alt-text.xlsx' : 'alt-text.csv')
  );

  const rootContext = await readTextIfExists(path.join(inputRoot, 'context.md'));
  const fileContext = options.contextFile
    ? (await fs.readFile(path.resolve(options.contextFile), 'utf8')).trim()
    : '';
  const globalContext = joinContextBlocks([rootContext, fileContext, options.context]);

  const promptTemplate = options.promptFile
    ? await fs.readFile(path.resolve(options.promptFile), 'utf8')
    : (options.prompt ?? DEFAULT_PROMPT);

  const images = await discoverImages(inputRoot);

  if (images.length === 0) {
    console.log(`No supported images found under product subdirectories in ${inputRoot}`);
    return;
  }

  console.log(`Input:      ${inputRoot}`);
  console.log(`Images:     ${images.length}`);
  console.log(`Model:      ${options.model}`);
  console.log(`Max chars:  ${options.maxChars}`);
  console.log(`Output:     ${outputFile}`);
  console.log(`Ollama:     ${options.host}`);
  console.log(`Context:    ${globalContext ? 'yes' : 'none'}`);
  if (options.imageTextInfo) {
    console.log(`Image Text Info: enabled`);
  }

  if (options.dryRun) {
    console.log('\nDry run — no Ollama calls will be made.\n');
    for (const image of images) {
      const productContext = await readTextIfExists(path.join(image.productDirectory, 'context.md'));
      const context = joinContextBlocks([globalContext, productContext]) || '(none supplied)';
      console.log(`${image.productName} | ${image.filename} | context=${context === '(none supplied)' ? 'none' : 'yes'}`);
    }
    return;
  }

  // Fail early if Ollama is unavailable or the requested model is not installed.
  await checkOllama(options.host, options.model);

  const rows = [];

  for (let index = 0; index < images.length; index += 1) {
    const image = images[index];
    const productContext = await readTextIfExists(path.join(image.productDirectory, 'context.md'));
    const context = joinContextBlocks([globalContext, productContext]) || '(none supplied)';

    const prompt = renderTemplate(promptTemplate, {
      productName: image.productName,
      filename: image.filename,
      maxChars: options.maxChars,
      context,
    });

    console.log(`[${index + 1}/${images.length}] ${image.productName}/${image.filename}`);
    if (options.verbose) {
      console.log('----- prompt -----');
      console.log(prompt);
      console.log('------------------');
    }

    let imageTextInfo = '';
    if (options.imageTextInfo) {
      try {
        imageTextInfo = await generateImageTextInfo({
          image,
          model: options.model,
          host: options.host,
        });
        if (options.verbose) {
          console.log('----- image text info -----');
          console.log(imageTextInfo);
          console.log('--------------------------');
        }
      } catch (err) {
        console.error(`  ERROR: Failed to generate image text info: ${err.message}`);
        imageTextInfo = '';
      }
    }

    try {
      const result = await generateOne({
        image,
        prompt,
        model: options.model,
        maxChars: options.maxChars,
        temperature: options.temperature,
        retries: options.retries,
        host: options.host,
      });

      rows.push({
        productName: image.productName,
        cell: image.cell,
        filename: image.filename,
        relativePath: image.relativePath,
        absolutePath: image.absolutePath,
        altText: result.altText,
        characterCount: result.altText.length,
        model: options.model,
        status: result.status,
        imageTextInfo,
      });

      console.log(`  ${result.altText}`);
      console.log(`  (${result.altText.length} chars${result.status !== 'ok' ? `, ${result.status}` : ''})`);
      if (options.imageTextInfo && imageTextInfo) {
        console.log(`  [Image Text Info]: ${imageTextInfo}`);
      }
    } catch (error) {
      console.error(`  ERROR: ${error.message}`);
      rows.push({
        productName: image.productName,
        cell: image.cell,
        filename: image.filename,
        relativePath: image.relativePath,
        absolutePath: image.absolutePath,
        altText: '',
        characterCount: 0,
        model: options.model,
        status: `error: ${error.message}`,
        imageTextInfo,
      });
    }
  }

  if (options.excel) {
    await exportToExcel(rows, outputFile, options.imageTextInfo);
  } else {
    const headers = [
      'product_name',
      'cell',
      'filename',
      'relative_path',
      'alt_text',
    ];
    if (options.imageTextInfo) {
      headers.push('image_text_info');
    }
    headers.push('character_count', 'model', 'status');

    const csvLines = [headers.join('\t')];
    for (const row of rows) {
      const values = [
        row.productName,
        row.cell,
        row.filename,
        row.relativePath,
        row.altText,
      ];
      if (options.imageTextInfo) {
        values.push(row.imageTextInfo);
      }
      values.push(row.characterCount, row.model, row.status);
      csvLines.push(values.map(csvEscape).join('\t'));
    }

    await fs.mkdir(path.dirname(outputFile), { recursive: true });
    await fs.writeFile(outputFile, `${csvLines.join('\n')}\n`, 'utf8');

    const errors = rows.filter((row) => row.status.startsWith('error:')).length;
    const trimmed = rows.filter((row) => row.status === 'trimmed').length;

    console.log(`\nWrote ${rows.length} row(s) to ${outputFile}`);
    if (trimmed) console.log(`Warning: ${trimmed} result(s) required final hard trimming.`);
    if (errors) console.log(`Warning: ${errors} image(s) failed; see the status column.`);
    return;
  }
}

main().catch((error) => {
  console.error(`Error: ${error.message}`);
  process.exitCode = 1;
});