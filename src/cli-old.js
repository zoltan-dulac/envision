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

function shortenAtWordBoundary(text, maxChars) {
  if (text.length <= maxChars) return text;
  if (maxChars <= 1) return text.slice(0, maxChars);

  const slice = text.slice(0, maxChars + 1);
  const lastSpace = slice.lastIndexOf(' ');
  const cutoff = lastSpace >= Math.floor(maxChars * 0.6) ? lastSpace : maxChars;

  return text.slice(0, cutoff).replace(/[\s,;:–—-]+$/u, '').trim();
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
    altText = shortenAtWordBoundary(altText, maxChars);
    status = 'trimmed';
  }

  return { altText, status, retryCount };
}

// --- Excel Export Function ---
async function exportToExcel(images, outputFile) {
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
    sheet.columns = [
      { header: 'Filename', key: 'filename', width: 30 },
      { header: 'Image', key: 'image', width: 20 },
      { header: 'Alt Text', key: 'altText', width: 60, style: { alignment: { wrapText: true, vertical: 'top' } } },
      { header: 'Character Count', key: 'characterCount', width: 18 },
      { header: 'Model', key: 'model', width: 18 },
      { header: 'Status', key: 'status', width: 18 },
    ];
    sheet.getRow(1).font = { bold: true };

    for (let i = 0; i < imgs.length; i++) {
      const { filename, absolutePath, altText, characterCount, model, status } = imgs[i];
      const rowIdx = i + 2; // 1-based, row 1 is header

      // Insert filename and alt text
      sheet.getCell(`A${rowIdx}`).value = filename;
      sheet.getCell(`C${rowIdx}`).value = altText;
      sheet.getCell(`D${rowIdx}`).value = characterCount;
      sheet.getCell(`E${rowIdx}`).value = model;
      sheet.getCell(`F${rowIdx}`).value = status;

      // Insert image if exists
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
        sheet.getRow(rowIdx).height = 60;
      } catch (err) {
        // If image can't be read, leave cell blank but log error
        console.error(`Warning: Could not embed image for ${filename}: ${err.message}`);
      }
    }

    // Set alt text column to wrap and auto-fit
    sheet.getColumn('C').alignment = { wrapText: true, vertical: 'top' };
    // Optionally, auto-size alt text column to fit content
    let maxAltLen = Math.max(...imgs.map(img => (img.altText || '').length), 10);
    sheet.getColumn('C').width = Math.min(Math.max(30, maxAltLen / 2), 100);
  }

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
      });

      console.log(`  ${result.altText}`);
      console.log(`  (${result.altText.length} chars${result.status !== 'ok' ? `, ${result.status}` : ''})`);
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
      });
    }
  }

  if (options.excel) {
    await exportToExcel(rows, outputFile);
  } else {
  const headers = [
    'product_name',
    'cell',
    'filename',
    'relative_path',
    'alt_text',
    'character_count',
    'model',
    'status',
  ];

  const csvLines = [headers.join(',')];
  for (const row of rows) {
    csvLines.push([
      row.productName,
      row.cell,
      row.filename,
      row.relativePath,
      row.altText,
      row.characterCount,
      row.model,
      row.status,
    ].map(csvEscape).join(','));
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