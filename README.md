# Envision

A local Node.js command-line tool that generates accessibility alt text for product images using a vision model running in Ollama.

It is designed to consume the directory structure produced by an Excel image extractor where each worksheet becomes a product directory and each image is named for its Excel cell.

---

## Installing Ollama and the Required Model

This tool requires [Ollama](https://ollama.com/) and a vision-capable model (`qwen3-vl:4b-instruct` by default).  
Follow the instructions for your operating system below.

### macOS

#### Using Homebrew (Recommended)
    ```bash
    brew install ollama
    ```
After installation, start the Ollama service:
    ```bash
    ollama serve
    ```

#### Using MacPorts
    ```bash
    sudo port install ollama
    ```
Then start the Ollama service:
    ```bash
    ollama serve
    ```

#### Download the Model
Once Ollama is running, pull the required model:
    ```bash
    ollama pull qwen3-vl:4b-instruct
    ```

### Linux

1. Download and run the official Ollama installer:
       ```bash
       curl -fsSL https://ollama.com/install.sh | sh
       ```
2. Start the Ollama service:
       ```bash
       ollama serve
       ```
3. Pull the required model:
       ```bash
       ollama pull qwen3-vl:4b-instruct
       ```


### Windows

1. Download the [Ollama Windows installer](https://ollama.com/download) and run it.
2. After installation, launch the Ollama app (it will start the service automatically).
3. Open Command Prompt or PowerShell and pull the required model:
       ```powershell
       ollama pull qwen3-vl:4b-instruct
       ```

---

**Note:**  
- The default model is `qwen3-vl:4b-instruct`. You can use a different model by specifying `--model &lt;name&gt;` when running the CLI.
- Make sure the Ollama service is running before using this tool.

---

## Input Layout

    ```text
    excel-images/
    ├── Lipstick - Rouge Signature/
    │   ├── A1.jpg
    │   ├── A2.jpg
    │   └── A3.png
    ├── Foundation - Warm Beige/
    │   ├── B2.jpg
    │   └── B3.jpg
    └── context.md                 # optional global context hint
    ```

The immediate parent directory is treated as the product name. `A1.jpg` also produces `A1` in the CSV's `cell` column.

---

## Requirements

- Node.js 20 or later
- Ollama installed and running locally
- A vision-capable Ollama model. The default is `qwen3-vl:4b-instruct`.
- For Excel export: `exceljs` and (optionally) `sharp` for image conversion.

Install dependencies:

    ```bash
    npm install exceljs@^4.3.0 sharp@^0.33.4
    ```

---

## Basic Usage

    ```bash
    node src/envision.js /path/to/excel-images
    ```

The default output is:

    ```text
    /path/to/excel-images/alt-text.csv
    ```

You can also install the CLI from this directory:

    ```bash
    npm link
    envision /path/to/excel-images
    ```

---

## CLI Options

    ```text
    envision &lt;image-root&gt; [options]
    
    Options:
      -o, --output &lt;file&gt;        Output file path (default: &lt;image-root&gt;/alt-text.csv or .xlsx)
      -m, --model &lt;name&gt;         Ollama model (default: qwen3-vl:4b-instruct)
          --max-chars &lt;number&gt;   Maximum alt-text length (default: 1000)
          --context &lt;text&gt;       Global context hint supplied on the command line
          --context-file &lt;file&gt;  Global context hint file
          --prompt &lt;text&gt;        Override the default prompt template
          --prompt-file &lt;file&gt;   Read the prompt template from a file
          --temperature &lt;n&gt;      Ollama temperature (default: 0.2)
          --host &lt;url&gt;           Ollama API URL (default: http://127.0.0.1:11434)
          --retries &lt;number&gt;     Retry attempts if output is too long (default: 2)
          --dry-run              Discover images and show configuration without Ollama calls
          --verbose              Print the rendered prompt for each image
          --excel                Export as Excel (.xlsx) instead of CSV
      -h, --help                 Show this help
    ```

---

## CSV Output

    ```csv
    product_name,cell,filename,relative_path,alt_text,character_count,model,status
    ```

Example:

    ```csv
    Lipstick - Rouge Signature,A1,A1.jpg,Lipstick - Rouge Signature/A1.jpg,"Rouge Signature liquid lipstick in a clear rectangular tube with a black cap and deep red shade visible inside.",112,qwen3-vl:4b-instruct,ok
    ```

The CSV should be treated as a review queue, not as automatically approved accessibility content.

---

## Excel Output

- Use `--excel` to export to Excel (`.xlsx`).
- Each product directory becomes a separate worksheet/tab.
- Each image file is a row with:
  - Filename
  - Embedded image (in the next cell, resizable with the cell, aspect ratio preserved)
  - Alt text (word-wrapped, column auto-sized for readability)
  - Character count, model, and status

Example command:

    ```bash
    node src/envision.js ./excel-images --excel --output ./results/alt-text.xlsx
    ```

**Excel Features:**
- Images are embedded at their original size and will scale with the cell (maintaining aspect ratio).
- Alt text column is set to wrap and auto-size so all text is visible without scrolling.
- If `sharp` is installed, non-PNG/JPEG images are converted to PNG for compatibility.

---

## Character Limit

The default maximum is 1000 characters:

    ```bash
    node src/envision.js ./excel-images --max-chars 150
    ```

The limit is enforced in three stages:

1. The original prompt tells the vision model the limit.
2. If the result is too long, the tool asks the model to shorten it without adding facts.
3. If it is still too long after the configured retries, the tool trims it at a word boundary and sets `status` to `trimmed`.

Set the retry count with:

    ```bash
    node src/envision.js ./excel-images --max-chars 150 --retries 3
    ```

---

## Context Hints

Context can supply authoritative catalogue information that the image alone cannot reliably establish.

### Root `context.md`

Put `context.md` directly in the input root:

    ```text
    excel-images/context.md
    ```

It will be supplied to every image.

### Product-specific `context.md`

A product directory can contain its own context:

    ```text
    excel-images/
    └── Rouge Signature 129/
        ├── A1.jpg
        ├── A2.jpg
        └── context.md
    ```

This is appended only when generating alt text for that product.

### Context file option

    ```bash
    node src/envision.js ./excel-images \
      --context-file ./catalogue-context.md
    ```

### Inline context

    ```bash
    node src/envision.js ./excel-images \
      --context "These are cosmetics product images. Exact shade names in product folder names are authoritative."
    ```

If several context sources are present, they are appended in this order:

1. root `context.md`
2. `--context-file`
3. `--context`
4. product-specific `context.md`

---

## Custom Prompts

The built-in prompt is accessibility-oriented, but it can be completely replaced.

From a file:

    ```bash
    node src/envision.js ./excel-images \
      --prompt-file ./prompt.example.txt
    ```

Or directly:

    ```bash
    node src/envision.js ./excel-images \
      --prompt "Write concise alt text for {{productName}}. Context: {{context}} Maximum {{maxChars}} characters. Return only alt text."
    ```

Available placeholders:

- `{{productName}}`
- `{{filename}}`
- `{{maxChars}}`
- `{{context}}`

---

## Choose a Different Model

For example, use the larger 8B instruct model:

    ```bash
    ollama pull qwen3-vl:8b-instruct
    node src/envision.js ./excel-images --model qwen3-vl:8b-instruct
    ```

---

## Ollama Host

The default Ollama API is `http://127.0.0.1:11434`. Override it if necessary:

    ```bash
    node src/envision.js ./excel-images --host http://127.0.0.1:11434
    ```

---

## Choose Output Location

    ```bash
    node src/envision.js ./excel-images \
      --output ./results/alt-text.csv
    ```

Or for Excel:

    ```bash
    node src/envision.js ./excel-images \
      --excel --output ./results/alt-text.xlsx
    ```

---

## Dry Run

Check which images will be processed without contacting Ollama:

    ```bash
    node src/envision.js ./excel-images --dry-run
    ```

---

## Verbose Mode

Print the fully rendered prompt for each image:

    ```bash
    node src/envision.js ./excel-images --verbose
    ```

---

## Supported Images

The scanner currently accepts:

- JPG/JPEG
- PNG
- WebP
- GIF
- BMP
- TIFF

Images must be immediately inside their product directory. Files in the root directory are not treated as products.

---

## Extraction of images from an Excel folder.

Included in this project is `extract-excel-images.js`, which can extract images placed in an Excel file.  This has proven useful in the past when 

## Why the Product Directory Name is Supplied to the Model

A vision model may recognize packaging or visible branding, but it should not be relied on to infer an exact catalogue product name. The directory name is therefore supplied as authoritative metadata while the model is instructed to describe only visual details it can actually see.

---

## Accessibility Note

Generated alt text should be human-reviewed. The correct alternative text depends on the purpose and context in which an image is used, not only on image recognition. The `status`, character count, product name, cell, and filename columns make the CSV and Excel outputs useful as a review workflow.

---
