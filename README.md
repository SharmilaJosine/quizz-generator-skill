# Interactive Quiz Generator Skill

A Python command-line utility that extracts content from **PDF**, **Excel (XLSX)**, and **Word (DOCX)** documents and generates a visually stunning, responsive, self-contained interactive **HTML/CSS/JS** quiz.

The generated quiz works completely offline, is portable, and features two built-in theme layouts:
- 🧸 **Playful Kids Theme (Ages 10-14, Default)**: Bouncy visual scales, soft card shadows, bubbly typography, large tap buttons, friendly emojis, and engaging layouts.
- 🎨 **Classic Dark Glassmorphism Theme**: Radial gradient glow, frosted-glass borders, glowing highlights, and premium modern developer interface styles.
- 💡 **Interactive Question Types**: Multiple choice (cards), Yes/No toggle boxes, and Drag-and-Drop matching (with full click-to-match mobile fallback).
- 🔊 **Synthesized Sound Effects**: Audio responses for correct/incorrect answers generated synthetically using the browser's built-in Web Audio API (no external audio assets required).
- 🏆 **Gamified Features**: Circular animated score ring, live progress bar, active timer, and an HTML5 Canvas-based confetti celebrate effect.
- 📝 **Detailed Explanations & Review List**: Full question review list showing correct answers and detailed contexts of why they are right.

---

## Prerequisites

- **Python 3.10+**
- A **Gemini API Key** (required to use the LLM parsing to synthesize intelligent questions from your documents)

---

## Setup & Installation

1. Open your terminal in the directory:
   ```bash
   cd C:/Users/IDK/.gemini/antigravity/scratch/quiz-generator-skill
   ```

2. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

Run the `generate_quiz.py` script and pass your input file.

### Environment Variable (Recommended)
Set your Gemini API key in your session environment, then run the script:
```powershell
# In PowerShell (Windows):
$env:GEMINI_API_KEY="your-api-key-here"

# This defaults to the Kids theme:
python generate_quiz.py sample_doc.pdf
```

### Specifying Theme Style
Choose between `kids` (default) and `default` (dark glassmorphism):
```powershell
# Bubbly Kids Theme (Ages 10-14):
python generate_quiz.py sample_doc.pdf --theme kids

# Dark Glassmorphism:
python generate_quiz.py sample_doc.pdf --theme default
```

### Direct CLI Argument
You can also supply your key directly using the `-k` or `--api-key` argument:
```powershell
python generate_quiz.py sample_doc.docx --api-key "your-api-key-here" --theme default
```

### Interactive Prompt
If the key is not in your environment variables or CLI arguments, the script will securely prompt you to paste it in:
```powershell
python generate_quiz.py sample_doc.xlsx
[!] GEMINI_API_KEY environment variable or --api-key argument not provided.
[?] Enter your Gemini API Key: <PASTE_API_KEY_HERE>
```

---

## CLI Options

```
usage: generate_quiz.py [-h] [-o OUTPUT] [-k API_KEY] [-n NUM_QUESTIONS] [-t TEMPLATE] [--theme {default,kids}] input_file

positional arguments:
  input_file            Path to the input document (PDF, Word, or Excel file).

options:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        Path to write the output HTML quiz file. Defaults to input name + .html
  -k API_KEY, --api-key API_KEY
                        Gemini API Key. Can also be set via GEMINI_API_KEY environment variable.
  -n NUM_QUESTIONS, --num-questions NUM_QUESTIONS
                        Number of questions to generate. Default is 10.
  -t TEMPLATE, --template TEMPLATE
                        Path to the quiz template file. If not set, loaded based on --theme.
  --theme {default,kids}
                        Select the theme template style: 'kids' (vibrant, playful for ages 10-14, default) or 'default' (dark glassmorphism).
```

---

## 🛠️ Customizing or Adding Templates

The CLI is structured to make custom template creation simple:
1. Create a new HTML template file in the script directory named:
   `quiz_template_<your_theme_name>.html`
2. Incorporate the standard data block:
   ```html
   <script id="quiz-data" type="application/json">
     /* {{QUIZ_DATA_PLACEHOLDER}} */
   </script>
   ```
3. Run the generator script with your theme name:
   ```powershell
   python generate_quiz.py sample_doc.pdf --theme <your_theme_name>
   ```
   *The script will automatically discover and load your template!*

---

## File Support & Details

| File Type | Library | What it Extracts |
| :--- | :--- | :--- |
| **PDF (`.pdf`)** | `PyMuPDF` | Full text formatting from all pages. |
| **Word (`.docx`)** | `python-docx` | Paragraph texts, headings, and cell tables. |
| **Excel (`.xlsx`)** | `openpyxl` | Grid row cells, and labels across all worksheets. |
