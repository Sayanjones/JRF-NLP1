# English-Hindi NLP Assignment

**Author:** Sayan H. Mandal  
**Tasks:** Dataset Processing (Assignment 1) + LLM Translation (Assignment 2)

---

## Repository Structure

```
├── assignment1.py                # Assignment 1: Dataset processing code
├── assignment2.py                # Assignment 2: LLM translation code
├── Dataset_English_Hindi.csv     # Input dataset (download from Kaggle)
├── assignment1_output.xlsx       # Final Excel output for Assignment 1
├── assignment2_translations.xlsx # Final Excel output for Assignment 2
├── assignment2_metrics.txt       # BLEU, CHRF, TER scores
└── README.md                     # This file
```

---

## Requirements

Install all dependencies with:

```bash
pip install pandas openpyxl sacrebleu anthropic
```

---

## Assignment 1: English-Hindi Dataset Processing

### What it does
- Loads the English-Hindi parallel CSV dataset
- Extracts English and Hindi sentences into two columns
- Computes word counts for both languages
- Filters sentences where both word counts are in range **[5, 50]**
- Filters pairs where word count difference (English - Hindi) is in **[-10, +10]**
- Saves the cleaned dataset to Excel with 5 columns

### How to run

1. Place `Dataset_English_Hindi.csv` in the same folder as `assignment1.py`
2. Run:

```bash
python assignment1.py
```

3. Output file `assignment1_output.xlsx` will be generated with these columns:
   - English Sentences
   - Hindi Sentences
   - Word Count (English)
   - Word Count (Hindi)
   - Difference between Word Count (English) and Word Count (Hindi)

---

## Assignment 2: Translation with LLM

### What it does
- Selects 100 English sentences from the Assignment 1 output
- Translates them to Hindi using **Claude claude-sonnet-4-6** (Anthropic LLM)
- Computes **BLEU**, **CHRF**, and **TER** evaluation scores
- Saves metrics to a `.txt` file
- Saves translations to Excel (Column A: English, Column B: Hindi)

### How to run

1. Make sure `assignment1_output.xlsx` exists (run Assignment 1 first)
2. Get your Anthropic API key from: https://console.anthropic.com
3. Set your API key:

```bash
# On Linux/Mac
export ANTHROPIC_API_KEY="api_key_here"

# On Windows (Command Prompt)
set ANTHROPIC_API_KEY=api_key_here
```

4. Run:

```bash
python assignment2.py
```

5. Output files generated:
   - `assignment2_translations.xlsx` - Column A: English, Column B: Hindi translation
   - `assignment2_metrics.txt` - BLEU, CHRF, TER scores

---

## Evaluation Metrics Explained

| Metric | Range | Better When |
|--------|-------|-------------|
| BLEU   | 0–100 | Higher      |
| CHRF   | 0–100 | Higher      |
| TER    | 0–∞   | Lower       |

- **BLEU**: Measures n-gram precision between predicted and reference translations
- **CHRF**: Character-level F-score, more sensitive to morphological differences
- **TER**: Translation Edit Rate - number of edits needed to match the reference

---

## Notes

- Run Assignment 1 **before** Assignment 2
- Do **not** modify the output Excel files after submission - any changes post-assessment are treated as violations per the assignment instructions
