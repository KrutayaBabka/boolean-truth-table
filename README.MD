# Boolean Truth Table Generator

This project generates truth tables from Boolean expressions and exports them to a DOCX file.

## Features
- Supports Boolean expressions with operators:
  - **NOT**: ¬
  - **AND**: * or ^
  - **OR**: + or ∨
- Variable names can include letters, digits, and underscores (e.g., F, x1, x_2)
- Outputs truth table in DOCX format with **Times New Roman**, size 12

## Project Structure
```

boolean-truth-table/
├── src/
│   └── truth_table_generator.py
├── requirements.txt
└── README.md

````

## Installation

1. **Clone the repository:**
```bash
git clone https://github.com/KrutayaBabka/boolean-truth-table.git
cd boolean-truth-table
````

2. **Create a virtual environment:**

```bash
python -m venv venv
```

3. **Activate the environment:**

```bash
# Windows
venv\Scripts\activate

# Linux / Mac
source venv/bin/activate
```

4. **Install dependencies:**

```bash
pip install -r requirements.txt
```

## Usage

Run the script:

```bash
python src/truth_table_generator.py
```

Then enter your Boolean expression when prompted. The truth table will be saved as `truth_table.docx` in the project folder.

### Examples

```
Enter Boolean expression: ¬(F*¬G)+¬(C+D)+E+F*G*¬H
```

### Notes

* You can also enter expressions copied from Word formulas, e.g.:

```
(F∙G̅)̅+(C+D)̅+E+F∙G∙H̅
```

* The script will parse and convert them automatically.