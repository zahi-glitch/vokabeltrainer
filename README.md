# Vokabeltrainer — Die neue Linie 1 B1

A small desktop flashcard app (Python + `tkinter`) for practising German **B1**
vocabulary from the *Die neue Linie 1 B1* coursebook. The vocabulary lives in
three Excel workbooks — verbs, nouns and adjectives — with **one sheet per
chapter** (Kapitel), so it is easy to read, edit and extend in Excel.

## Try it online — no install

A web version runs in any browser (phone, tablet or computer), with the same
vocabulary as the Excel files:

**▶ https://claude.ai/artifact/M62hhMXyYrySZARfCF3KFN**

It is a static snapshot built from the `.xlsx` files (currently chapters K1–K4);
the link stays the same and the content refreshes each time the page is rebuilt
from updated workbooks.

## Features

- Flashcards filtered by **chapter**, **word type** (Verben / Nomen / Adjektive) and **theme**
- **English → German** prompts; reveal the solution with all forms, an example and notes
- **"No repeat"** mode: every card is shown once before any repeats, with a progress counter
- Colour-coded by chapter; articles shown as part of each noun
- Reads vocabulary straight from the `.xlsx` files — no database, just Excel

## Requirements

- Python 3.9+
- [`openpyxl`](https://openpyxl.readthedocs.io/) (see `requirements.txt`)
- `tkinter` — bundled with the python.org build of Python. On macOS with
  Homebrew Python, install it once with `brew install python-tk`.

## Install & run

```bash
git clone https://github.com/zahi-glitch/vokabeltrainer.git
cd vokabeltrainer

python3 -m venv env
source env/bin/activate          # Windows: env\Scripts\activate
pip install -r requirements.txt

python trainer_gui.py
```

By default the app loads `B1_Verben.xlsx`, `B1_Nomen.xlsx` and
`B1_Adjektive.xlsx` from the project folder. You can also pass files
explicitly — `python trainer_gui.py B1_Verben.xlsx B1_Nomen.xlsx` — or pick
them with the **Dateien…** button.

**Keyboard:** `Space` / `Enter` = next · `→` = skip · `q` / `Esc` = quit.

## Data format

Each workbook has one sheet per chapter, named like `K01 Neue Nachbarn`.
Columns are detected automatically from the header row, for example:

| Type       | Columns |
|------------|---------|
| Verben     | Thema · Infinitiv · English · Präteritum · Partizip II · Verwendung |
| Nomen      | Thema · Artikel · Nomen · Plural · English · Beispiel · Hinweis |
| Adjektive  | Thema · Adjektiv · English · Komparativ · Superlativ · Gegenteil · Beispiel · Hinweis |

To add a chapter, add a new sheet (`K05 …`); to add a word, add a row.

## License

Released under the [MIT License](LICENSE).

---

A personal language-learning project. Vocabulary selections are inspired by the
chapter themes of *Die neue Linie 1 B1* (Ernst Klett Sprachen) but are written
independently; this repository is not affiliated with or endorsed by the publisher.
