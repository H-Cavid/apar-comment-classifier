# Multilingual Customer Comment Classifier (AZ / RU / EN)

A transparent, offline Python pipeline that sorts free-text customer comments from a bike- and scooter-sharing service into **39 complaint categories** – brakes, battery, app login, pricing, parking, positive feedback and more.

Built during my work as a Data Analyst for [Apar](https://apar.az/en), Azerbaijan's bike-sharing platform, to turn **23,933 unstructured comments** into a structured overview the operations team could act on.

> The real customer data is confidential and is **not** included in this repository. `examples/sample_comments.xlsx` contains invented comments for demonstration only.

---

## The problem

Riders leave short comments after a trip. They are:

- **multilingual** – Azerbaijani, Russian and English, often mixed in one sentence
- **informal** – Azerbaijani is frequently typed without diacritics (`islemir` instead of `işləmir`), with typos and slang
- **multi-label** – one comment can mention a broken brake *and* a high price

The operations team needed to know *which* problems occur, *how often*, and *where to act first* – without reading 24,000 comments by hand.

## Approach

**Rule-based multilingual keyword classification**, iteratively refined against human-labeled data.

1. **Normalization** – lowercase text and map Azerbaijani-specific letters to ASCII (`ə→e`, `ı→i`, `ş→s` …), so `işləmir` and `islemir` match the same rule.
2. **Regex dictionaries** – each of the 39 categories has hand-built patterns in Azerbaijani, Russian and English (synonyms, inflections, common misspellings).
3. **Multi-label matching** – every category whose patterns match is assigned; comments with no match fall into `O01 – Unclassified`.
4. **Conflict rules** – e.g. the generic "technical fault" bucket is dropped when a specific technical category already explains the comment; strong imperative requests ("please add …") are tagged as suggestions, not as current shortages.

### Why rules instead of an LLM?

I also experimented with **local LLMs (Ollama)**, **Hugging Face zero-shot classification** and **unsupervised topic discovery**. For this use case I chose the rule-based pipeline because it:

- runs **fully offline at zero API cost** and processes 24,000+ comments in seconds
- is **fully transparent** – every label can be traced back to the exact word that triggered it, so results are easy to audit
- supports multi-label output out of the box

The exploratory work was useful for discovering categories, but it is not part of this repository.

## Evaluation

I hand-labeled **500 comments** and compared them with the pipeline's output (`eval_manual_labels.py`). Error analysis on the mismatches revealed six concrete bugs; fixing them improved the results:

| Metric | Before | After |
|---|---|---|
| Micro-F1 | 63.7% | **78.1%** |
| Exact match (all labels correct) | 50.7% | **72.9%** |
| Jaccard similarity (multi-label) | 60.8% | **77.7%** |

Final micro precision / recall: **78.0% / 78.3%**.

### Bugs found through error analysis

1. The generic "technical fault" category fired even when a specific category (e.g. brakes) already applied.
2. Russian **«дорого»** (*expensive*) was matched as **«дорога»** (*road*) – a Latin/Cyrillic "o" mix-up.
3. The Azerbaijani word *"sep"* actually meant *chain*, not *basket*.
4. *"təmir"* (*repair*) was too broad for the maintenance category.
5. The alignment category was triggered too easily by *"crooked handlebar"* alone.
6. Suggestions and parking shortages were being confused.

The same frequency analysis on unclassified comments led to **four new categories** (shock absorber, sudden stop during a ride, display, handlebar grip) – real recurring problems missing from the original taxonomy.

## Limitations

- Keyword matching does not understand irony, paraphrases or complex sentences.
- About 28% of the unclassified comments are generic sentiment without detail ("very bad", "great") – they cannot be mapped to a specific problem without more context.
- Human labelers themselves disagreed on some borderline categories, so 100% accuracy is not a realistic target for rules.
- The four categories added after validation (T24–T27) have not been separately evaluated yet.
- A natural next step would be an LLM-based classifier for the remaining hard cases, with the rule-based output as a baseline.

---

## Usage

```bash
pip install -r requirements.txt
python run_classification.py examples/sample_comments.xlsx
```

The input file needs the comments in the **first column of the first sheet** (row 1 is a header).
The output `*_kateqoriyalasdirilmis.xlsx` contains two sheets:

- `Kategoriyalar` – every comment with one 0/1 column per category code
- `Xulase` – a summary with the count and share of each category

### Evaluate against your own labels

```bash
python eval_manual_labels.py manual_validation.xlsx   # human labels vs. pipeline
python eval_accuracy.py sample_to_check.xlsx          # quick accuracy check on a filled-in sample
```

Both scripts report exact match, Jaccard similarity and per-category precision / recall / F1.

### Run the tests

```bash
pip install -r requirements-dev.txt
pytest
```

## Category codes

| Group | Codes |
|---|---|
| Technical faults | T01–T27 (brakes, pedal, chain, seat, handlebar, tyres, lock, battery, lights …) |
| App & digital | A01 GPS, A02 app / login / QR / session, A03 Bluetooth |
| Payment & price | P01 price complaints, P02 payment errors |
| Infrastructure | I01 parking, I02 service area, I03 vehicle availability, I04 bike lanes |
| Suggestions | S01 new features / equipment |
| Positive | M01 satisfaction |
| Other | O01 unclassified |

Category names and code comments are in Azerbaijani, as the tool was built for an Azerbaijani team.

## Tech

Python 3 · regular expressions · `openpyxl` · pytest
