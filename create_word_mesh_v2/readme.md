# 🧠 Word Mesh Generator (Multilingual NLP Word Cloud Tool)

A Python tool that generates **high-quality word clouds** from `.docx` files using **modern NLP preprocessing** (lemmatization + multilingual stopword filtering).

Supports:
- 🇳🇱 Dutch text
- 🇬🇧 English text
- 📄 DOCX input
- 🖼️ PNG export
- 📄 PDF export
- 🧠 Lemmatization (spaCy-based NLP)

---

# 🚀 Features

✔ Extract text from Word documents (`.docx`)  
✔ Clean text using NLP (lemmatization + stopword removal)  
✔ Supports Dutch + English  
✔ Generates word clouds  
✔ Exports to:
- PNG (high resolution)
- PDF (vector format)

✔ CLI-based tool (easy automation)  
✔ Stable NLP pipeline (no broken NLTK dependencies)

---

# 📦 Installation

## 1. Install dependencies

```bash
pip install python-docx matplotlib wordcloud spacy
```

## 2. Download spaCy language models

```bash
python -m spacy download en_core_web_sm
python -m spacy download nl_core_news_sm
```

---

# 📂 Project Structure

```
create_word_mesh_v4.py   # main script
README.md                # this file
demo_english.docx       # example input
demo_dutch.docx         # example input
```

---

# ⚙️ Usage

## Basic usage

```bash
python create_word_mesh_v2.py --file demo_english.docx
```

## Specify output format

```bash
python create_word_mesh_v2.py --file demo_dutch.docx --format png
python create_word_mesh_v2.py --file demo_dutch.docx --format pdf
python create_word_mesh_v2.py --file demo_dutch.docx --format both
```

## Change output name

```bash
python create_word_mesh_v2.py --file demo_english.docx --output result
```

---

# 🌍 Language support

| Option | Description |
|--------|-------------|
| `en`   | English only |
| `nl`   | Dutch only |
| `both` | Mixed language processing |

---

# 🧪 Demo Files

Two example files are included:

- 📄 `demo_english.docx` → English NLP test text
- 📄 `demo_dutch.docx` → Dutch NLP test text

These demonstrate:
- Stopword removal
- Lemmatization
- Meaningful keyword extraction

---

# 🖼️ Output Example

The script generates:

- `wordcloud.png` → preview image
- `wordcloud.pdf` → print-ready vector file

---

# 🧠 NLP Pipeline

The system uses spaCy:

1. Tokenization
2. Lemmatization
3. Stopword filtering (EN + NL)
4. Custom domain stopwords
5. Word frequency aggregation
6. Word cloud generation

---

# 🔥 Example Command

```bash
python create_word_mesh_v4.py --file demo_english.docx --format both --output english_cloud
```

Output:
```
english_cloud.png
english_cloud.pdf
```

---

# 📌 Use Cases

- Academic text analysis
- Thesis keyword extraction
- Report summarization
- Content analysis
- Language comparison (NL vs EN)

---

# 📄 License

Free to use and modify for personal and commercial projects.

---

# 🚀 Future improvements

Planned enhancements:

- TF-IDF weighted word clouds
- Interactive HTML visualization
- Batch processing (folders)
- Keyword clustering
- GUI version

