import os
import string
import argparse

import docx
import matplotlib.pyplot as plt
from wordcloud import WordCloud

import spacy


# -----------------------------
# HELP
# -----------------------------
def show_help():
    print("""
Word Mesh Generator (Stable NLP version)

USAGE:
    python create_word_mesh_v4.py --file mesh.docx --format both --lang both

OPTIONS:
    --file      Input .docx file
    --output    Output name (default: wordcloud)
    --lang      en | nl | both (default: both)
    --format    png | pdf | both (default: both)

EXAMPLES:
    python create_word_mesh_v4.py --file mesh.docx
    python create_word_mesh_v4.py --file mesh.docx --format pdf
""")


# -----------------------------
# DOCX READER
# -----------------------------
def extract_text_from_docx(path):
    doc = docx.Document(path)
    return " ".join(p.text for p in doc.paragraphs)


# -----------------------------
# LOAD NLP MODEL (SPA CY)
# -----------------------------
def load_model(lang):
    if lang == "nl":
        return spacy.load("nl_core_news_sm")
    return spacy.load("en_core_web_sm")


# -----------------------------
# CLEAN + LEMMATIZE
# -----------------------------
def clean_text(text, lang="both"):
    extra_stopwords = {
        "figuur", "tabel", "hoofdstuk", "bijlage", "daarnaast",
        "verder", "gemaakt", "gebruik", "binnen", "volgende",
        "komen", "gaan", "kunnen", "zullen", "worden", "zijn",
        "hebben", "heeft", "is", "was", "een", "de", "het",
        "van", "te", "in", "op", "en", "of", "aan", "met"
    }

    stop_words = set(extra_stopwords)

    if lang in ("en", "both"):
        stop_words.update(spacy.load("en_core_web_sm").Defaults.stop_words)

    if lang in ("nl", "both"):
        stop_words.update(spacy.load("nl_core_news_sm").Defaults.stop_words)

    nlp = spacy.load("nl_core_news_sm")  # good general baseline for NL-heavy docs

    doc = nlp(text.lower())

    words = []

    for token in doc:
        lemma = token.lemma_.strip().lower()

        if token.is_alpha and lemma not in stop_words:
            words.append(lemma)

    return " ".join(words)

# -----------------------------
# WORDCLOUD
# -----------------------------
def generate_wordcloud(text, output="wordcloud", fmt="both"):
    wc = WordCloud(
        width=1000,
        height=500,
        background_color="white",
        colormap="viridis"
    ).generate(text)

    fig = plt.figure(figsize=(10, 5))
    plt.imshow(wc, interpolation="bilinear")
    plt.axis("off")

    if fmt in ("png", "both"):
        plt.savefig(f"{output}.png", dpi=300, bbox_inches="tight")

    if fmt in ("pdf", "both"):
        plt.savefig(f"{output}.pdf", bbox_inches="tight")

    # IMPORTANT: prevent blocking
    plt.close(fig)


# -----------------------------
# MAIN
# -----------------------------
def main(file, output, lang, fmt):
    if not os.path.exists(file):
        print(f"File not found: {file}")
        return

    text = extract_text_from_docx(file)
    cleaned = clean_text(text, lang)
    generate_wordcloud(cleaned, output, fmt)


# -----------------------------
# CLI
# -----------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(add_help=False)

    parser.add_argument("--file")
    parser.add_argument("--output", default="wordcloud")
    parser.add_argument("--lang", default="both")
    parser.add_argument("--format", default="both")
    parser.add_argument("--help", action="store_true")

    args = parser.parse_args()

    if args.help or not args.file:
        show_help()
    else:
        main(args.file, args.output, args.lang, args.format)