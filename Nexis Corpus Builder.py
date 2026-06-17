from pathlib import Path
from docx import Document
from xml.etree.ElementTree import Element, SubElement, ElementTree, indent
import re
from datetime import datetime




input_folder = input("Paste the folder containing the DOCX files: ").strip().strip('"')
output_parent = input("Paste the folder where the corpus folder should be created: ").strip().strip('"')
corpus_name = input("Name this corpus: ").strip()

if not corpus_name:
    corpus_name = "CORPUS_XML"

corpus_name = re.sub(r'[<>:"/\\|?*]+', "_", corpus_name)

PASTA_ENTRADA = Path(input_folder)
PASTA_SAIDA = Path(output_parent) / corpus_name

PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

SECTION_MAP = {
    "sports": "Sport",
    "sport": "Sport",
    "business": "Business",
    "business day": "Business",
    "world": "World",
    "middleeast": "World",
    "americas": "World",
    "africa": "World",
    "europe": "World",
    "asia": "World",
    "u.s.": "US",
    "us": "US",
    "national": "US",
    "politics": "Politics",
    "opinion": "Opinion",
    "editorial": "Opinion",
    "arts": "Arts",
    "books": "Books",
    "style": "Style",
    "travel": "Travel",
    "science": "Science",
    "health": "Health",
    "technology": "Technology",
    "tech": "Technology",
    "climate": "Climate",
    "real estate": "RealEstate",
    "food": "Dining",
    "dining": "Dining",
    "nyregion": "NYRegion",
    "ny region": "NYRegion",
    "magazine": "Magazine",
    "movies": "Movies",
}


CATEGORY_FOLDERS = {
    "Sport": "Sports",
    "Business": "Business",
    "World": "World",
    "US": "US",
    "Politics": "Politics",
    "Opinion": "Opinion",
    "Arts": "Arts",
    "Style": "Style",
    "Travel": "Travel",
    "Science": "Science",
    "Health": "Health",
    "Technology": "Technology",
    "Climate": "Climate",
    "RealEstate": "RealEstate",
    "Dining": "Dining",
    "NYRegion": "NYRegion",
    "Magazine": "Magazine",
    "Books": "Books",
    "Movies":"Movies",
    "Other": "Other",
}


BOILERPLATE_STARTS = [
    "the times is committed to",
    "follow the new york times",
    "reporting was contributed by",
]

BOILERPLATE_CONTAINS = [
    "contributed reporting",
    "by getty images",
]


def normalize_text(text):
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")
    text = text.replace("\xa0", " ")

    invisible_chars = [
        "\u200b",
        "\u200c",
        "\u200d",
        "\ufeff",
    ]

    for char in invisible_chars:
        text = text.replace(char, "")

    replacements = {
        "\u2013": "-",
        "\u2212": "-",
        "\u2014": "-",
        "\u2010": "-",
        "-": "-",
        "\u201c": '"',
        "\u201d": '"',
        "\u2033": '"',
        "''": '"',
        "\u2018": "'",
        "\u2019": "'",
        "`": "'",
        "\u00b4": "'",
        "\u00e2\u20ac\u201c": "-",
        "\u00e2\u02c6\u2019": "-",
        "\u00e2\u20ac\u201d": "-",
        "\u00e2\u20ac\u0090": "-",
        "\u00e2\u20ac\u0153": '"',
        "\u00e2\u20ac\u009d": '"',
        "\u00e2\u20ac\u00b3": '"',
        "\u00e2\u20ac\u02dc": "'",
        "\u00e2\u20ac\u2122": "'",
        "\u00c2\u00b4": "'",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"\n{2,}", "\n", text)
    text = re.sub(r"[^\S\n]{2,}", " ", text)

    lines = []
    for line in text.split("\n"):
        line = re.sub(r"^\s+", "", line)
        line = re.sub(r"\s+$", "", line)

        if line:
            lines.append(line)

    return "\n".join(lines).strip()


def get_paragraphs(docx_path):
    doc = Document(docx_path)
    paragraphs = []

    for p in doc.paragraphs:
        text = normalize_text(p.text)

        if text:
            paragraphs.append(text)

    return paragraphs


def get_value_after_colon(text):
    return text.split(":", 1)[1].strip()


def is_boilerplate(paragraph):
    lower = paragraph.lower().strip()

    if lower.startswith("copyright"):
        return True

    if lower.startswith("load-date"):
        return True

    if lower == "end of document":
        return True

    if any(item in lower for item in BOILERPLATE_CONTAINS):
        return True

    if any(lower.startswith(item) for item in BOILERPLATE_STARTS):
        return True

    return False


def clean_body_paragraphs(paragraphs):
    cleaned = []

    for paragraph in paragraphs:
        paragraph = normalize_text(paragraph)

        if not paragraph:
            continue

        if is_boilerplate(paragraph):
            continue

        cleaned.append(paragraph)

    return cleaned


def extract_abstract_and_full_text(body_paragraphs):
    highlight = ""
    body = body_paragraphs

    upper_paragraphs = [p.upper().strip() for p in body_paragraphs]

    if "ABSTRACT" in upper_paragraphs:
        abstract_index = upper_paragraphs.index("ABSTRACT")

        full_text_index = None
        if "FULL TEXT" in upper_paragraphs:
            full_text_index = upper_paragraphs.index("FULL TEXT")

        if full_text_index is not None and full_text_index > abstract_index:
            abstract_paragraphs = body_paragraphs[abstract_index + 1:full_text_index]
            body = body_paragraphs[full_text_index + 1:]
        else:
            abstract_paragraphs = body_paragraphs[abstract_index + 1:]
            body = body_paragraphs[:abstract_index]

        abstract_paragraphs = clean_body_paragraphs(abstract_paragraphs)
        highlight = " ".join(abstract_paragraphs).strip()

    elif "FULL TEXT" in upper_paragraphs:
        full_text_index = upper_paragraphs.index("FULL TEXT")
        body = body_paragraphs[full_text_index + 1:]

    body = clean_body_paragraphs(body)

    return highlight, body


def classify_section(section):
    lower = section.lower().strip()

    for keyword, category in SECTION_MAP.items():
        if keyword in lower:
            return category

    return "Other"


def extract_year(date_text):
    match = re.search(r"\b(19|20)\d{2}\b", date_text)

    if match:
        return match.group(0)

    return "Unknown"


def save_readme(output_folder, category_counters, year_counters):
    readme_path = output_folder / "README.txt"
    total = sum(category_counters.values())

    lines = [
        "Nexis Corpus XML",
        "",
        f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Total texts: {total}",
        "",
        "Texts by category:",
    ]

    for category in sorted(category_counters):
        lines.append(f"- {category}: {category_counters[category]}")

    lines.extend([
        "",
        "Texts by year:",
    ])

    for year in sorted(year_counters):
        lines.append(f"- {year}: {year_counters[year]}")

    readme_path.write_text("\n".join(lines), encoding="utf-8")
    print("README created:", readme_path)

def next_output_path(base_folder, category, counters):
    folder_name = CATEGORY_FOLDERS.get(category, category)
    category_folder = base_folder / folder_name
    category_folder.mkdir(parents=True, exist_ok=True)

    counters[category] = counters.get(category, 0) + 1

    filename = f"NYTNews{category}{counters[category]:03}.xml"

    return category_folder / filename


def parse_nexis_docx(docx_path):
    paragraphs = get_paragraphs(docx_path)

    body_index = next(
        i for i, p in enumerate(paragraphs)
        if p.strip().lower() == "body"
    )

    before_body = paragraphs[:body_index]
    after_body = paragraphs[body_index + 1:]

    before_body = [
        p for p in before_body
        if not is_boilerplate(p)
        and not p.lower().startswith("length:")
    ]

    headline = before_body[0] if len(before_body) > 0 else ""
    source = before_body[1] if len(before_body) > 1 else ""
    date = before_body[2] if len(before_body) > 2 else ""

    section = ""
    author = ""
    dateline = ""
    highlight = ""

    for i, paragraph in enumerate(before_body):
        lower = paragraph.lower().strip()

        if lower.startswith("section:"):
            section = get_value_after_colon(paragraph)

        elif lower.startswith("byline:"):
            author = get_value_after_colon(paragraph)

        elif lower.startswith("dateline:"):
            dateline = get_value_after_colon(paragraph)

        elif lower.startswith("highlight:"):
            highlight = get_value_after_colon(paragraph)

        elif lower == "highlight":
            if i + 1 < len(before_body):
                highlight = before_body[i + 1].strip()

    abstract_highlight, body_paragraphs = extract_abstract_and_full_text(after_body)

    if abstract_highlight:
        highlight = abstract_highlight

    return {
        "headline": headline,
        "source": source,
        "date": date,
        "section": section,
        "author": author,
        "dateline": dateline,
        "highlight": highlight,
        "body": body_paragraphs,
    }


def save_xml(data, output_path):
    attributes = {}

    for key in ["source", "date", "section", "author", "dateline", "highlight"]:
        if data.get(key):
            attributes[key] = data[key]

    root = Element("text", attributes)

    headline_el = SubElement(root, "headline")
    headline_el.text = data["headline"]

    for paragraph in data["body"]:
        p_el = SubElement(root, "p")
        p_el.text = paragraph

    indent(root, space="  ")

    tree = ElementTree(root)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)


counters = {}
year_counters = {}

docx_files = sorted({p.resolve() for p in PASTA_ENTRADA.glob("*.docx") if p.is_file()})

for docx_path in sorted(docx_files):
    try:
        data = parse_nexis_docx(docx_path)

        category = classify_section(data["section"])
        year = extract_year(data["date"])
        year_counters[year] = year_counters.get(year, 0) + 1

        output_path = next_output_path(PASTA_SAIDA, category, counters)

        save_xml(data, output_path)

        print("Created:", output_path)

    except StopIteration:
        print("ERROR: Body not found in:", docx_path.name)

    except Exception as error:
        print("ERROR in:", docx_path.name)
        print(error)

save_readme(PASTA_SAIDA, counters, year_counters)


