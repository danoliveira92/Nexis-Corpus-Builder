from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from xml.etree.ElementTree import Element, SubElement, ElementTree, indent
import re

from docx import Document


SCRIPT_VERSION = "2026-06-17 guardian-heading-fix"
MIN_ARTICLES_FOR_SECTION_FOLDER = 3

input_folder = input("Paste the folder containing the DOCX files: ").strip().strip('"')
output_parent = input("Paste the folder where the corpus folder should be created: ").strip().strip('"')
corpus_name = input("Name this corpus: ").strip()

if not corpus_name:
    corpus_name = "CORPUS_XML"

corpus_name = re.sub(r'[<>:"/\\|?*]+', "_", corpus_name)

PASTA_ENTRADA = Path(input_folder)
PASTA_SAIDA = Path(output_parent) / corpus_name

if PASTA_SAIDA.exists() and any(PASTA_SAIDA.iterdir()):
    raise SystemExit(
        f"Output corpus folder already exists and is not empty: {PASTA_SAIDA}\n"
        "Choose a new corpus name or manually delete/rename the existing folder before running again."
    )

PASTA_SAIDA.mkdir(parents=True, exist_ok=True)

SOURCE_PREFIXES = {
    "nyt": "NYT",
    "guardian": "THEGUARDIAN",
    "unknown": "UNKNOWN",
}

SECTION_LABEL_OVERRIDES = {
    "australia news": "World",
    "australianews": "World",
    "world news": "World",
    "worldnews": "World",
    "world": "World",
    "football": "Sports",
    "sport": "Sports",
    "sports": "Sports",
    "media": "Media",
    "film": "Media",
    "movies": "Media",
    "film and movies": "Media",
    "film & movies": "Media",
    "u.s.": "US",
    "us": "US",
    "us news": "US",
    "usnews": "US",
    "u.s. news": "US",
    "nyregion": "NYRegion",
    "ny region": "NYRegion",
    "real estate": "RealEstate",
}

NYT_BOILERPLATE_STARTS = [
    "the times is committed to",
    "follow the new york times",
    "reporting was contributed by",
]

NYT_BOILERPLATE_CONTAINS = [
    "contributed reporting",
    "by getty images",
]

GUARDIAN_REMOVE_SECTION_HEADINGS = {
    "spotlight",
    "top picks",
    "rights and freedom",
    "opinion and analysis",
    "southern frontlines",
    "in pictures",
    "read this",
    "the world in brief",
}

GUARDIAN_REMOVE_SECTION_PREFIXES = [
    "read this:",
    "the world in brief",
]


GENERIC_BOILERPLATE_STARTS = [
    "copyright",
    "load-date",
]



def fix_mojibake(text):
    text = str(text)
    mojibake_fragment = re.compile(r"[\u00c2\u00c3][\u0080-\u00ff\u0192]+")

    def repair_fragment(match):
        current = match.group(0)

        for _ in range(3):
            try:
                fixed = current.encode("cp1252").decode("utf-8")
            except UnicodeError:
                break

            if fixed == current:
                break

            current = fixed

        return current

    return mojibake_fragment.sub(repair_fragment, text)
def normalize_text(text):
    text = fix_mojibake(str(text))
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

    for paragraph in doc.paragraphs:
        text = normalize_text(paragraph.text)

        if text:
            paragraphs.append(text)

    return paragraphs


def get_value_after_colon(text):
    if ":" not in text:
        return ""

    return text.split(":", 1)[1].strip()


def extract_year(date_text):
    match = re.search(r"\b(19|20)\d{2}\b", date_text)

    if match:
        return match.group(0)

    return "Unknown"


def simplify_for_matching(text):
    return re.sub(r"\s+", " ", text.lower().strip())


def sanitize_filename_part(text, default="Other"):
    text = normalize_text(str(text))
    text = text.replace("&", " and ")
    text = re.sub(r'[<>:"/\\|?*]+', "_", text)
    text = re.sub(r"\s+", "_", text.strip())
    text = re.sub(r"_+", "_", text)
    text = text.strip("._ ")

    return text or default


def make_section_label(section):
    if not section:
        return "Other"

    base_section = normalize_text(section).split(";", 1)[0].strip()

    if not base_section:
        return "Other"

    key = simplify_for_matching(base_section)

    if key in SECTION_LABEL_OVERRIDES:
        return SECTION_LABEL_OVERRIDES[key]

    words = re.split(r"[\s_/-]+", base_section)
    label = "".join(word[:1].upper() + word[1:].lower() for word in words if word)
    label = sanitize_filename_part(label)

    return label or "Other"


def detect_source_type(source, paragraphs):
    search_text = " ".join([source] + paragraphs[:8]).lower()

    if "new york times" in search_text:
        return "nyt"

    if "guardian" in search_text:
        return "guardian"

    return "unknown"


def log_removal(removed_items, reason, paragraph):
    paragraph = normalize_text(paragraph)

    if paragraph:
        removed_items.append({
            "reason": reason,
            "text": paragraph,
        })


def generic_removal_reason(paragraph):
    lower = simplify_for_matching(paragraph)

    if lower == "end of document":
        return "generic boilerplate: end of document"

    if lower.startswith("length:"):
        return "generic metadata: length"

    for start in GENERIC_BOILERPLATE_STARTS:
        if lower.startswith(start):
            return f"generic boilerplate: {start}"

    return None


def nyt_removal_reason(paragraph):
    lower = simplify_for_matching(paragraph)

    for item in NYT_BOILERPLATE_CONTAINS:
        if item in lower:
            return f"NYT boilerplate contains: {item}"

    for start in NYT_BOILERPLATE_STARTS:
        if lower.startswith(start):
            return f"NYT boilerplate starts with: {start}"

    return None


def guardian_removal_reason(paragraph):
    lower = simplify_for_matching(paragraph)

    if lower.startswith("related:"):
        return "Guardian related link"

    return None


def get_source_specific_removal_reason(paragraph, source_type):
    if source_type == "nyt":
        return nyt_removal_reason(paragraph)

    if source_type == "guardian":
        return guardian_removal_reason(paragraph)

    return None


def clean_metadata_paragraphs(paragraphs, removed_items):
    cleaned = []

    for paragraph in paragraphs:
        reason = generic_removal_reason(paragraph)

        if reason:
            log_removal(removed_items, reason, paragraph)
            continue

        cleaned.append(paragraph)

    return cleaned



def is_guardian_remove_section_heading(paragraph):
    lower = simplify_for_matching(paragraph)
    first_line = simplify_for_matching(paragraph.split("\n", 1)[0])

    candidates = [lower, first_line]

    for candidate in candidates:
        if candidate in GUARDIAN_REMOVE_SECTION_HEADINGS:
            return True

        if any(candidate.startswith(f"{heading}:") for heading in GUARDIAN_REMOVE_SECTION_HEADINGS):
            return True

        if any(candidate.startswith(prefix) for prefix in GUARDIAN_REMOVE_SECTION_PREFIXES):
            return True

    return False

def clean_body_paragraphs(paragraphs, source_type, removed_items):
    cleaned = []
    index = 0

    while index < len(paragraphs):
        paragraph = normalize_text(paragraphs[index])

        if not paragraph:
            index += 1
            continue

        if source_type == "guardian" and is_guardian_remove_section_heading(paragraph):
            heading = paragraph
            log_removal(removed_items, "Guardian section heading", heading)

            if index + 1 < len(paragraphs):
                following_paragraph = normalize_text(paragraphs[index + 1])

                if is_guardian_remove_section_heading(following_paragraph):
                    index += 1
                else:
                    log_removal(
                        removed_items,
                        f"Guardian paragraph after section heading: {heading}",
                        following_paragraph,
                    )
                    index += 2
            else:
                index += 1

            continue

        reason = generic_removal_reason(paragraph)

        if not reason:
            reason = get_source_specific_removal_reason(paragraph, source_type)

        if reason:
            log_removal(removed_items, reason, paragraph)
            index += 1
            continue

        cleaned.append(paragraph)
        index += 1

    return cleaned


def extract_abstract_and_full_text(body_paragraphs, source_type, removed_items):
    highlight = ""
    body = body_paragraphs
    upper_paragraphs = [paragraph.upper().strip() for paragraph in body_paragraphs]

    if "ABSTRACT" in upper_paragraphs:
        abstract_index = upper_paragraphs.index("ABSTRACT")
        log_removal(removed_items, "Nexis label", body_paragraphs[abstract_index])

        full_text_index = None
        if "FULL TEXT" in upper_paragraphs:
            full_text_index = upper_paragraphs.index("FULL TEXT")
            log_removal(removed_items, "Nexis label", body_paragraphs[full_text_index])

        if full_text_index is not None and full_text_index > abstract_index:
            abstract_paragraphs = body_paragraphs[abstract_index + 1:full_text_index]
            body = body_paragraphs[full_text_index + 1:]
        else:
            abstract_paragraphs = body_paragraphs[abstract_index + 1:]
            body = body_paragraphs[:abstract_index]

        abstract_paragraphs = clean_body_paragraphs(abstract_paragraphs, source_type, removed_items)
        highlight = " ".join(abstract_paragraphs).strip()

    elif "FULL TEXT" in upper_paragraphs:
        full_text_index = upper_paragraphs.index("FULL TEXT")
        log_removal(removed_items, "Nexis label", body_paragraphs[full_text_index])
        body = body_paragraphs[full_text_index + 1:]

    body = clean_body_paragraphs(body, source_type, removed_items)

    return highlight, body


def find_body_index(paragraphs):
    return next(
        index for index, paragraph in enumerate(paragraphs)
        if paragraph.strip().lower() == "body"
    )


def parse_nexis_docx(docx_path):
    removed_items = []
    paragraphs = get_paragraphs(docx_path)
    body_index = find_body_index(paragraphs)

    before_body = paragraphs[:body_index]
    after_body = paragraphs[body_index + 1:]

    source_guess = before_body[1] if len(before_body) > 1 else ""
    source_type = detect_source_type(source_guess, paragraphs)

    before_body = clean_metadata_paragraphs(before_body, removed_items)

    headline = before_body[0] if len(before_body) > 0 else ""
    source = before_body[1] if len(before_body) > 1 else ""
    date = before_body[2] if len(before_body) > 2 else ""

    section = ""
    author = ""
    dateline = ""
    highlight = ""

    for index, paragraph in enumerate(before_body):
        lower = simplify_for_matching(paragraph)

        if lower.startswith("section:"):
            section = get_value_after_colon(paragraph)

        elif lower.startswith("byline:"):
            author = get_value_after_colon(paragraph)

        elif lower.startswith("dateline:"):
            dateline = get_value_after_colon(paragraph)

        elif lower.startswith("highlight:"):
            highlight = get_value_after_colon(paragraph)

        elif lower == "highlight":
            if index + 1 < len(before_body):
                highlight = before_body[index + 1].strip()

    abstract_highlight, body_paragraphs = extract_abstract_and_full_text(
        after_body,
        source_type,
        removed_items,
    )

    if abstract_highlight:
        highlight = abstract_highlight

    source_type = detect_source_type(source, paragraphs)
    source_prefix = SOURCE_PREFIXES.get(source_type, SOURCE_PREFIXES["unknown"])
    section_label = make_section_label(section)
    year = extract_year(date)

    return {
        "docx_path": docx_path,
        "headline": headline,
        "source": source,
        "source_type": source_type,
        "source_prefix": source_prefix,
        "date": date,
        "year": year,
        "section": section,
        "section_label": section_label,
        "author": author,
        "dateline": dateline,
        "highlight": highlight,
        "body": body_paragraphs,
        "removed_items": removed_items,
    }


def choose_final_section(article, section_counts):
    section_label = article["section_label"]

    if section_label == "Other":
        return "Other"

    if section_counts[section_label] >= MIN_ARTICLES_FOR_SECTION_FOLDER:
        return section_label

    return "Other"


def build_output_path(base_folder, article, filename_counters):
    source_prefix = article["source_prefix"]
    final_section = article["final_section"]
    year = article["year"]

    folder_name = sanitize_filename_part(final_section)
    category_folder = base_folder / folder_name
    category_folder.mkdir(parents=True, exist_ok=True)

    counter_key = (source_prefix, final_section, year)
    filename_counters[counter_key] += 1

    filename = f"{source_prefix}_{final_section}_{year}_{filename_counters[counter_key]:03}.xml"
    filename = sanitize_filename_part(filename, default="article.xml")

    if not filename.lower().endswith(".xml"):
        filename = f"{filename}.xml"

    return category_folder / filename


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


def save_readme(output_folder, articles, category_counters, year_counters, source_counters, errors):
    readme_path = output_folder.parent / f"{output_folder.name}_README.txt"
    total = len(articles)

    lines = [
        "Nexis Corpus XML",
        "",
        f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Total texts: {total}",
        f"Minimum articles required for a section folder: {MIN_ARTICLES_FOR_SECTION_FOLDER}",
        "",
        "Texts by source:",
    ]

    for source in sorted(source_counters):
        lines.append(f"- {source}: {source_counters[source]}")

    lines.extend([
        "",
        "Texts by category:",
    ])

    for category in sorted(category_counters):
        lines.append(f"- {category}: {category_counters[category]}")

    lines.extend([
        "",
        "Texts by year:",
    ])

    for year in sorted(year_counters):
        lines.append(f"- {year}: {year_counters[year]}")

    if errors:
        lines.extend([
            "",
            "Files not processed:",
        ])

        for filename, error in errors:
            lines.append(f"- {filename}: {error}")

    readme_path.write_text("\n".join(lines), encoding="utf-8")
    print("README created:", readme_path)


def save_removal_readme(output_folder, articles, errors):
    readme_path = output_folder.parent / f"{output_folder.name}_README_2.txt"

    lines = [
        "Nexis Corpus Removal Log",
        "",
        f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "This file lists paragraphs or lines removed from each processed article.",
    ]

    articles_with_removals = [article for article in articles if article["removed_items"]]

    if not articles_with_removals:
        lines.extend([
            "",
            "No removals were logged.",
        ])

    for article in articles_with_removals:
        lines.extend([
            "",
            f"Article: {article['docx_path'].name}",
            f"Output XML: {article.get('output_filename', '')}",
            f"Source: {article.get('source', '')}",
            f"Section: {article.get('section', '') or 'Other'}",
            "Removed:",
        ])

        for item in article["removed_items"]:
            lines.append(f"- {item['reason']}: {item['text']}")

    if errors:
        lines.extend([
            "",
            "Files not processed:",
        ])

        for filename, error in errors:
            lines.append(f"- {filename}: {error}")

    readme_path.write_text("\n".join(lines), encoding="utf-8")
    print("README_2 created:", readme_path)


def main():
    print(f"Nexis Corpus Builder version: {SCRIPT_VERSION}")
    print(f"Running script: {Path(__file__).resolve()}")

    counters = Counter()
    year_counters = Counter()
    source_counters = Counter()
    filename_counters = defaultdict(int)
    errors = []

    docx_files = sorted({path.resolve() for path in PASTA_ENTRADA.glob("*.docx") if path.is_file()})

    if not docx_files:
        print("No DOCX files found in:", PASTA_ENTRADA)
        return

    articles = []

    print("Scanning DOCX files...")

    for docx_path in docx_files:
        try:
            article = parse_nexis_docx(docx_path)
            articles.append(article)
        except StopIteration:
            errors.append((docx_path.name, "Body not found"))
            print("ERROR: Body not found in:", docx_path.name)
        except Exception as error:
            errors.append((docx_path.name, str(error)))
            print("ERROR in:", docx_path.name)
            print(error)

    section_counts = Counter(
        article["section_label"]
        for article in articles
        if article["section_label"] != "Other"
    )

    print("Saving XML files...")

    for article in articles:
        article["final_section"] = choose_final_section(article, section_counts)
        output_path = build_output_path(PASTA_SAIDA, article, filename_counters)

        article["output_path"] = output_path
        article["output_filename"] = output_path.name

        save_xml(article, output_path)

        counters[article["final_section"]] += 1
        year_counters[article["year"]] += 1
        source_counters[article["source_prefix"]] += 1

        print("Created:", output_path)

    save_readme(PASTA_SAIDA, articles, counters, year_counters, source_counters, errors)
    save_removal_readme(PASTA_SAIDA, articles, errors)


if __name__ == "__main__":
    main()
