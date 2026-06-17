# Nexis Corpus Builder

Nexis Corpus Builder is a Python script for converting batches of newspaper articles downloaded from Nexis into clean, structured XML files for corpus analysis.

The current version is designed for articles from **The New York Times** and **The Guardian**. It reads `.docx` files exported from Nexis, extracts article metadata, cleans unwanted boilerplate, normalizes text, classifies articles by section, and saves each article as an individual XML file.

The output is organized into section-based folders and can be used in corpus analysis tools such as LancsBox.

## What The Script Does

For each `.docx` file, the script:

1. reads the article text from the DOCX body;
2. extracts metadata such as headline, source, date, section, byline/author, dateline, and highlight/abstract;
3. identifies whether the article is from The New York Times or The Guardian;
4. applies general text normalization;
5. applies source-specific cleaning rules;
6. saves the cleaned article as an XML file;
7. organizes XML files into section-based folders;
8. creates summary README files describing the corpus and the removed content.

## Supported Sources

### The New York Times

For The New York Times articles, the script removes boilerplate such as:

- paragraphs containing `contributed reporting`;
- paragraphs beginning with `The Times is committed to`;
- paragraphs beginning with `Follow The New York Times`;
- paragraphs beginning with `Reporting was contributed by`;
- lines containing `by Getty Images`;
- copyright notices;
- length information;
- load-date lines;
- end-of-document markers.

Generated XML files from The New York Times begin with `NYT_`.

Example:

```text
NYT_Business_2025_001.xml
```

### The Guardian

For The Guardian articles, the script removes:

- paragraphs beginning with `Related:`;
- selected promotional or navigation section headings, including `SPOTLIGHT`, `TOP PICKS`, `RIGHTS AND FREEDOM`, `OPINION AND ANALYSIS`, `Southern frontlines`, `in pictures`, and `read this`.

When one of these Guardian section headings is found, the script removes the heading and the paragraph immediately after it.

Generated XML files from The Guardian begin with `THEGUARDIAN_`.

Example:

```text
THEGUARDIAN_World_2024_001.xml
```

## Section Classification

The script first scans all DOCX files before creating the final XML files. This allows it to count how many articles belong to each section.

A section receives its own folder only if it appears in more than two articles, meaning three or more articles. Sections with fewer articles are placed in `Other`.

Some related sections are automatically merged before counting. For example:

```text
AustraliaNews -> World
WorldNews -> World
World -> World

Football -> Sports
Sport -> Sports
Sports -> Sports

Media -> Media
Film -> Media
Movies -> Media
Film and Movies -> Media

US -> US
USNews -> US
U.S. News -> US
```

This prevents the corpus from being split into too many small folders.

## File Naming

Each XML file follows this naming pattern:

```text
SOURCE_SECTION_YEAR_NUMBER.xml
```

Examples:

```text
NYT_Business_2025_001.xml
THEGUARDIAN_World_2024_001.xml
NYT_Sports_2023_002.xml
```

The source prefix identifies the newspaper:

```text
NYT
THEGUARDIAN
```

## Output Structure

When the script runs, it asks for:

1. the folder containing the Nexis DOCX files;
2. the folder where the corpus should be created;
3. the name of the corpus.

Example output:

```text
My_Corpus/
  Business/
    NYT_Business_2025_001.xml
  World/
    THEGUARDIAN_World_2024_001.xml
  Sports/
    NYT_Sports_2023_001.xml
  Other/
    THEGUARDIAN_Other_2022_001.xml

My_Corpus_README.txt
My_Corpus_README_2.txt
```

The main corpus folder contains only XML files and section folders. The README files are saved outside the corpus folder so they do not interfere with corpus tools such as LancsBox.

## Generated README Files

The script creates two summary files.

### `CORPUS_NAME_README.txt`

This file summarizes the corpus. It includes:

- total number of processed articles;
- number of articles by source;
- number of articles by category;
- number of articles by year;
- files that could not be processed, if any.

### `CORPUS_NAME_README_2.txt`

This file logs removed material for each article. For each removal, it records:

- the original DOCX filename;
- the output XML filename;
- the article source;
- the article section;
- the removed text;
- the reason it was removed.

This makes the cleaning process more transparent and easier to review.

## Text Normalization

The script applies the same character normalization to all supported sources.

It normalizes:

- different dash characters to `-`;
- curly double quotation marks to `"`;
- curly single quotation marks to `'`;
- Windows and Unix line breaks;
- repeated blank lines;
- repeated spaces;
- non-breaking spaces;
- invisible Unicode characters;
- leading and trailing spaces;
- common mojibake produced by encoding problems.

The mojibake repair step is conservative and targets typical broken UTF-8 sequences found in Nexis exports. For example:

```text
ZÃºÃ±iga -> Zúñiga
AntÃ³nio -> António
SÃ£o Paulo -> São Paulo
FranÃ§ois -> François
```

## XML Format

Each article is saved as one XML file.

Example:

```xml
<?xml version='1.0' encoding='utf-8'?>
<text source="The New York Times" date="March 13, 2025 Friday" section="Business" author="Jane Doe" highlight="Example highlight.">
  <headline>Example Headline</headline>
  <p>First paragraph of the article.</p>
  <p>Second paragraph of the article.</p>
</text>
```

## Requirements

The script requires Python and the `python-docx` package.

Install `python-docx` with:

```powershell
pip install python-docx
```

## Important Note

This script does not download articles from Nexis, The New York Times, or The Guardian. It only processes `.docx` files that the user has already downloaded from Nexis.

The original DOCX files are not modified.

Generated XML files should be used according to the user's institutional access, copyright obligations, and research permissions.