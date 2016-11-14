openoffice-macros
========

My list of macros for LibreOffice / OpenOffice.org. Written in Star Basic. Licensed under GPL v3 – see [LICENSE.md](LICENSE.md) or [LICENSE.txt](LICENSE.txt).

# Procedure list

## Count Subheadings

`CountSubheadings.bas` – Reports a count of second-level headings for each first-level heading inthe document.

## Find Broken internal links

`FindBrokenInternalLinks.bas` – Looks for broken internal links in the current document. Suggests fixes to links to outline elements (headings), whenever possible. Optionally, auto-fixes links where the text matches but the numbering doesn’t.

*TODO:* Suggest opening the link for manual editing.

## Straight Quotes to Curly Quotes

`StraightQuotes2CurlyQuotes.bas` – “Educates” quotes by turning single and double straight quotes to curly opening or closing single or double quotes, as the case may be. Follows English language conventions.

# Changelog

#### 1.1

* GPL v3 licensing info
* `FindBrokenInternalLinks.bas`: possibility of auto-fixing links where the text matches but the numbering doesn’t

#### 1.0

* New procedure: `CountSubheadings.bas`
* New procedure: `FindBrokenInternalLinks.bas`

#### 0.0

* Initial commit: `StraightQuotes2CurlyQuotes.bas`
