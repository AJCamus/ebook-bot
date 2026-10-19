# Ebook Bot 📚

Takes a `.docx` manuscript and produces everything you need to publish on
Amazon KDP:

1. **A clean, KDP-ready EPUB** — built by hand with nothing but the standard
   library plus `python-docx` (no ebooklib needed)
2. **A `listing-pack.md`** — every field Amazon's upload form asks for
   (title, blurb, categories, keywords, pricing notes), ready to copy-paste

## How it works

Run it, and it interviews you — title, author, blurb, categories — the way
posting to Facebook walks you through a post, but for books. Then it:

- Reads your `.docx`, splitting it into chapters on Heading 1
- Converts paragraphs to clean XHTML (bold/italic survive the trip)
- Zips it all into a valid EPUB (correct mimetype-first ordering, OPF manifest,
  NCX + nav)
- Writes the listing pack alongside

```bash
pip install python-docx
python ebook_bot.py
```

Output lands in `./ebook-bot-output/<slug>/`.

## Also a Python lesson

The source is heavily commented as a learn-Python-by-reading walkthrough —
helpers, EPUB spec quirks, and why each step exists. Read it top to bottom;
it teaches while it works.

Built by A.J. Camus & Ember. This bot produced real live Kindle books
(*Growing With Muse*, *AI Training Learning Path* — the Gardener's A.I. Library).
