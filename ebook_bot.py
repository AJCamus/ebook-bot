# ebook_bot.py — the Kindle-prep bot, v1
#
# Takes a .docx manuscript and produces:
#   1. A clean, KDP-ready EPUB (built by hand — no extra libraries needed)
#   2. A listing-pack.md with everything Amazon's upload form asks for
#
# Run:  python ebook_bot.py
# (It will ask for the manuscript file, then walk you through the rest,
#  like posting to Facebook but for books.)
#
# This file is ALSO a Python lesson. Read the comments.

import os
import re
import sys
import zipfile
from datetime import date
from docx import Document

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def slugify(text):
    # "AI Training Learning Path" -> "ai-training-learning-path"
    # Used for folder and file names.
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-") or "untitled-book"


def esc(text):
    # Escape the five characters that break XHTML.
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;"))


def para_to_html(p):
    """Turn one docx paragraph into an XHTML string, keeping bold/italic."""
    style = p.style.name if p.style else "Normal"
    # Build the inner text run by run so bold/italic survive.
    inner = ""
    for run in p.runs:
        t = esc(run.text)
        if not t:
            continue
        if run.bold and run.italic:
            t = f"<strong><em>{t}</em></strong>"
        elif run.bold:
            t = f"<strong>{t}</strong>"
        elif run.italic:
            t = f"<em>{t}</em>"
        inner += t
    if not inner.strip():
        return ""  # skip empty paragraphs

    if style == "Title":
        return f"<h1 class=\"booktitle\">{inner}</h1>"
    if style == "Subtitle":
        return f"<p class=\"subtitle\">{inner}</p>"
    if style == "Heading 1":
        return f"<h1>{inner}</h1>"
    if style == "Heading 2":
        return f"<h2>{inner}</h2>"
    if style == "Heading 3":
        return f"<h3>{inner}</h3>"
    if "List Bullet" in style:
        return f"<ul><li>{inner}</li></ul>"
    if "List Number" in style:
        return f"<ol><li>{inner}</li></ol>"
    return f"<p>{inner}</p>"


def ask(prompt, default=""):
    # Ask a question with a suggested default in [brackets].
    # Empty answer = take the default.
    hint = f" [{default}]" if default else ""
    ans = input(f"{prompt}{hint}: ").strip()
    return ans if ans else default


# ---------------------------------------------------------------------------
# Step 1: read the manuscript
# ---------------------------------------------------------------------------

def read_docx(path):
    doc = Document(path)
    title = ""
    subtitle = ""
    # chapters = list of (chapter_title, [html strings])
    chapters = []
    current = None

    for p in doc.paragraphs:
        style = p.style.name if p.style else "Normal"
        text = p.text.strip()

        if style == "Title" and not title:
            title = text
            continue
        if style == "Subtitle" and not subtitle:
            subtitle = text
            continue
        if style == "Heading 1":
            # A new chapter begins. Stash the old one first.
            if current:
                chapters.append(current)
            current = (text, [para_to_html(p)])
        elif current is None:
            # Text before the first chapter -> front matter chapter.
            current = ("Front matter", [])
            html = para_to_html(p)
            if html:
                current[1].append(html)
        else:
            html = para_to_html(p)
            if html:
                current[1].append(html)

    if current:
        chapters.append(current)
    return title, subtitle, chapters


# ---------------------------------------------------------------------------
# Step 2: build the EPUB (by hand — it's just a zip with rules)
# ---------------------------------------------------------------------------

CSS = """body{font-family:Georgia,serif; line-height:1.6; margin:5%%}
h1{font-size:1.5em; margin-top:1.5em}
h2{font-size:1.25em; margin-top:1.2em}
.booktitle{font-size:2em; text-align:center; margin-top:3em}
.subtitle{text-align:center; font-style:italic; color:#555}
p{text-align:justify; margin:0.6em 0}
li{margin:0.3em 0}
"""

CONTAINER_XML = """<?xml version="1.0"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles><rootfile full-path="OEBPS/content.opf"
    media-type="application/oebps-package+xml"/></rootfiles>
</container>"""

NCX_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
<head><meta name="dtb:uid" content="{uid}"/></head>
<docTitle><text>{title}</text></docTitle>
<navMap>{items}</navMap>
</ncx>"""

NAV_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">
<head><title>Contents</title><link rel="stylesheet" href="style.css"/></head>
<body><nav epub:type="toc"><h1>Contents</h1><ol>{items}</ol></nav></body>
</html>"""

CHAPTER_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>{ctitle}</title><link rel="stylesheet" href="style.css"/></head>
<body>{body}</body>
</html>"""


def build_epub(title, author, chapters, out_path, cover_path=None):
    uid = "urn:uuid:" + slugify(title) + "-" + str(date.today())
    files = []   # (arcname, data) in spine order

    # One XHTML file per chapter.
    for i, (ctitle, blocks) in enumerate(chapters, start=1):
        fname = f"chapter{i}.xhtml"
        body = "\n".join(blocks)
        files.append((fname, CHAPTER_TEMPLATE.format(
            ctitle=esc(ctitle), body=body).encode("utf-8")))

    # content.opf — the manifest + metadata + spine.
    manifest = ['<item id="css" href="style.css" media-type="text/css"/>',
                '<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>',
                '<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>']
    spine = []
    for i, (fname, _) in enumerate(files, start=1):
        manifest.append(
            f'<item id="ch{i}" href="{fname}" media-type="application/xhtml+xml"/>')
        spine.append(f'<itemref idref="ch{i}"/>')
    cover_meta = ""
    if cover_path:
        ext = os.path.splitext(cover_path)[1].lower()
        mt = "image/jpeg" if ext in (".jpg", ".jpeg") else "image/png"
        manifest.append(f'<item id="cover" href="cover{ext}" media-type="{mt}" properties="cover-image"/>')
        cover_meta = '<meta name="cover" content="cover"/>'

    opf = f"""<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" unique-identifier="uid" version="3.0">
<metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
<dc:identifier id="uid">{uid}</dc:identifier>
<dc:title>{esc(title)}</dc:title>
<dc:creator>{esc(author)}</dc:creator>
<dc:language>en</dc:language>
<dc:date>{date.today()}</dc:date>
{cover_meta}
</metadata>
<manifest>{''.join(manifest)}</manifest>
<spine toc="ncx">{''.join(spine)}</spine>
</package>"""

    # toc.ncx (older Kindles) and nav.xhtml (newer).
    ncx_items = "".join(
        f'<navPoint id="np{i}" playOrder="{i}"><navLabel><text>{esc(c)}</text>'
        f"</navLabel><content src=\"{f}\"/></navPoint>"
        for i, ((c, _), (f, _)) in enumerate(zip(chapters, files), start=1))
    nav_items = "".join(
        f"<li><a href=\"{f}\">{esc(c)}</a></li>"
        for (c, _), (f, _) in zip(chapters, files))

    # Write the zip. The mimetype file MUST be first and uncompressed.
    with zipfile.ZipFile(out_path, "w") as z:
        z.writestr("mimetype", "application/epub+zip",
                   compress_type=zipfile.ZIP_STORED)
        z.writestr("META-INF/container.xml", CONTAINER_XML,
                   compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("OEBPS/content.opf", opf.encode("utf-8"),
                   compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("OEBPS/toc.ncx",
                   NCX_TEMPLATE.format(uid=uid, title=esc(title),
                                       items=ncx_items).encode("utf-8"),
                   compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("OEBPS/nav.xhtml",
                   NAV_TEMPLATE.format(items=nav_items).encode("utf-8"),
                   compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("OEBPS/style.css", CSS.encode("utf-8"),
                   compress_type=zipfile.ZIP_DEFLATED)
        for fname, data in files:
            z.writestr("OEBPS/" + fname, data,
                       compress_type=zipfile.ZIP_DEFLATED)
        if cover_path:
            ext = os.path.splitext(cover_path)[1].lower()
            with open(cover_path, "rb") as fh:
                z.writestr("OEBPS/cover" + ext, fh.read(),
                           compress_type=zipfile.ZIP_DEFLATED)


# ---------------------------------------------------------------------------
# Step 3: the listing pack — everything KDP's form asks for
# ---------------------------------------------------------------------------

LISTING_TEMPLATE = """# Listing pack — {title}

Generated {today} by the ebook bot. Fill the blanks, then copy/paste
into Amazon KDP. Nothing here is final until you say so.

## The book file
- EPUB: `{epub}`
- Upload this at kdp.amazon.com → Create → Kindle eBook.

## Title / subtitle
- Title: {title}
- Subtitle: {subtitle}

## Author
- Author name: {author}
- Bio (edit me!):
> {bio}

## Description (the blurb — edit me!)
{blurb}

## Keywords (7 slots — comma separated)
{keywords}

## Categories (pick 2–3 on KDP)
- Suggested: {categories}

## Cover
- KDP wants 2560 × 1600 px (ratio 1.6:1), JPEG or TIFF.
- Cover used: {cover}
- No cover yet? Prompt an image model with:
  "Book cover art, {title}. {cover_prompt}"

## Pricing
- Sweet spot for 70% royalty: $2.99–$9.99.
- Your call: {price}

## Pre-upload checklist
- [ ] Read the EPUB on a phone/tablet once (catches weird breaks)
- [ ] Cover is 2560×1600 and looks good as a thumbnail
- [ ] Blurb has no typos and ends with a hook
- [ ] Bio is in third person
- [ ] NOT enrolled in KDP Select (keeps your direct-sales options open)
"""


# ---------------------------------------------------------------------------
# Main — the friendly interview
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("  📚 EBOOK BOT v1 — manuscript in, Kindle-ready out")
    print("=" * 60)

    default_doc = os.path.expanduser(
        "~/workspace/goals/seeds-seedlings-companion-prototype/files/"
        "ai-training-learning-path/AI Training Learning Path.docx")
    src = ask("\nManuscript (.docx)", default_doc).strip('"')
    if not os.path.exists(src):
        print(f"\nCan't find that file: {src}\n")
        sys.exit(1)

    print("\nReading manuscript…")
    title, subtitle, chapters = read_docx(src)
    n_ch = len(chapters)
    print(f"Found: {n_ch} chapters. Title in file: {title or '(none)'}")

    print("\n--- About the book (Enter = accept suggestion) ---")
    title = ask("Book title", title or "Untitled")
    subtitle = ask("Subtitle", subtitle or "")
    author = ask("Author name", "Aaron Camus")
    blurb = ask("One-paragraph blurb", "")
    if not blurb:
        print("  (tip: write the blurb later — the pack will remind you)")
        blurb = "[WRITE ME: what the reader gets, who it's for, one hook]"
    bio = ask("Author bio (one or two sentences)",
              "Aaron Camus writes and builds at the intersection of "
              "gardening and AI.")
    keywords = ask("Keywords (comma separated)",
                   "AI training, machine learning basics, beginners guide")
    categories = ask("Categories",
                     "Computers & Technology / AI & Machine Learning")
    price = ask("Price", "4.99")
    cover = ask("Cover image path (blank = none yet)", "")
    if cover and not os.path.exists(cover):
        print("  (cover not found — noted as missing)")
        cover = ""
    cover_prompt = ask("Cover art prompt (if generating later)",
                       "garden rows under a sunrise, warm and hopeful")

    slug = slugify(title)
    # Output goes next to wherever you run the bot from — no surprises.
    outdir = os.path.join(os.getcwd(), "ebook-bot-output", slug)
    os.makedirs(outdir, exist_ok=True)
    epub_path = os.path.join(outdir, slug + ".epub")

    print("\nBuilding EPUB…")
    build_epub(title, author, chapters, epub_path,
               cover_path=cover or None)

    pack = LISTING_TEMPLATE.format(
        title=title, subtitle=subtitle or "(none)", author=author,
        bio=bio, blurb=blurb, keywords=keywords, categories=categories,
        cover=cover or "NONE YET — see prompt below",
        cover_prompt=cover_prompt, price="$" + price,
        epub=os.path.basename(epub_path), today=date.today())
    pack_path = os.path.join(outdir, "listing-pack.md")
    with open(pack_path, "w") as f:
        f.write(pack)

    print(f"\nDone! Your Kindle-ready package is in:\n  {outdir}")
    print(f"  📕 {os.path.basename(epub_path)}")
    print("  📋 listing-pack.md  (everything KDP's form asks for)")
    print("\nNext: open the EPUB on your phone, read it once, then upload.")


if __name__ == "__main__":
    main()

# TRY THIS (homework):
#  1. Open the generated EPUB: rename .epub -> .zip and peek inside.
#     It's just files. Formats are less scary than they look.
#  2. Add .txt support: read a text file, split on blank lines into
#     paragraphs, treat lines starting with "# " as Heading 1.
#  3. The blurb prompt: can you make the bot SUGGEST a blurb from the
#     manuscript's first chapter? (Hint: read the first 3 paragraphs.)
