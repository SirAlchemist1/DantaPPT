"""Build the Danta-Med pitch deck.

Usage:  python build.py            -> single-file HTML
        python build.py --export   -> HTML + PDF + PPTX + PPSX + slide PNGs

Demo media (cropped, re-encoded H.264) lives in /source/media and is inlined.
Optional drop-in (next to deck.html in /source):
  danta-workflow.png   replaces the drawn workflow on slide 03
"""
import base64, pathlib, re, sys, asyncio, zipfile

SRC = pathlib.Path(__file__).resolve().parent
OUT = SRC.parent
A = SRC / "assets"
NAME = "DANTA-Med_ICT_Award_2026"


def b64(p):
    return base64.b64encode(p.read_bytes()).decode()


FONTS = [
    ("Bric", "bricolage-latin.woff2", "200 800", "normal", "U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD"),
    ("Bric", "bricolage-latinext.woff2", "200 800", "normal", "U+0100-02AF,U+1E00-1EFF,U+2020,U+20A0-20AB,U+20AD-20C0,U+2113,U+2C60-2C7F,U+A720-A7FF"),
    ("ISerif", "instrument-serif-italic.woff2", "400", "italic", None),
    ("Anek", "anek-devanagari.woff2", "100 800", "normal", None),
    ("GMono", "mono-400.woff2", "400", "normal", None),
    ("GMono", "mono-600.woff2", "600", "normal", None),
]


def build_html():
    css = []
    for fam, f, w, st, rng in FONTS:
        r = f";unicode-range:{rng}" if rng else ""
        css.append(f'@font-face{{font-family:"{fam}";src:url(data:font/woff2;base64,{b64(A / f)}) format("woff2");'
                   f"font-weight:{w};font-style:{st};font-display:block{r}}}")
    html = (SRC / "deck.html").read_text()
    html = html.replace("/*FONTS*/", "\n".join(css))
    logo = (A / "danta-logo.svg").read_text()
    vb = re.search(r'viewBox="([^"]+)"', logo).group(1)
    inner = re.search(r"<svg[^>]*>(.*)</svg>", logo, re.S).group(1)
    html = html.replace("<!--LOGO_SYMBOL-->",
                        f'<svg width="0" height="0" style="position:absolute"><symbol id="danta-logo" viewBox="{vb}">{inner}</symbol></svg>')
    html = html.replace("url(ICT_L)", f"url(data:image/png;base64,{b64(A / 'ict-light.png')})")
    html = html.replace("url(ICT_D)", f"url(data:image/png;base64,{b64(A / 'ict-dark.png')})")
    wf = SRC / "danta-workflow.png"
    if wf.exists():
        html = html.replace('src="danta-workflow.png"', f'src="data:image/png;base64,{b64(wf)}"')
    mime = {".mp4": "video/mp4", ".jpg": "image/jpeg", ".png": "image/png"}
    for f in sorted((SRC / "media").iterdir()):
        if f.suffix in mime:
            html = html.replace(f'"media/{f.name}"', f'"data:{mime[f.suffix]};base64,{b64(f)}"')
    out = OUT / f"{NAME}.html"
    out.write_text(html)
    print("HTML ->", out)
    return out


async def export(html):
    from playwright.async_api import async_playwright
    shots = OUT / "slides-png"
    shots.mkdir(exist_ok=True)
    async with async_playwright() as p:
        b = await p.chromium.launch()
        pg = await b.new_page(viewport={"width": 1920, "height": 1080})
        await pg.goto(html.as_uri() + "?print")
        await pg.wait_for_function("document.body.dataset.ready==1", timeout=20000)
        await pg.evaluate("document.fonts.ready")
        await pg.wait_for_timeout(600)
        n = await pg.evaluate("document.querySelectorAll('.slide').length")
        notes = await pg.evaluate("[...document.querySelectorAll('.slide')].map(s=>{const a=s.querySelector('aside.notes');return a?a.textContent.trim():''})")
        files = []
        for i in range(n):
            f = shots / f"slide-{i:02d}.png"
            await pg.locator(".slide").nth(i).screenshot(path=str(f))
            files.append(f)
        await pg.pdf(path=str(OUT / f"{NAME}.pdf"), width="1920px", height="1080px",
                     print_background=True, margin={"top": "0", "bottom": "0", "left": "0", "right": "0"})
        await b.close()
    print("PDF  ->", OUT / f"{NAME}.pdf")
    return files, notes


def pptx(files, notes):
    from pptx import Presentation
    from pptx.util import Emu
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    for f, note in zip(files, notes):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes.add_picture(str(f), 0, 0, prs.slide_width, prs.slide_height)
        if note:
            s.notes_slide.notes_text_frame.text = note
    p = OUT / f"{NAME}.pptx"
    prs.save(p)
    print("PPTX ->", p)
    ppsx = OUT / f"{NAME}.ppsx"
    with zipfile.ZipFile(p) as zi, zipfile.ZipFile(ppsx, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            data = zi.read(it.filename)
            if it.filename == "[Content_Types].xml":
                data = data.replace(b"presentationml.presentation.main+xml", b"presentationml.slideshow.main+xml")
            zo.writestr(it, data)
    print("PPSX ->", ppsx)


if __name__ == "__main__":
    h = build_html()
    if "--export" in sys.argv:
        pptx(*asyncio.run(export(h)))
