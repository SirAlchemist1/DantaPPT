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


def build_html():
    html = (SRC / "deck.html").read_text()
    html = re.sub(r'url\("assets/([\w.-]+\.woff2)"\)',
                  lambda m: f'url(data:font/woff2;base64,{b64(A / m.group(1))})', html)
    html = re.sub(r"url\(assets/([\w.-]+\.png)\)",
                  lambda m: f"url(data:image/png;base64,{b64(A / m.group(1))})", html)
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
