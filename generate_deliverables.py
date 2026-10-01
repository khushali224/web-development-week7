from pathlib import Path
from struct import unpack
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree.ElementTree import Element, SubElement, tostring, register_namespace

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


OUTPUT_DIR = Path(__file__).parent / "submission"
GREEN = "173D32"
GREEN_MID = "275B48"
LIME = "D4F270"
CORAL = "E87952"
PAPER = "FBFCF8"
INK = "1D2C26"
MUTED = "66736C"
LINE = "E2E8E2"
WHITE = "FFFFFF"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
for prefix, namespace in (("w", W), ("r", R), ("wp", WP), ("a", A), ("pic", PIC)):
    register_namespace(prefix, namespace)


def word_tag(name):
    return f"{{{W}}}{name}"


def add_word_paragraph(body, text, kind="body", color=INK):
    paragraph = SubElement(body, word_tag("p"))
    ppr = SubElement(paragraph, word_tag("pPr"))
    spacing = SubElement(ppr, word_tag("spacing"))
    spacing.set(word_tag("after"), "130" if kind == "body" else "90")
    spacing.set(word_tag("line"), "285")
    run = SubElement(paragraph, word_tag("r"))
    rpr = SubElement(run, word_tag("rPr"))
    fonts = SubElement(rpr, word_tag("rFonts"))
    fonts.set(word_tag("ascii"), "Aptos")
    fonts.set(word_tag("hAnsi"), "Aptos")
    size = SubElement(rpr, word_tag("sz"))
    text_size = {"eyebrow": 18, "title": 44, "subtitle": 25, "h1": 32, "h2": 25, "body": 21, "small": 17, "bullet": 20}.get(kind, 21)
    size.set(word_tag("val"), str(text_size))
    color_node = SubElement(rpr, word_tag("color"))
    color_node.set(word_tag("val"), color)
    if kind in {"eyebrow", "title", "h1", "h2", "bullet"}:
        SubElement(rpr, word_tag("b"))
    node = SubElement(run, word_tag("t"))
    node.text = text
    if text[:1].isspace() or text[-1:].isspace():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return paragraph


def add_word_table(body, rows):
    table = SubElement(body, word_tag("tbl"))
    props = SubElement(table, word_tag("tblPr"))
    borders = SubElement(props, word_tag("tblBorders"))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = SubElement(borders, word_tag(edge))
        border.set(word_tag("val"), "single")
        border.set(word_tag("sz"), "5")
        border.set(word_tag("color"), LINE)
    widths = SubElement(table, word_tag("tblGrid"))
    for width in (2700, 6660):
        col = SubElement(widths, word_tag("gridCol"))
        col.set(word_tag("w"), str(width))
    for row_index, values in enumerate(rows):
        row = SubElement(table, word_tag("tr"))
        for value in values:
            cell = SubElement(row, word_tag("tc"))
            cell_props = SubElement(cell, word_tag("tcPr"))
            if row_index == 0:
                shade = SubElement(cell_props, word_tag("shd"))
                shade.set(word_tag("fill"), GREEN)
            paragraph = add_word_paragraph(cell, value, "small", WHITE if row_index == 0 else INK)
            paragraph_props = paragraph.find(word_tag("pPr"))
            if paragraph_props is not None:
                spacing = paragraph_props.find(word_tag("spacing"))
                if spacing is not None:
                    spacing.set(word_tag("after"), "45")


def add_word_image(body, image_path, relationship_id, figure_id, description):
    image_data = image_path.read_bytes()
    width_px, height_px = unpack(">II", image_data[16:24])
    width_emu = int(6.7 * 914400)
    height_emu = int(width_emu * height_px / width_px)
    paragraph = SubElement(body, word_tag("p"))
    ppr = SubElement(paragraph, word_tag("pPr"))
    alignment = SubElement(ppr, word_tag("jc"))
    alignment.set(word_tag("val"), "center")
    run = SubElement(paragraph, word_tag("r"))
    drawing = SubElement(run, word_tag("drawing"))
    inline = SubElement(drawing, f"{{{WP}}}inline", {"distT": "0", "distB": "0", "distL": "0", "distR": "0"})
    SubElement(inline, f"{{{WP}}}extent", {"cx": str(width_emu), "cy": str(height_emu)})
    SubElement(inline, f"{{{WP}}}docPr", {"id": str(figure_id), "name": description, "descr": description})
    frame = SubElement(inline, f"{{{WP}}}cNvGraphicFramePr")
    SubElement(frame, f"{{{A}}}graphicFrameLocks", {"noChangeAspect": "1"})
    graphic = SubElement(inline, f"{{{A}}}graphic")
    graphic_data = SubElement(graphic, f"{{{A}}}graphicData", {"uri": PIC})
    picture = SubElement(graphic_data, f"{{{PIC}}}pic")
    nonvisual = SubElement(picture, f"{{{PIC}}}nvPicPr")
    SubElement(nonvisual, f"{{{PIC}}}cNvPr", {"id": str(figure_id), "name": image_path.name, "descr": description})
    SubElement(nonvisual, f"{{{PIC}}}cNvPicPr")
    fill = SubElement(picture, f"{{{PIC}}}blipFill")
    SubElement(fill, f"{{{A}}}blip", {f"{{{R}}}embed": relationship_id})
    stretch = SubElement(fill, f"{{{A}}}stretch")
    SubElement(stretch, f"{{{A}}}fillRect")
    shape = SubElement(picture, f"{{{PIC}}}spPr")
    transform = SubElement(shape, f"{{{A}}}xfrm")
    SubElement(transform, f"{{{A}}}off", {"x": "0", "y": "0"})
    SubElement(transform, f"{{{A}}}ext", {"cx": str(width_emu), "cy": str(height_emu)})
    geometry = SubElement(shape, f"{{{A}}}prstGeom", {"prst": "rect"})
    SubElement(geometry, f"{{{A}}}avLst")


def add_page_break(body):
    paragraph = SubElement(body, word_tag("p"))
    run = SubElement(paragraph, word_tag("r"))
    SubElement(run, word_tag("br"), {word_tag("type"): "page"})


def make_report():
    document = Element(word_tag("document"))
    body = SubElement(document, word_tag("body"))
    add_word_paragraph(body, "DG INTERNS HUB  /  WEEK 07", "eyebrow", GREEN_MID)
    add_word_paragraph(body, "Internship discovery website", "title", GREEN)
    add_word_paragraph(body, "Project report  |  React JS", "subtitle", CORAL)
    add_word_paragraph(body, "A responsive, multi-page experience for exploring early-career opportunities.", "body", MUTED)
    add_word_paragraph(body, "PROJECT SNAPSHOT", "h2", GREEN_MID)
    add_word_table(body, [
        ("Item", "Implementation"),
        ("Frontend", "React 19 with functional components"),
        ("Routing", "React Router routes: Home, Jobs, Contact"),
        ("Tooling", "Vite, JavaScript, plain CSS, Lucide icons"),
        ("Opportunities", "Six illustrative internship listings with field, mode, location, tags, duration, and stipend"),
    ])

    add_word_paragraph(body, "1. Executive summary", "h1", GREEN)
    add_word_paragraph(body, "DG Interns Hub is a responsive internship discovery website built as a Week 7 React JS project. It gives students a single place to browse sample early-career opportunities, narrow listings by role or work arrangement, and reach an application contact through an email link. The experience includes Home, Jobs, and Contact pages, a consistent navigation system, responsive layouts, and an editable local dataset.")
    add_word_paragraph(body, "2. Objective and scope", "h1", GREEN)
    add_word_paragraph(body, "The objective is to make internship opportunities easy to discover across desktop and mobile, while demonstrating functional React components, multi-page routing, card-based listings, basic interactions, and responsive styling.")
    add_word_paragraph(body, "Included in this build", "h2", GREEN_MID)
    for item in [
        "Home page with introductory search and featured roles.",
        "Jobs page with keyword search, field filter, work-mode filter, result count, and an empty state.",
        "Contact page with required fields and client-side submission confirmation.",
        "Job cards with company mark, location, duration, tags, stipend, and an apply email link.",
        "Responsive navigation, desktop and mobile layouts, keyboard focus styling, and page metadata.",
    ]:
        add_word_paragraph(body, f"-  {item}", "bullet", INK)
    add_word_paragraph(body, "3. Page structure", "h1", GREEN)
    add_word_table(body, [
        ("Page", "Purpose and interactions"),
        ("Home", "Brand introduction, role search, featured listing cards, and a link into the full catalogue."),
        ("Jobs", "Six listings with query search across role/company/field/skills and filters for field and work arrangement."),
        ("Contact", "Contact details, required name/email/message inputs, and local confirmation state on submit."),
    ])
    add_word_paragraph(body, "4. Technical implementation", "h1", GREEN)
    add_word_paragraph(body, "The app is scaffolded with Vite. React Router DOM provides client-side routing, with BrowserRouter at the entry point and Routes/Route declarations for the three pages. The page components use React state for search input, selected filters, and contact-form confirmation. Job data is currently a local JavaScript array, making the sample catalogue easy to replace or connect to an API later. Lucide React supplies interface icons; CSS defines the layout, breakpoints, colors, typography, cards, and interaction states.")
    add_word_paragraph(body, "5. User journey and interactions", "h1", GREEN)
    add_word_paragraph(body, "A visitor can start on Home, enter a role or skill, and continue to Jobs with the query applied. On Jobs, the result set responds immediately to the text and select filters. Each Apply control opens a mailto link with the role in the subject. Contact validates required fields in the browser and changes to a confirmation view after submission; it does not transmit or store the message.")
    add_word_paragraph(body, "6. Visual design and accessibility", "h1", GREEN)
    add_word_paragraph(body, "The visual direction pairs deep green with a bright lime accent, warm coral details, and a pale neutral canvas. The home page uses a workplace photograph as its primary visual, while listings remain scan-friendly. Fixed breakpoints reorganize the hero, filters, cards, and contact form for narrow screens. Labels and accessible names are supplied for search controls; navigation uses semantic links and keyboard focus receives a visible outline.")
    add_word_paragraph(body, "7. Validation and known limitations", "h1", GREEN)
    add_word_paragraph(body, "The production build and Oxlint check passed during development, and editor diagnostics reported no React errors after the final source edits. A final build rerun and browser preview were not completed. The six listings are illustrative, the careers/contact email addresses are placeholders, and the contact form has no backend. The hero image and web fonts load from external services and require an internet connection.")
    add_word_paragraph(body, "8. Run and deployment", "h1", GREEN)
    add_word_paragraph(body, "Install packages with npm install, then start the local development server with npm run dev. Verify with npm run lint and npm run build. For Netlify or Vercel, use npm run build and publish the dist directory. Configure SPA fallback so direct visits and refreshes on /jobs and /contact serve index.html. Before launch, replace sample listings and email addresses with verified DG Interns Hub content.")
    add_word_paragraph(body, "9. Conclusion", "h1", GREEN)
    add_word_paragraph(body, "DG Interns Hub demonstrates the requested React fundamentals in a complete, responsive internship discovery experience. The project is ready for real listings and a backend integration once the sample content, application contacts, and message delivery flow are confirmed.")
    add_page_break(body)
    add_word_paragraph(body, "Website Screenshots", "h1", GREEN)
    add_word_paragraph(body, "Desktop captures of the implemented Home, Jobs, and Contact pages.", "body", MUTED)
    screenshots = [
        ("home-desktop.png", "Home page: introduction, role search, and featured opportunities."),
        ("jobs-desktop.png", "Jobs page: search filters and internship listing cards."),
        ("contact-desktop.png", "Contact page: contact information and message form."),
    ]
    for index, (filename, caption) in enumerate(screenshots, start=1):
        if index > 1:
            add_page_break(body)
        add_word_paragraph(body, caption.split(":")[0], "h2", GREEN_MID)
        add_word_image(body, OUTPUT_DIR / filename, f"rId{index}", index, caption)
        add_word_paragraph(body, caption, "small", MUTED)
    section = SubElement(body, word_tag("sectPr"))
    margins = SubElement(section, word_tag("pgMar"))
    margins.set(word_tag("top"), "1080")
    margins.set(word_tag("right"), "1080")
    margins.set(word_tag("bottom"), "1080")
    margins.set(word_tag("left"), "1080")

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'''
    relationships = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
    image_relationships = Element("Relationships", xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
    for index, (filename, _) in enumerate(screenshots, start=1):
        SubElement(image_relationships, "Relationship", {
            "Id": f"rId{index}",
            "Type": f"{R}/image",
            "Target": f"media/{filename}",
        })
    with ZipFile(OUTPUT_DIR / "DG_Interns_Hub_Project_Report.docx", "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", relationships)
        archive.writestr("word/document.xml", tostring(document, encoding="utf-8", xml_declaration=True))
        archive.writestr("word/_rels/document.xml.rels", tostring(image_relationships, encoding="utf-8", xml_declaration=True))
        for filename, _ in screenshots:
            archive.write(OUTPUT_DIR / filename, f"word/media/{filename}")


def rgb(hex_value):
    return RGBColor.from_string(hex_value)


def add_shape(slide, kind, x, y, w, h, fill, line=None):
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.fill.background() if line is None else None
    if line:
        shape.line.color.rgb = rgb(line)
    return shape


def add_text(slide, text, x, y, w, h, size=18, color=INK, bold=False, font="Aptos", align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.02)
    frame.margin_right = Inches(0.02)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    for index, line in enumerate(text.split("\n")):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = align
        paragraph.space_after = Pt(5)
        run = paragraph.add_run()
        run.text = line
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = rgb(color)
    return box


def add_footer(slide, page, dark=False):
    color = "D5E0D1" if dark else MUTED
    add_text(slide, "DG INTERNS HUB", 0.55, 7.08, 2.6, 0.2, 8, color, True)
    add_text(slide, f"WEEK 07  /  {page:02d}", 10.85, 7.08, 1.9, 0.2, 8, color, True, align=PP_ALIGN.RIGHT)


def add_heading(slide, eyebrow, title, page, dark=False):
    add_text(slide, eyebrow.upper(), 0.68, 0.42, 7.8, 0.25, 9, LIME if dark else GREEN_MID, True)
    add_text(slide, title, 0.68, 0.77, 11.7, 0.65, 27, WHITE if dark else INK, True, font="Aptos Display")
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.68, 1.55, 11.95, 0.015, "46665A" if dark else LINE)
    add_footer(slide, page, dark)


def card(slide, x, y, w, h, title, body, accent=GREEN_MID, dark=False):
    base = "21493D" if dark else WHITE
    text_color = WHITE if dark else INK
    body_color = "D5E0D1" if dark else MUTED
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, base, "46665A" if dark else LINE)
    add_shape(slide, MSO_SHAPE.RECTANGLE, x, y, 0.07, h, accent)
    add_text(slide, title, x + 0.25, y + 0.18, w - 0.45, 0.35, 15, text_color, True)
    add_text(slide, body, x + 0.25, y + 0.62, w - 0.45, h - 0.75, 11, body_color)


def make_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    slide = prs.slides.add_slide(blank)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(GREEN)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 8.5, 0, 4.833, 7.5, GREEN_MID)
    add_shape(slide, MSO_SHAPE.OVAL, 9.15, 1.0, 3.2, 3.2, LIME)
    add_text(slide, "DG", 9.68, 1.45, 2.1, 0.8, 45, GREEN, True, font="Aptos Display", align=PP_ALIGN.CENTER)
    add_text(slide, "INTERNS\nHUB", 9.68, 2.18, 2.1, 0.8, 15, GREEN, True, align=PP_ALIGN.CENTER)
    add_text(slide, "WEEK 07  /  REACT JS PROJECT", 0.85, 0.72, 6.5, 0.3, 11, LIME, True)
    add_text(slide, "Your next chapter\nstarts here.", 0.85, 1.7, 7.2, 2.1, 36, WHITE, True, font="Aptos Display")
    add_text(slide, "A responsive internship discovery website", 0.9, 4.1, 6.5, 0.5, 18, "D5E0D1")
    add_shape(slide, MSO_SHAPE.RECTANGLE, 0.9, 5.2, 1.2, 0.08, CORAL)
    add_text(slide, "PROJECT REPORT + PRESENTATION", 0.9, 5.55, 5.8, 0.3, 10, "D5E0D1", True)
    add_footer(slide, 1, True)

    slide = prs.slides.add_slide(blank)
    add_heading(slide, "01 / Purpose", "A first opportunity should be easier to find.", 2)
    add_text(slide, "THE NEED", 0.8, 2.0, 2.0, 0.3, 10, CORAL, True)
    add_text(slide, "Students need a clear starting point to explore internships by skill, field, and work arrangement.", 0.8, 2.45, 4.6, 1.45, 23, INK, True, font="Aptos Display")
    add_shape(slide, MSO_SHAPE.RECTANGLE, 6.0, 2.05, 0.025, 3.55, LINE)
    card(slide, 6.6, 2.0, 2.65, 1.65, "Discover", "Browse a curated set of early-career roles.", GREEN_MID)
    card(slide, 9.55, 2.0, 2.65, 1.65, "Narrow", "Search and filter without losing context.", CORAL)
    card(slide, 6.6, 4.0, 2.65, 1.65, "Take action", "Start an application through the role’s email link.", GREEN_MID)
    card(slide, 9.55, 4.0, 2.65, 1.65, "Stay connected", "Reach the hub from the contact page.", CORAL)

    slide = prs.slides.add_slide(blank)
    add_heading(slide, "02 / Solution", "Three pages. One connected experience.", 3)
    steps = [("01", "Home", "Search entry point\nFeatured openings"), ("02", "Jobs", "Six sample roles\nSearch + filters"), ("03", "Contact", "Questions +\ncontact form")]
    for index, (number, title, copy) in enumerate(steps):
        x = 0.85 + index * 4.05
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.15, 3.55, 2.8, WHITE, LINE)
        add_text(slide, number, x + 0.25, 2.38, 0.65, 0.45, 18, CORAL, True)
        add_text(slide, title, x + 0.25, 3.0, 3.0, 0.5, 22, GREEN, True, font="Aptos Display")
        add_text(slide, copy, x + 0.25, 3.7, 2.9, 0.9, 14, MUTED)
        if index < 2:
            add_text(slide, ">", x + 3.62, 3.25, 0.35, 0.5, 24, GREEN_MID, True, align=PP_ALIGN.CENTER)
    add_text(slide, "Shared navigation keeps the journey predictable across routes.", 0.9, 5.65, 9.0, 0.4, 14, MUTED)

    slide = prs.slides.add_slide(blank)
    add_heading(slide, "03 / Home", "A welcoming start with a useful next step.", 4)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.85, 1.95, 7.2, 4.55, GREEN)
    add_text(slide, "GOOD WORK STARTS SOMEWHERE", 1.2, 2.35, 4.8, 0.3, 9, LIME, True)
    add_text(slide, "Your next chapter\nstarts here.", 1.2, 2.85, 5.7, 1.35, 30, WHITE, True, font="Aptos Display")
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 1.2, 4.55, 5.8, 0.62, WHITE)
    add_text(slide, "Role, skill, or company", 1.45, 4.7, 3.6, 0.25, 10, MUTED)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 5.55, 4.63, 1.27, 0.45, LIME)
    add_text(slide, "Find a role", 5.62, 4.74, 1.12, 0.18, 9, GREEN, True, align=PP_ALIGN.CENTER)
    add_text(slide, "Open doors, right now.", 1.2, 5.65, 5.5, 0.4, 16, WHITE, True)
    card(slide, 8.55, 2.05, 3.55, 1.25, "Search first", "Home search carries its query into Jobs.", GREEN_MID)
    card(slide, 8.55, 3.65, 3.55, 1.25, "See a sample", "Featured role cards preview the catalogue.", CORAL)
    card(slide, 8.55, 5.25, 3.55, 1.0, "Move forward", "A clear route to all opportunities.", GREEN_MID)

    slide = prs.slides.add_slide(blank)
    add_heading(slide, "04 / Jobs", "Help people find the right fit faster.", 5)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.8, 1.95, 11.75, 0.78, WHITE, LINE)
    add_text(slide, "⌕   Title, skill, or company", 1.1, 2.18, 3.2, 0.28, 12, MUTED)
    add_text(slide, "All fields  ▾", 5.0, 2.18, 1.65, 0.28, 11, INK)
    add_text(slide, "Any location  ▾", 7.2, 2.18, 2.0, 0.28, 11, INK)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 10.55, 2.1, 1.55, 0.45, GREEN)
    add_text(slide, "Search  →", 10.7, 2.22, 1.2, 0.2, 10, WHITE, True, align=PP_ALIGN.CENTER)
    roles = [("N", "Frontend Developer Intern", "Remote", "₹18,000 / month"), ("C", "Product Design Intern", "Hybrid", "₹15,000 / month"), ("G", "Digital Marketing Intern", "On-site", "₹12,000 / month")]
    for index, (mark, role, mode, pay) in enumerate(roles):
        x = 0.8 + index * 4.0
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, 3.15, 3.72, 2.65, WHITE, LINE)
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x + 0.25, 3.42, 0.48, 0.48, ["CBE5D3", "F6D6C6", "C9DFE8"][index])
        add_text(slide, mark, x + 0.38, 3.53, 0.2, 0.2, 14, INK, True)
        add_text(slide, role, x + 0.25, 4.1, 3.2, 0.56, 15, INK, True)
        add_text(slide, f"{mode}  ·  {pay}", x + 0.25, 4.87, 3.1, 0.3, 10, MUTED)
        add_text(slide, "APPLY  ↗", x + 0.25, 5.35, 1.2, 0.2, 9, GREEN_MID, True)
    add_text(slide, "Live filters: title / company / category / skills + field + work mode", 0.9, 6.15, 9.5, 0.3, 12, MUTED)

    slide = prs.slides.add_slide(blank)
    add_heading(slide, "05 / Design", "A calm interface with a bright signal.", 6)
    swatches = [(GREEN, "Deep green", "#173D32"), (LIME, "Fresh lime", "#D4F270"), (CORAL, "Warm coral", "#E87952"), (PAPER, "Paper", "#FBFCF8")]
    for index, (color, label, hex_code) in enumerate(swatches):
        x = 0.85 + index * 1.85
        add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.0, 1.55, 1.1, color, LINE if color == PAPER else None)
        add_text(slide, label, x, 3.22, 1.55, 0.25, 10, INK, True)
        add_text(slide, hex_code, x, 3.52, 1.55, 0.25, 9, MUTED)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 8.85, 1.9, 3.0, 4.2, "EDF3E9")
    add_text(slide, "MOBILE READY", 9.18, 2.15, 2.3, 0.3, 9, GREEN_MID, True, align=PP_ALIGN.CENTER)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 9.62, 2.72, 1.48, 2.8, WHITE, LINE)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 9.78, 3.05, 1.16, 0.6, GREEN)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 9.78, 3.82, 1.16, 0.23, LINE)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 9.78, 4.16, 1.16, 0.23, LINE)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 9.78, 4.58, 1.16, 0.55, LIME)
    add_text(slide, "Responsive layout\nVisible focus states\nSemantic navigation\nExternal photo + fonts", 0.9, 4.45, 5.9, 1.45, 15, INK)

    slide = prs.slides.add_slide(blank)
    add_heading(slide, "06 / Implementation", "Small React pieces, clear ownership.", 7)
    layers = [("UI", "Header / Footer\nJobCard\nHome / Jobs / Contact", GREEN_MID), ("STATE", "Search query\nCategory + mode\nContact confirmation", CORAL), ("DATA", "Local opportunities[]\nSix illustrative entries\nReplace with real feed", GREEN_MID)]
    for index, (label, content, accent) in enumerate(layers):
        x = 0.9 + index * 3.5
        card(slide, x, 2.25, 2.95, 2.2, label, content, accent)
        if index < 2:
            add_text(slide, "→", x + 3.03, 3.02, 0.4, 0.45, 23, GREEN_MID, True, align=PP_ALIGN.CENTER)
    card(slide, 1.15, 5.15, 3.2, 0.95, "React Router", "BrowserRouter + 3 routes", GREEN_MID)
    card(slide, 5.05, 5.15, 3.2, 0.95, "Vite", "Dev server + production build", CORAL)
    card(slide, 8.95, 5.15, 3.2, 0.95, "CSS + Lucide", "Responsive styles + icons", GREEN_MID)

    slide = prs.slides.add_slide(blank)
    add_heading(slide, "07 / Quality", "A working front end, with clear boundaries.", 8)
    card(slide, 0.85, 2.0, 3.55, 1.65, "Verified earlier", "Production build completed. Oxlint completed. Editor diagnostics clean after source edits.", GREEN_MID)
    card(slide, 4.85, 2.0, 3.55, 1.65, "Current boundary", "Contact form confirms locally only. Apply links use sample email addresses. Listings are sample data.", CORAL)
    card(slide, 8.85, 2.0, 3.55, 1.65, "Before launch", "Run a final build. Check mobile in a browser. Confirm contact routes, real listings, and delivery flow.", GREEN_MID)
    add_text(slide, "No API, database, account system, or message service is included in this assignment build.", 0.9, 4.35, 10.7, 0.55, 17, INK, True)
    add_text(slide, "External image and font requests require an internet connection.", 0.9, 5.1, 8.4, 0.35, 12, MUTED)

    slide = prs.slides.add_slide(blank)
    add_heading(slide, "08 / Next steps", "Turn the demo into a live opportunity hub.", 9)
    actions = [
        ("01", "Replace sample roles", "Use verified listings, application links, and clear expiry dates."),
        ("02", "Connect submissions", "Send contact messages to a real inbox or service; add success and error states."),
        ("03", "Publish", "Push to GitHub, deploy dist on Netlify/Vercel, configure SPA fallback."),
        ("04", "Share", "Add the live URL to the report and post; tag DG Interns Hub on LinkedIn."),
    ]
    for index, (number, title, body) in enumerate(actions):
        y = 1.95 + index * 1.12
        add_shape(slide, MSO_SHAPE.OVAL, 0.95, y, 0.55, 0.55, LIME)
        add_text(slide, number, 1.05, y + 0.13, 0.35, 0.2, 10, GREEN, True, align=PP_ALIGN.CENTER)
        add_text(slide, title, 1.85, y - 0.01, 3.0, 0.3, 15, INK, True)
        add_text(slide, body, 5.0, y - 0.02, 6.5, 0.5, 12, MUTED)
    add_shape(slide, MSO_SHAPE.RECTANGLE, 1.22, 2.52, 0.02, 2.92, LINE)

    slide = prs.slides.add_slide(blank)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(GREEN)
    add_text(slide, "DG INTERNS HUB  /  REACT JS", 0.9, 0.82, 6.8, 0.3, 10, LIME, True)
    add_text(slide, "Good work starts\nsomewhere.", 0.9, 1.75, 8.0, 1.8, 37, WHITE, True, font="Aptos Display")
    add_text(slide, "A connected, responsive place to discover a first opportunity.", 0.95, 4.0, 7.2, 0.6, 17, "D5E0D1")
    add_shape(slide, MSO_SHAPE.OVAL, 9.2, 1.55, 2.65, 2.65, CORAL)
    add_text(slide, "NEXT\nCHAPTER", 9.57, 2.3, 1.9, 0.7, 17, WHITE, True, align=PP_ALIGN.CENTER)
    add_text(slide, "Thank you", 0.95, 5.85, 4.0, 0.6, 22, LIME, True)
    add_footer(slide, 10, True)
    prs.save(OUTPUT_DIR / "DG_Interns_Hub_Presentation.pptx")


if __name__ == "__main__":
    OUTPUT_DIR.mkdir(exist_ok=True)
    make_report()
    make_presentation()
    print(f"Created files in {OUTPUT_DIR}")