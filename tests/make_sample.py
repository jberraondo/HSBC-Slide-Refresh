"""Builds tests/sample.pptx: a made-up deck (no client data) that covers every
slide type the tool recognises. Run with:  python3 tests/make_sample.py
Needs: pip install python-pptx pillow
"""
import io
import random
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

HERE = Path(__file__).parent
random.seed(4)


def photo(w=1600, h=1000, hue=(40, 110, 60)):
    """A soft gradient 'photo' so image handling can be tested."""
    im = Image.new("RGB", (w, h))
    d = ImageDraw.Draw(im)
    for y in range(h):
        t = y / h
        d.line([(0, y), (w, y)], fill=(int(hue[0] + 120 * t), int(hue[1] + 80 * (1 - t)), int(hue[2] + 60 * t)))
    for _ in range(12):
        x, y, r = random.randint(0, w), random.randint(0, h), random.randint(40, 160)
        d.ellipse([x - r, y - r, x + r, y + r], fill=(230, 230, 230))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=85)
    buf.seek(0)
    return buf


def text(slide, x, y, w, h, lines, size=12, bold_first=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        for r in p.runs:
            r.font.size = Pt(size)
            r.font.bold = bold_first and i == 0
    return tb


prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
blank, title_only = prs.slide_layouts[6], prs.slide_layouts[5]

# 1 cover
s = prs.slides.add_slide(prs.slide_layouts[0])
s.shapes.title.text = "Sample Sport Report"
s.placeholders[1].text = "Q1 Report 2027"

# 2 contents
s = prs.slides.add_slide(blank)
text(s, 0.5, 0.4, 5, 0.6, ["Contents"], 28)
text(s, 0.5, 1.3, 5, 4, ["Headline results", "Top content", "Year on year", "Partnerships", "Learnings"], 16)
s.shapes.add_picture(photo(), Inches(6.5), Inches(0), Inches(6.8), Inches(7.5))

# 3 divider
s = prs.slides.add_slide(blank)
s.shapes.add_picture(photo(hue=(30, 60, 120)), Inches(3.5), Inches(1), Inches(9.8), Inches(6.5))
text(s, 0.6, 2.6, 5, 0.8, ["Headline Results"], 32)

# 4 headline figures
s = prs.slides.add_slide(title_only)
s.shapes.title.text = "Sample Sport – Overall Results"
text(s, 0.4, 1.4, 4, 0.8, ["40 Posts Q1 '27", "52 Posts Q1 '26"], 10)
text(s, 0.4, 2.3, 4, 3, ["Reach grew on the back of one big moment, which delivered most of the quarter's impressions.",
                          "Fewer posts went out, but each worked harder than last year."], 11)
for i, (num, label, c1, c2) in enumerate([
        ("12.4M", "Impressions", "+40% on Q1 2026 (8.9M)", "+12% on Q1 2025 (11.1M)"),
        ("5.1M", "Views", "+22% on Q1 2026 (4.2M)", "+60% on Q1 2025 (3.2M)"),
        ("410K", "Engagements", "+8% on Q1 2026 (380K)", "-5% on Q1 2025 (431K)"),
        ("3.3%", "Engagement Rate", "-23% on Q1 2026 (4.3%)", "-15% on Q1 2025 (3.9%)")]):
    text(s, 5 + (i % 2) * 4, 1.5 + (i // 2) * 2.4, 3.8, 2, [f"{num} {label}", c1, c2], 14)

# 5 content cards
s = prs.slides.add_slide(title_only)
s.shapes.title.text = "Sample Sport – top performing content"
for i, (name, imp, views) in enumerate([("Range session film", "2,401,220", "1,204,118"),
                                         ("Caddie story", "1,880,410", "990,320"),
                                         ("Fan questions", "640,200", "402,900")]):
    s.shapes.add_picture(photo(hue=(60 + 40 * i, 100, 70)), Inches(0.6 + i * 4.2), Inches(1.5), Inches(3), Inches(3))
    text(s, 0.6 + i * 4.2, 4.7, 3.6, 2.6, [f"Tournament – {name} (Instagram)", "Short description of the post and why it worked.",
                                          "Key Results", f"Impressions: {imp}", f"Views: {views}", "Engagements: 41,200", "Engagement Rate: 3.4%"], 10, True)

# 6 line charts
s = prs.slides.add_slide(title_only)
s.shapes.title.text = "Sample Sport – how we're tracking YOY"
months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
for i, metric in enumerate(["Impressions", "Views", "Engagements"]):
    cd = CategoryChartData()
    cd.categories = months
    base = [10, 14, 30, 33, 35, 42, 70, 72, 79, 82, 98, 115]
    cd.add_series("2026", [v * (i + 1) * 1e5 for v in base])
    cd.add_series("2027", [v * (i + 1) * 1.6e5 for v in base[:3]] + [None] * 9)
    gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(0.4 + i * 4.3), Inches(1.6), Inches(4.2), Inches(3.5), cd)
    gf.chart.has_title = True
    gf.chart.chart_title.text_frame.text = f"{metric} 26 vs 27"

# 7 partnership bars
s = prs.slides.add_slide(title_only)
s.shapes.title.text = "Sample Sport – partnership comparison"
text(s, 0.4, 1.4, 3.6, 1, ["Golf – 30 posts, £20,000", "Tennis – 6 posts, £2,000", "Rugby – 4 posts, no paid media"], 10)
text(s, 0.4, 2.6, 3.6, 3, ["Golf carried the quarter, with most of the reach and engagement."], 11)
for i, (metric, vals) in enumerate([("Impressions Per Partnership", (9.8e6, 1.9e6, 0.7e6)), ("Views Per Partnership", (4.1e6, 0.8e6, 0.2e6)),
                                    ("Engagements Per Partnership", (330e3, 52e3, 28e3)), ("Engagement Rate Per Partnership", (3.4, 2.7, 4.0))]):
    cd = CategoryChartData()
    cd.categories = ["GOLF", "TENNIS", "RUGBY"]
    cd.add_series("Value", vals)
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(4.3 + (i % 2) * 4.5), Inches(1.4 + (i // 2) * 2.9), Inches(4.3), Inches(2.7), cd)
    gf.chart.has_title = True
    gf.chart.chart_title.text_frame.text = metric

# 8 donut
s = prs.slides.add_slide(title_only)
s.shapes.title.text = "Sample Sport – Content Audit"
text(s, 0.4, 1.5, 6.5, 5, ["The Blockbuster delivered:", "One piece made up a small share of posts but most of the reach.",
                           "Outside Framework fell:", "Only a handful of posts sat outside the framework."], 12)
cd = CategoryChartData()
cd.categories = ["Blockbuster", "Partnership Storytelling", "Insider Pulse", "Outside Framework"]
cd.add_series("Number of Posts", (3, 12, 22, 3))
gf = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(7.4), Inches(1.4), Inches(5.5), Inches(5.5), cd)

# 9 learnings
s = prs.slides.add_slide(title_only)
s.shapes.title.text = "Key Learnings"
text(s, 0.5, 1.5, 12, 5, ["Planning around the next event pays off",
                          "Filming ahead gave the team a finished hero piece before the tournament started, and it became one of the most viewed posts.",
                          "Collabs should be the default for organic ideas",
                          "Posts published with a partner reached far more people than those posted alone, at no extra cost.",
                          "Fewer, better posts work",
                          "Cutting volume did not hurt the totals, and each post now works harder."], 12)

# 10 table
s = prs.slides.add_slide(title_only)
s.shapes.title.text = "Sample Sport – Demographics"
text(s, 0.4, 1.4, 3.6, 4, ["Hong Kong performed above average thanks to new targeting."], 11)
rows = [["Market", "% of Views"], ["Hong Kong", "18%"], ["UK", "15%"], ["India", "12%"], ["UAE", "8%"]]
tbl = s.shapes.add_table(len(rows), 2, Inches(4.5), Inches(1.5), Inches(4), Inches(2.5)).table
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        tbl.cell(r, c).text = val

prs.save(HERE / "sample.pptx")
print("wrote", HERE / "sample.pptx")
