import pandas as pd
import datetime as dt
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 1) authorize with Google Sheets
scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]
creds  = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# 2) open sheet
sheet = client.open('Cotopaxi Media Report - 2025').worksheet('ONLINE')
values = sheet.get_all_values()
rows   = values[1:]  # skip header

# 3) load into DataFrame (positional columns)
df = pd.DataFrame(rows)

# 4) parse dates (column E → index 4)
df[4] = pd.to_datetime(df[4], errors='coerce')

# 5) row-by-row Impressions parser
def parse_impr(row):
    for idx in (5, 6):        # F or G
        val = row[idx].replace(',', '').strip()
        if val.isdigit():
            return int(val)
    return 0

df['Impressions'] = df.apply(parse_impr, axis=1)

# 6) categorize (I/J → idx 8/9)
def get_category(r):
    if r[8]=='P' and r[9]=='A': return 'APPAREL FEATURES'
    if r[8]=='P' and r[9]=='P': return 'PACKS & BAG FEATURES'
    if r[8]=='B':               return 'BRAND/SUSTAINABILITY'
    return None

df['Category'] = df.apply(get_category, axis=1)

# 7) define Thu→Wed window
today        = dt.date.today()
delta_to_wed = (2 - today.weekday()) % 7  # Wed=2
end_date     = today + dt.timedelta(days=delta_to_wed)
start_date   = end_date - dt.timedelta(days=6)

mask    = df[4].dt.date.between(start_date, end_date)
week_df = df.loc[mask].sort_values([ 'Category', 4 ], ascending=[True, True])

# 8) totals
total_impr = week_df['Impressions'].sum()
total_hits = len(week_df)

# 9) build Word doc
def add_hyperlink(p, url, text):
    part = p.part
    r_id = part.relate_to(
        url,
        reltype="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )
    link = OxmlElement('w:hyperlink'); link.set(qn('r:id'), r_id)
    r    = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    c    = OxmlElement('w:color'); c.set(qn('w:val'),"0000FF"); rPr.append(c)
    u    = OxmlElement('w:u'); u.set(qn('w:val'), "single"); rPr.append(u)
    r.append(rPr)
    t    = OxmlElement('w:t'); t.text = text
    r.append(t); link.append(r); p._p.append(link)

doc = Document()
doc.styles['Normal'].font.name = 'Calibri'; doc.styles['Normal'].font.size = Pt(11)

# Header
h = doc.add_paragraph()
h.add_run(f"Date Range: {start_date:%B %d, %Y} - {end_date:%B %d, %Y}").bold = True
doc.add_paragraph(f"Total Impressions: {total_impr:,}")
doc.add_paragraph(f"Total Hits: {total_hits}")
doc.add_paragraph("")

# Sections
for sec in ['APPAREL FEATURES','PACKS & BAG FEATURES','BRAND/SUSTAINABILITY']:
    sub = week_df[week_df['Category']==sec]
    if sub.empty: continue
    p = doc.add_paragraph(); p.add_run(sec).bold = True
    for _, r in sub.iterrows():
        date_str = r[4].strftime('%m/%d/%Y')
        outlet   = r[0]
        headline = r[3]
        url      = r[7]
        imp      = r['Impressions']
        para = doc.add_paragraph()
        para.add_run(f"{date_str} – {outlet}: ")
        add_hyperlink(para, url, headline)
        para.add_run(f" - {imp:,} Impressions")

# Save
output = 'cotopaxi_weekly_hits.docx'
doc.save(output)
print("Generated:", output)

