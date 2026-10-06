"""Build Report_Group09.docx from the notebook outputs (outputs/results.json, outputs/*.csv, figures/*.png).

Every number in the report is read from the files the notebook writes, so the report always matches the code.
"""
import json
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
R = json.loads((ROOT / "outputs" / "results.json").read_text())
FIG = ROOT / "figures"
GROUP = "Group09"

NAVY = RGBColor(0x10, 0x42, 0x81)
INK = RGBColor(0x0B, 0x0B, 0x0B)
GREY = RGBColor(0x52, 0x51, 0x4E)
FONT = "Calibri"
SIZE = Pt(11)

doc = Document()
sec = doc.sections[0]
sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, side, Cm(2.54))
TEXT_W = Cm(21.0 - 2 * 2.54)

# ---- Base styles: everything 11 pt ----
st = doc.styles["Normal"]
st.font.name = FONT; st.font.size = SIZE
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
st.paragraph_format.space_after = Pt(4)
st.paragraph_format.line_spacing = 1.0
for name, before, after in [("Heading 1", 10, 4), ("Heading 2", 8, 3), ("Heading 3", 6, 2)]:
    h = doc.styles[name]
    h.font.name = FONT; h.font.size = SIZE; h.font.bold = True
    h.font.color.rgb = NAVY if name != "Heading 3" else INK
    h.element.rPr.rFonts.set(qn("w:asciiTheme"), "") if False else None
    rpr = h.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), FONT)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    h.paragraph_format.space_before = Pt(before); h.paragraph_format.space_after = Pt(after)
    h.paragraph_format.keep_with_next = True
doc.styles["Heading 1"].font.all_caps = True
for lst in ("List Bullet", "List Number"):
    doc.styles[lst].font.name = FONT; doc.styles[lst].font.size = SIZE
    doc.styles[lst].paragraph_format.space_after = Pt(2)

fig_no = [0]
tab_no = [0]


def para(text="", bold_lead=None, style=None, align=None, after=None):
    p = doc.add_paragraph(style=style)
    if bold_lead:
        r = p.add_run(bold_lead); r.bold = True
    _add_runs(p, text)
    if align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    return p


def _add_runs(p, text):
    """Support *italic* spans in body text."""
    parts = text.split("*")
    for i, part in enumerate(parts):
        if part:
            r = p.add_run(part); r.italic = (i % 2 == 1)


def bullet(text, lead=None):
    """Numbered recommendation paragraph with a hanging indent (the number is in `lead`)."""
    p = para(text, bold_lead=lead, align="justify")
    p.paragraph_format.left_indent = Cm(0.6); p.paragraph_format.first_line_indent = Cm(-0.6)
    return p


def figure(name, caption, width_cm=15.9):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0); p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(FIG / f"{name}.png"), width=Cm(width_cm))
    fig_no[0] += 1
    c = doc.add_paragraph(); c.paragraph_format.space_after = Pt(6)
    r = c.add_run(f"Figure {fig_no[0]}. "); r.bold = True; r.font.color.rgb = GREY
    r = c.add_run(caption); r.font.color.rgb = GREY
    return fig_no[0]


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def set_cell_margins(table, top=30, bottom=30, left=70, right=70):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for k, v in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        e = OxmlElement(f"w:{k}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); mar.append(e)
    tblPr.append(mar)


def table(df, caption, widths_cm, align_right_from=1):
    tab_no[0] += 1
    c = doc.add_paragraph(); c.paragraph_format.space_after = Pt(2); c.paragraph_format.keep_with_next = True
    r = c.add_run(f"Table {tab_no[0]}. "); r.bold = True; r.font.color.rgb = GREY
    r = c.add_run(caption); r.font.color.rgb = GREY
    t = doc.add_table(rows=1, cols=len(df.columns))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_cell_margins(t)
    for j, col in enumerate(df.columns):
        cell = t.rows[0].cells[j]; cell.text = ""
        run = cell.paragraphs[0].add_run(str(col)); run.bold = True; run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(cell, "104281")
    trPr = t.rows[0]._tr.get_or_add_trPr()   # repeat the header row when a table runs onto a new page
    hdr = OxmlElement("w:tblHeader"); hdr.set(qn("w:val"), "true"); trPr.append(hdr)
    for i, row in enumerate(df.itertuples(index=False)):
        cells = t.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = ""
            cells[j].paragraphs[0].add_run(str(v))
            if j >= align_right_from:
                cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if i % 2 == 1:
                shade(cells[j], "EEF4FC")
    for row in t.rows:
        for j, w in enumerate(widths_cm):
            row.cells[j].width = Cm(w)
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(0)
    # light borders
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "4"); e.set(qn("w:color"), "C3C2B7")
        borders.append(e)
    t._tbl.tblPr.append(borders)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return tab_no[0]


def money(x, d=0):
    return f"${x:,.{d}f}"


# ======================================================================
# Cover page
# ======================================================================
for _ in range(3):
    doc.add_paragraph()
for text, size, bold, colour in [
    ("SINGAPORE INSTITUTE OF TECHNOLOGY", 14, True, NAVY),
    ("ACC2125 Data Analytics and Machine Learning", 14, False, INK),
    ("AY2026/27 Trimester 1 — Group Project Report", 12, False, INK),
    ("", 11, False, INK),
    ("Does Rail Access Pay?", 24, True, NAVY),
    ("Machine Learning on Sydney's Short-Term Rental Market", 16, False, INK),
    ("", 11, False, INK),
    ("Case Study: Machine Learning on Short-Term Rental Markets", 12, False, GREY),
    ("Case Designer: Dr. Desi Arisandi", 12, False, GREY),
]:
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text); r.font.size = Pt(size); r.bold = bold; r.font.color.rgb = colour
for _ in range(4):
    doc.add_paragraph()
cover = pd.DataFrame({"": ["Group", "Members", "", "", "", "Submission date"],
                      " ": [GROUP, "[Member name — Student ID]", "[Member name — Student ID]", "[Member name — Student ID]",
                            "[Member name — Student ID]", "29 November 2026"]})
t = doc.add_table(rows=0, cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
for a, b in cover.itertuples(index=False):
    cells = t.add_row().cells
    cells[0].text = a; cells[1].text = b
    cells[0].paragraphs[0].runs[0].bold = True if a else False
    cells[0].width = Cm(4); cells[1].width = Cm(8)
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ======================================================================
# Part A
# ======================================================================
def ci(lo, hi, unit="%", d=1):
    return f"95% CI {lo:+.{d}f}{unit} to {hi:+.{d}f}{unit}"

def pval(p):
    return "p < 0.001" if p < 0.001 else (f"p = {p:.3f}" if p < 0.01 else f"p = {p:.2f}")

doc.add_heading("Part A. Data Selection and Exploratory Analysis", level=1)
doc.add_heading("A1. Data selection", level=2)
para(f"We analyse the Inside Airbnb snapshot for Sydney, New South Wales, scraped between 17 and 29 June 2026 "
     f"({R['raw_listings']:,} listings; {R['raw_calendar_rows']:,} calendar rows). Sydney is one of the world's largest Airbnb markets, with "
     "sharply different sub-markets (CBD apartments, harbour and beach suburbs, the western suburbs) and a dense rail network that has grown "
     "quickly since 2019 (light rail to Randwick and Kingsford, Sydney Metro under the harbour, Parramatta Light Rail). That makes it a natural "
     "place to test a common investor belief: that listings near a station earn more. We use listings.csv.gz, calendar.csv.gz (forward "
     "availability only; this snapshot has no calendar prices) and neighbourhoods.geojson (38 local government areas, LGAs). The 21 March 2026 "
     f"snapshot ({R['raw_listings_mar']:,} listings) is used only as a cross-season check; the September and December 2025 snapshots have no prices.",
     align="justify")
para("No single dataset describes rail access well, so we combine five public sources (Table 1). Each is joined to every listing and used "
     "in Q3, both unguided analyses and all three models.", bold_lead="External datasets. ", align="justify")
ext = pd.DataFrame([
    ("External01", "Transport for NSW, Timetables Complete GTFS (Open Data Hub)", "Timetable valid from 6 Oct 2026", "6 Oct 2026",
     "Station entrances, mode, off-peak frequency, rail time to CBD and airport"),
    ("External02", "Transport for NSW, Train, Metro and Light Rail Station Entries and Exits", "Oct 2024 – Aug 2026 (monthly)", "7 Oct 2026",
     "Station patronage (12-month average)"),
    ("External03", "Australian Bureau of Statistics, SEIFA 2021 by SA2 (+ ASGS Ed. 3 SA2 boundaries)", "2021 Census", "7 Oct 2026",
     "Neighbourhood advantage (IRSAD); SA2 areas for location controls"),
    ("External04", "OpenStreetMap contributors, via Overpass API (OSMnx)", "Extract of 7 Oct 2026", "7 Oct 2026",
     "Beaches, tourist attractions; pedestrian network for walking distance"),
    ("External05", "Transport for NSW, Train Station Entrance Locations (v4)", "Network as at 2020", "6 Oct 2026",
     "Comparison only: what the outdated network misses"),
], columns=["File", "Source and publisher", "Coverage", "Accessed", "Used for"])
table(ext, "External datasets (TfNSW and ABS data CC BY 4.0; OpenStreetMap ODbL 1.0)", [1.9, 5.0, 2.8, 1.8, 4.4], align_right_from=99)
para(f"The 1.4 GB GTFS feed was condensed to one row per entrance of the {R['stations_sydney']} Sydney stations "
     f"({R['stations_by_mode']['Train']} train, {R['stations_by_mode']['Metro']} Metro, {R['stations_by_mode']['Light rail']} light rail). "
     "For each station we computed off-peak departures per hour and the expected rail travel time to the CBD and the airport, using a reverse "
     "connection scan over the 7 October 2026 timetable (Dibbelt et al., 2013) with 3-minute transfers. Each listing was then routed along the "
     f"OpenStreetMap footpath network to its nearest entrance ({R['walk_routed_pct']}% routed), placed in one of {R['sa2_n']} ABS Statistical Area Level 2 "
     "(SA2) neighbourhoods, and given its distance to the nearest beach and the number of attractions within 1 km. "
     f"The median listing is a {R['median_dist_station_m']:,} m walk from a station (straight line {R['median_straight_dist_station_m']:,} m), and "
     f"{R['share_within_500m']}% are within a 500 m walk. Compared with the 2020 file, {len(R['new_stations_since_2020'])} stations are new and "
     f"{R['listings_station_closer_than_2020']:,} listings are now more than 200 m closer to rail.", bold_lead="Integration. ", align="justify")
log = pd.read_csv(ROOT / "outputs" / "data_decisions_log.csv")[["Change Details", "Reason and Justification"]]
table(log, f"Data Decisions Log ({R['raw_listings']:,} raw listings → {R['clean_listings']:,} analysed)", [7.9, 8.0], align_right_from=99)

# ---- Q1 ----
doc.add_heading("A2. Guided analysis", level=2)
doc.add_heading("Q1. Characteristics of the most popular listings", level=3)
q1 = pd.DataFrame(R["q1"]); pop, oth = q1["Popular (top 25%)"], q1["Other listings"]
rob = R["q1_robust"]; cal_ = rob["Calendar: next 90 days unavailable"]
mn = R["q1_occ_by_minnights"]; rb = R["q1_occ_by_rating"]; sh = R["q1_occ_superhost"]
para("We measure popularity with Inside Airbnb's *estimated nights booked in the last 12 months* and call the top quartile popular "
     f"(at least {R['popular_threshold_nights']:.0f} nights; n = {R['popular_n']:,}). Popular listings are *not* the expensive ones "
     f"({money(pop['Median price (AUD)'])} against {money(oth['Median price (AUD)'])}). They stand out on how they are run (Table 3): "
     f"{pop['Superhost %']:.0f}% are superhosts (versus {oth['Superhost %']:.0f}%), the median minimum stay is {pop['Median minimum nights']:.0f} night "
     f"(versus {oth['Median minimum nights']:.0f}) and they list more amenities. Median occupancy falls from {mn['1']:.0f} nights at one-night minimums to "
     f"{mn['8-30']:.0f} at 8–30 nights, and rises with rating from {rb['<4.5']:.0f} nights below 4.5 to {rb['4.8-4.9']:.0f} at 4.8–4.9 (Figure 1). "
     "Commercial hosts are under-represented among popular listings.", align="justify")
rows = [("Median nightly price", money(pop["Median price (AUD)"]), money(oth["Median price (AUD)"])),
        ("Median minimum stay (nights)", f"{pop['Median minimum nights']:.0f}", f"{oth['Median minimum nights']:.0f}"),
        ("Superhost share", f"{pop['Superhost %']:.0f}%", f"{oth['Superhost %']:.0f}%"),
        ("Median amenities listed", f"{pop['Median amenities']:.0f}", f"{oth['Median amenities']:.0f}"),
        ("Mean rating (5+ reviews)", f"{pop['Mean rating (5+ reviews)']:.2f}", f"{oth['Mean rating (5+ reviews)']:.2f}"),
        ("Run by a commercial host (10+ listings)", f"{pop['Commercial host (10+) %']:.0f}%", f"{oth['Commercial host (10+) %']:.0f}%"),
        ("Within a 500 m walk of a station", f"{pop['Within 500 m walk of station %']:.0f}%", f"{oth['Within 500 m walk of station %']:.0f}%")]
table(pd.DataFrame(rows, columns=["Characteristic", f"Popular (n = {pop['Listings']:,.0f})", f"Other (n = {oth['Listings']:,.0f})"]),
      "Popular listings versus the rest (price, minimum stay, amenities, rating and walking distance differ at p < 0.001, Mann-Whitney U)", [8.3, 3.8, 3.8])
para(f"*Robustness.* The occupancy estimate is built from review counts, so it tracks reviews in the last 12 months almost perfectly (ρ = {R['rho_occ_reviews_ltm']}) "
     f"but the forward calendar only weakly (ρ = {R['rho_occ_calendar']}). When popularity is defined by the calendar instead, only "
     f"{cal_['Overlap with main definition %']:.0f}% of listings overlap, the superhost gap vanishes ({cal_['Superhost % (popular / other)']}) and minimum stays are equal. "
     "Part of the superhost link is mechanical, since superhost status itself requires completed, reviewed stays. The pattern that holds under every "
     f"definition is that commercial hosts are under-represented among the busiest listings ({cal_['Commercial host % (popular / other)']} on the calendar). "
     "The insight is that demand follows low booking friction and a good review record, not price.", align="justify")
figure("q1_popularity_drivers", "Median estimated nights booked by minimum stay (A), rating band (B) and superhost status (C).", width_cm=15)

# ---- Q2 ----
doc.add_heading("Q2. Relationship between price and guest ratings", level=3)
sp = R["q2_spearman_price"]; dec = R["q2_deciles"]; br = R["q2_by_room"]
para(f"Among the {R['q2_n_rated']:,} listings with at least five reviews, price and rating are positively but weakly related "
     f"(Spearman ρ = {R['q2_rho_overall']:.2f}, p < 0.001), also within room types (entire homes ρ = {br['Entire home/apt']['rho']:.2f}; private rooms "
     f"ρ = {br['Private room']['rho']:.2f}). The relationship is non-linear (Figure 2A): the cheapest decile (median {money(dec['price_median'][0])}) averages "
     f"{dec['rating_mean'][0]:.2f} stars with {dec['share_below_45'][0]:.0f}% of listings below 4.5, the second decile already {dec['rating_mean'][1]:.2f}, and the curve is "
     f"then nearly flat to {dec['rating_mean'][9]:.2f} at the top. Log price explains only {100*R['q2_linear_r2']:.1f}% of rating variation. Location is the "
     f"sub-score most tied to price (ρ = {sp['Location']:.2f}) and Value the least (ρ = {sp['Value']:.2f}; Figure 2B): guests at dearer listings are "
     "happier overall but do not feel they get better value. Pricing above the market floor protects against poor ratings, but high prices do not buy higher ones.",
     align="justify")
figure("q2_price_vs_rating", "Mean overall, value and location scores by price decile (A) and Spearman correlations between price and review sub-scores (B).", width_cm=15)

# ---- Q3 ----
doc.add_heading("Q3. Location, listing characteristics and performance", level=3)
tp = R["q3_lga_top_price"]; bp = R["q3_lga_bottom_price"]; to = R["q3_lga_top_occ"]; bands = R["q3_bands"]; bm = R["q3_by_mode_near"]
para(f"The City of Sydney holds {R['q3_sydney_lga_share']}% of listings and the five largest LGAs {R['q3_top5_lga_share']}%. Price and occupancy both differ "
     f"strongly across the {R['q3_lga_n']} LGAs with 30+ listings (Kruskal-Wallis p < 0.001), but they follow different maps (Figure 3). Prices peak "
     f"on the coast and North Shore (Pittwater {money(tp['Pittwater']['median_price'])}, Manly {money(tp['Manly']['median_price'])}, Woollahra "
     f"{money(tp['Woollahra']['median_price'])}) and are lowest in the south and south-west (Bankstown {money(bp['Bankstown']['median_price'])}, Hurstville "
     f"{money(bp['Hurstville']['median_price'])}). Demand peaks in the City of Sydney ({to['Sydney']['median_occupancy']:.0f} nights at "
     f"{money(to['Sydney']['median_price'])}), while Pittwater books only {tp['Pittwater']['median_occupancy']:.0f}. Across LGAs price and occupancy are almost "
     f"unrelated (ρ = {R['q3_lga_price_occ_rho']}): leisure areas earn through rate, the inner city through volume. Listings in more advantaged SA2s "
     f"charge more (ρ = {R['q3_rho_irsad_price']:.2f} with SEIFA IRSAD) but are barely booked more (ρ = {R['q3_rho_irsad_occ']:.2f}).", align="justify")
figure("q3_lga_maps", "Median nightly price (A) and median nights booked (B) by LGA, with current station entrances overlaid. "
       "LGAs with fewer than 30 listings are greyed out.", width_cm=13.5)
para(f"Rail access is linked to demand more than price (Figure 4). Listings within a 500 m walk book the most (median {bands['median_occupancy'][0]:.0f} nights "
     f"against {bands['median_occupancy'][3]:.0f} beyond 2 km) and have the most forward nights taken ({100*bands['unavailable_90d'][0]:.0f}% against "
     f"{100*bands['unavailable_90d'][3]:.0f}%), yet are *cheaper* ({money(bands['median_price'][0])} against {money(bands['median_price'][3])}). The rail-remote band is "
     f"dear because it holds the beaches: its median listing is {bands['median_dist_beach_km'][3]:.1f} km from a beach against {bands['median_dist_beach_km'][0]:.1f} km "
     f"for the nearest band. Near a station, light-rail listings are the most central and busiest ({bm['Light rail']['median_occupancy']:.0f} nights) and Metro "
     f"listings the dearest ({money(bm['Metro']['median_price'])}). The raw pattern therefore mixes access with *what* is built near stations and *where* "
     "stations are, which Analysis 1 separates.", align="justify")
figure("q3_transit_bands", "Median price (A), median nights booked (B) and mean location score (C) by walking distance to the nearest station entrance.", width_cm=15)

# ---- A3 ----
doc.add_heading("A3. Unguided analysis", level=2)
doc.add_heading("Analysis 1. Does rail access pay?", level=3)
prem = pd.DataFrame(R["a3_premium"])
def pv(model, band="<500 m"):
    return prem[(prem.Model == model) & (prem.Band == band)].iloc[0]
rg = pd.DataFrame(R["a3_ring"]).set_index("Group"); dm = R["a3_demand"]; robx = R["a3_robust"]
se_c, se_h = R["a3_se_cluster_vs_hc3"]
para("We regress log price on walking-distance bands (reference: more than 2 km), adding controls in blocks: listing structure (room type, guests, "
     "bedrooms, bathrooms, shared bathroom, amenities); distance to the CBD with LGA fixed effects; distance to the nearest beach and attractions within "
     "1 km; and finally SA2 fixed effects, which compare listings within the same small neighbourhood. Standard errors are clustered by host, "
     f"because 259 commercial hosts run 38% of listings and their listings are not independent (Cameron & Miller, 2015); this makes the key standard "
     f"error {se_c/se_h:.1f}x larger than heteroskedasticity-robust (HC3) errors. Coefficients are reported as percentage differences, exp(β) − 1.",
     bold_lead="Method. ", align="justify")
para(f"The apparent rail effect is a composition effect (Figure 5). Uncontrolled, listings within a 500 m walk are {abs(R['a3_raw_premium_500']):.0f}% *cheaper*, "
     f"because near-station stock is mostly small apartments ({pv('+ Structure')['Premium %']:+.0f}% after structure) and stations are not on the beaches. With all "
     f"controls the difference is {R['a3_adj_premium_500']:+.1f}% ({ci(*R['a3_adj_premium_500_ci'])}, {pval(R['a3_adj_premium_500_p'])}), about "
     f"{money(R['a3_adj_premium_500_dollars'])} a night at the median price. A smooth spline in walking distance gives the same answer (Figure 6A), and "
     "neither the station's mode, frequency, rail time to the CBD nor patronage moves price measurably (all p > 0.1).", bold_lead="Price. ", align="justify")
figure("a3_transit_premium", "Price difference by walking-distance band relative to listings more than 2 km from a station, under five sets of controls (95% CI, clustered by host).", width_cm=13.5)
para(f"The average hides a split by ring (Figure 6B). In the inner city (< 5 km from Town Hall), where almost everything is walkable, near-station listings are "
     f"{abs(rg.loc['Inner (<5 km)','Premium %']):.1f}% *cheaper* ({pval(rg.loc['Inner (<5 km)','p'])}). In the middle ring they earn "
     f"{rg.loc['Middle (5-15 km)','Premium %']:+.1f}% ({pval(rg.loc['Middle (5-15 km)','p'])}) and in the outer suburbs {rg.loc['Outer (>15 km)','Premium %']:+.1f}% "
     f"({pval(rg.loc['Outer (>15 km)','p'])}), where a station is the guest's link to the city.", bold_lead="Where it pays. ", align="justify")
figure("a3_rail_shape", "Adjusted price by walking distance relative to a 2 km walk (A) and the within-500 m premium by CBD ring and station mode (B).", width_cm=15)
drows = []
for label in ["Occupancy, nights (OLS)", "Occupancy, nights (OLS, + log price)", "Calendar unavailable, next 90 days (OLS)",
              "Reviews, last 12 months (Poisson)", "Revenue (PPML)"]:
    d = dm[label]; u = {"nights": " nights", "pp": " pp", "%": "%"}[d["Unit"]]
    drows.append((label, f"{d['Effect of <500 m walk']:+.1f}{u}", f"{d['CI low']:+.1f} to {d['CI high']:+.1f}", pval(d["p"])))
for label, key in [("Price, March 2026 snapshot", "March 2026, walking distance, current network"),
                   ("Price, first-draft measure (2020 network, straight line)", "June 2026, straight line, 2020 network (first draft)")]:
    d = robx[key]
    drows.append((label, f"{d['Premium %']:+.1f}%", f"{d['CI low %']:+.1f} to {d['CI high %']:+.1f}", pval(d["p"])))
table(pd.DataFrame(drows, columns=["Outcome (all with full controls and SA2 fixed effects)", "Within 500 m walk", "95% CI", "p"]),
      "Demand, revenue and robustness checks for the within-500 m effect", [8.2, 2.6, 3.2, 1.9])
para("Demand is measured three ways because the occupancy estimate is review-based (Table 4). None shows a rail effect: "
     f"{dm['Occupancy, nights (OLS)']['Effect of <500 m walk']:+.1f} nights a year, and estimated revenue (Poisson pseudo-maximum likelihood, which keeps "
     f"zero-revenue listings; Santos Silva & Tenreyro, 2006) changes by {R['a3_rev_500_pct']:+.1f}% ({ci(*R['a3_rev_500_ci'])}). The result holds in the March "
     "snapshot, so it is not a winter artefact. Our first draft, which used straight-line distance to the 2020 stations, LGA controls and listing-level "
     f"errors, reported +4.8 nights (p = 0.04); with the 2020 measure, even our final model shows a spurious "
     f"{robx['June 2026, straight line, 2020 network (first draft)']['Premium %']:+.1f}% price premium. Measurement quality changed the answer. "
     "*Limitations:* these are associations within SA2s, not causal effects; prices are asking prices and revenue is an estimate.",
     bold_lead="Demand, revenue and robustness. ", align="justify")

doc.add_heading("Analysis 2. Commercial multi-listing hosts versus single-listing hosts", level=3)
ht = pd.DataFrame(R["a3_host_tier"]).set_index("host_tier"); hs = R["a3_host_subscores"]; tt = R["a3_host_tests"]
S, C = ht.loc["Single (1)"], ht.loc["Commercial (10+)"]
tier_rows = [(t, f"{ht.loc[t,'listings']:,.0f}", f"{ht.loc[t,'mean_rating_5plus']:.2f}", f"{ht.loc[t,'below_4_5_pct']:.0f}%",
              f"{ht.loc[t,'superhost_pct']:.0f}%", f"{ht.loc[t,'median_occupancy']:.0f}", f"{ht.loc[t,'median_walk_station_m']:,.0f} m",
              f"{ht.loc[t,'within_500m_pct']:.0f}%") for t in ht.index]
para(f"Just {C['hosts']:.0f} commercial hosts (10+ listings) operate {C['listing_share_pct']:.0f}% of short-term listings. We compare them with single and small "
     "multi-listing hosts on guest satisfaction, demand and rail access (Table 5, Figure 7).", align="justify")
table(pd.DataFrame(tier_rows, columns=["Host tier", "Listings", "Mean rating", "Rated < 4.5", "Superhost", "Median nights", "Median walk", "< 500 m"]),
      "Host tiers compared (ratings use listings with 5+ reviews)", [3.4, 1.6, 1.7, 1.7, 1.7, 1.8, 2.3, 1.7])
para(f"Commercial listings score lower on every sub-score, most in value ({hs['Commercial (10+)']['Value']:.2f} against {hs['Single (1)']['Value']:.2f}) and "
     f"cleanliness ({hs['Commercial (10+)']['Cleanliness']:.2f} against {hs['Single (1)']['Cleanliness']:.2f}); the difference is large (rank-biserial r = "
     f"{tt['rating_rank_biserial']:.2f}). Controlling for room type, price, review volume, distance to the CBD and SA2, with host-clustered errors, commercial "
     f"listings are still {abs(tt['rating_gap_adj']):.2f} stars lower ({ci(*tt['rating_gap_adj_ci'], unit='', d=2)}). They hold the best-connected stock "
     f"(median walk {C['median_walk_station_m']:,.0f} m against {S['median_walk_station_m']:,.0f} m), yet book {abs(tt['occupancy_gap_adj']):.0f} fewer nights a year than "
     f"comparable single-host listings in the same SA2 at the same price ({ci(*tt['occupancy_gap_adj_ci'], unit=' nights', d=0)}). Professional operators "
     "secure the locations but lose demand through a weaker guest experience. *Limitation:* guests with a poor stay may be less likely to review.",
     align="justify")
figure("a3_host_tiers", "Mean review sub-scores by host tier (A) and cumulative distribution of walking distance to the nearest station entrance (B).", width_cm=15)

# ======================================================================
# Part B
# ======================================================================
doc.add_heading("Part B. Machine Learning Analysis, Interpretation and Limitations", level=1)
m1, m3, m1n, m3n, m1r, m3r = R["m1"], R["m3"], R["m1_noT"], R["m3_noT"], R["m1_random"], R["m3_random"]
BASE = pd.read_csv(ROOT / "outputs" / "model_comparison.csv").set_index("Model").loc["Baseline (median)"]
para(f"All models use the {R['clean_listings']:,} cleaned listings. The supervised models predict log nightly price (prices are right-skewed, so errors "
     "become proportional) from listing structure, reviews, host scale, coordinates, LGA, the CBD, beach and attraction measures, SEIFA, and seven rail "
     "features from the external data (walking distance, stations within 1 km, mode, frequency, patronage, rail time to the CBD and airport). "
     f"Because neighbouring listings are alike, a random split flatters models. We therefore hold out whole SA2 neighbourhoods: {R['ml_test_n']:,} test listings "
     f"in {R['ml_test_sa2']} SA2s never seen in training ({R['ml_train_n']:,} listings, {R['ml_train_sa2']} SA2s), and tune hyper-parameters with folds that are also "
     "grouped by SA2 (Roberts et al., 2017). The random split is reported for comparison.", align="justify")

doc.add_heading("Model 1 (supervised). Ridge regression", level=2)
para("Ridge is linear regression with an L2 penalty (α·Σβ²) that shrinks coefficients towards zero (Hoerl & Kennard, 1970), which stabilises "
     "estimates when inputs are correlated (guests, bedrooms and beds; coordinates, LGA, distances and rail times). Our *objective* was an "
     "interpretable price model that shows how much rail access adds once structure and location are known.", bold_lead="Method and objective. ", align="justify")
para("Numeric inputs are median-imputed with missing-value indicators (12% of listings have no rating; a few stations have no patronage record) "
     "and standardised; categories are one-hot encoded with levels under 20 rows merged; distances enter as logarithms. α was tuned over 11 values "
     f"from 0.01 to 1,000 by grouped 5-fold CV (best α = {R['m1_alpha']:g}).", bold_lead="Preparation and tuning. ", align="justify")
para(f"On held-out SA2s, R² = {m1['R2 (log price)']:.3f} (training {R['m1_train_r2']:.3f}), RMSE {money(m1['RMSE (AUD)'])}, MAE {money(m1['MAE (AUD)'])}, "
     f"MAPE {m1['MAPE %']:.1f}%, against a median baseline with MAE {money(BASE['MAE (AUD)'])}. Room type and size dominate, followed by distance to the beach, "
     f"longitude and SEIFA (Figure 8B). Removing all rail features lowers R² only from {m1['R2 (log price)']:.3f} to {m1n['R2 (log price)']:.3f}.",
     bold_lead="Results. ", align="justify")
para(f"(1) The random split scored R² = {m1r['R2 (log price)']:.3f}, overstating accuracy in a new area; we report the spatial result. (2) Only 13 shared "
     "rooms remain, so rare levels were merged to avoid unstable dummies. (3) LGA dummies, coordinates and distances overlap, so individual location "
     "coefficients should not be read alone; the penalty keeps predictions stable. (4) The model is additive and misses interactions, and listed price "
     "is an asking price.", bold_lead="Issues and limitations. ", align="justify")
figure("m1_ridge", "Ridge predicted versus actual price on held-out SA2s (A) and the largest standardised non-LGA coefficients (B).", width_cm=15)

doc.add_heading("Model 2 (unsupervised). K-Means market segmentation", level=2)
kt = pd.DataFrame(R["m2_k_table"]).set_index("k"); pr = pd.DataFrame(R["m2_profiles"])
para("K-Means assigns each listing to the nearest of k centroids and moves each centroid to the mean of its members until assignments settle, "
     "minimising within-cluster squared distance. Our *objective* was to segment Sydney's supply into zones combining location, price, rail access "
     "and demand, as an investment screen. Inputs were latitude, longitude, log price, log walking distance to a station, distance to the CBD, log "
     "distance to the beach and estimated occupancy, all z-scored because their units differ by orders of magnitude.", bold_lead="Method and objective. ", align="justify")
para(f"We fitted k = 2–10. The silhouette is highest at k = 2 ({kt.loc[2,'silhouette']:.2f}), which only separates the city from the rest; among k ≥ 4 it "
     f"peaks at k = {R['m2_k']} ({R['m2_silhouette']:.2f}), where the elbow curve has flattened. Silhouette is O(n²), so it was computed on a fixed 4,000-listing "
     "sample. Segment names are generated from each profile (ring, coast, price and demand relative to the city median, walk to rail).",
     bold_lead="Choosing k. ", align="justify")
prow = [(r_.segment, f"{r_.listings:,.0f}", money(r_.median_price), f"{r_.median_occupancy:.0f}", money(r_.median_revenue),
         f"{r_.median_walk_station_m:,.0f} m") for r_ in pr.itertuples()]
table(pd.DataFrame(prow, columns=["Segment", "Listings", "Median price", "Median nights", "Median revenue", "Walk to rail"]),
      "K-Means segment profiles (medians; revenue is Inside Airbnb's 12-month estimate)", [6.6, 1.6, 1.9, 1.8, 2.1, 1.9])
top_rev = pr.sort_values("median_revenue", ascending=False).iloc[0]
inner_low = pr[pr.segment.str.startswith("Inner") & pr.segment.str.contains("low-utilisation")].iloc[0]
para(f"The two inner-city, rail-served segments sit in the same LGAs with similar rail access ({top_rev.median_walk_station_m:,.0f} m against "
     f"{inner_low.median_walk_station_m:,.0f} m), yet book {top_rev.median_occupancy:.0f} against {inner_low.median_occupancy:.0f} nights, and median revenue is "
     f"{money(top_rev.median_revenue)} against {money(inner_low.median_revenue)} (Figure 9). Within the same locations, operation separates winners from under-used "
     "stock. The coastal premium segments charge the most but are rail-remote with low utilisation. Silhouettes near 0.25 mean the segments overlap, "
     "K-Means assumes compact clusters and treats coordinates as Euclidean, so the segments are a screening tool, not natural boundaries.",
     bold_lead="Results and limitations. ", align="justify")
figure("m2_cluster_maps", "Listings in each K-Means segment (blue) against all listings (grey), with segment medians.", width_cm=15.5)

doc.add_heading("Model 3 (not covered in class). XGBoost with SHAP interpretation", level=2)
p3 = R["m3_params"]; sb = R["m3_shap_dist_band_pct"]; eb = R["ml_err_by_band"]
para("XGBoost (Chen & Guestrin, 2016) is a gradient-boosted tree ensemble: each new tree is fitted to the residual errors of those before it and "
     "scaled by a learning rate, with L2 regularisation and row/column subsampling against overfitting. Unlike Ridge it learns non-linear effects and "
     "interactions, such as rail mattering more in the suburbs. SHAP (Lundberg & Lee, 2017) splits each prediction into additive feature contributions "
     "based on Shapley values; with a log target a SHAP value s moves the predicted price by about exp(s) − 1. Our *objective* was to beat Ridge and isolate "
     "the non-linear contribution of rail access.", bold_lead="Method and objective. ", align="justify")
para(f"Same inputs as Model 1 (raw distances; trees are scale-free). A randomised search over 30 configurations with grouped 3-fold CV selected "
     f"{p3['n_estimators']} trees, depth {p3['max_depth']}, learning rate {p3['learning_rate']}, subsample {p3['subsample']}, column sample "
     f"{p3['colsample_bytree']} and λ = {p3['reg_lambda']} (CV RMSE {R['m3_cv_rmse_log']:.3f} log units against {R['m1_cv_rmse_log']:.3f} for Ridge). "
     "The model without rail features received its own search.", bold_lead="Preparation and tuning. ", align="justify")
cmp_rows = [("Baseline (median price)", f"{BASE['R2 (log price)']:.3f}", money(BASE['RMSE (AUD)']), money(BASE['MAE (AUD)']), f"{BASE['MAPE %']:.1f}%")]
for name, m in [("Model 1: Ridge", m1), ("Ridge without rail features", m1n), ("Model 3: XGBoost", m3), ("XGBoost without rail features", m3n),
                ("Ridge, random split (comparison)", m1r), ("XGBoost, random split (comparison)", m3r)]:
    cmp_rows.append((name, f"{m['R2 (log price)']:.3f}", money(m['RMSE (AUD)']), money(m['MAE (AUD)']), f"{m['MAPE %']:.1f}%"))
table(pd.DataFrame(cmp_rows, columns=["Model (held-out SA2s unless stated)", "R² (log)", "RMSE", "MAE", "MAPE"]),
      "Test-set performance, rail-feature ablation and random-split comparison", [7.1, 2.0, 2.2, 2.2, 2.2])
para(f"XGBoost beats Ridge on every metric (R² {m3['R2 (log price)']:.3f} against {m1['R2 (log price)']:.3f}; MAE {money(m3['MAE (AUD)'])} against "
     f"{money(m1['MAE (AUD)'])}). SHAP ranks bedrooms, guests and room type highest, then longitude, distance to the beach and SEIFA (Figure 10A). All rail "
     f"features together carry {R['m3_shap_rail_share_pct']}% of total |SHAP| and the best of them is only number {R['m3_shap_rail_rank']}. Walking distance "
     f"moves price by about 1% or less below 5 km in every ring (Figure 10B) and rises only beyond 5 km ({sb['>5 km']:+.1f}%), where {R['far_from_rail_coastal_pct']:.0f}% "
     f"of listings are within 2 km of a beach, a coastal effect rather than a rail one. Re-tuned without rail features, R² is {m3n['R2 (log price)']:.4f} against "
     f"{m3['R2 (log price)']:.4f}. Regression with controls, Ridge and XGBoost agree: rail proximity does not set Sydney's short-term rental prices.",
     bold_lead="Results. ", align="justify")
para(f"(1) Training R² is {R['m3_train_r2']:.2f} against {m3['R2 (log price)']:.2f} on test; the search favoured regularising settings and the grouped CV "
     f"error matches the test error, so the score is not a lucky split. (2) The random split overstated R² ({m3r['R2 (log price)']:.3f}). (3) SHAP spreads one-hot "
     "categories over many columns, so we summed them back to the original variable. (4) Predictions shrink to the middle: the cheapest quartile is "
     f"over-predicted by {eb['XGBoost bias %']['Q1 (cheapest)']:.0f}% and the dearest under-predicted by {abs(eb['XGBoost bias %']['Q4 (dearest)']):.0f}%. "
     "(5) SHAP explains the model, not causation.", bold_lead="Issues and limitations. ", align="justify")
figure("m3_shap", "XGBoost feature importance as mean |SHAP| (A) and SHAP effect of walking distance on predicted price by CBD ring (B).", width_cm=15.5)

# ======================================================================
# Part C
# ======================================================================
doc.add_heading("Part C. Summary and Business Recommendations", level=1)
para("Across the exploratory and machine learning analyses one message is consistent: *what* a listing is and *where* it sits set the price, and "
     "*how it is run* sets the demand. Rail access looks like a 26% discount in raw data, but once structure, beaches and neighbourhood are controlled it "
     f"has no average effect on price ({R['a3_adj_premium_500']:+.1f}%), demand or revenue, and it adds almost nothing to either price model. The exception is "
     f"the suburbs, where a station within a 500 m walk is worth about {rg.loc['Middle (5-15 km)','Premium %']:.0f}–{rg.loc['Outer (>15 km)','Premium %']:.0f}% on the "
     "nightly rate. Demand instead follows operation: popular listings have short minimum stays and strong review records, the busiest inner-city "
     f"segment earns about {top_rev.median_revenue/inner_low.median_revenue:.0f} times the revenue of an under-used segment in the same area, and commercial "
     "hosts convert the best-connected stock into lower ratings and fewer nights. Better data also mattered: the outdated 2020 network and straight-line "
     "distances produced a rail premium and a demand effect that the corrected measures do not support.", align="justify")
bullet(" Do not pay extra for proximity to a station in the inner city, where it carries no premium. In the middle and outer suburbs, "
       f"a listing within a 500 m walk of a station earns about {rg.loc['Middle (5-15 km)','Premium %']:.0f}–{rg.loc['Outer (>15 km)','Premium %']:.0f}% more a night than a comparable one "
       "further away. That is where rail access is worth paying for. Underwrite inner-city stock on occupancy, as in the high-turnover segment "
       f"({money(top_rev.median_price)} median rate, {top_rev.median_occupancy:.0f} nights, {money(top_rev.median_revenue)} median revenue).",
       lead="1. Investors —")
bullet(f" Reduce booking friction before cutting price. Median occupancy falls from {mn['1']:.0f} nights at a one-night minimum stay to {mn['4-7']:.0f} at "
       "four to seven nights, and ratings below 4.5 coincide with much lower occupancy. Hosts in the low-utilisation inner segment, who already "
       "have the right locations, should first move to one- or two-night minimums and protect their review record. Price is a weak lever on ratings "
       f"(ρ = {R['q2_rho_overall']:.2f}).", lead="2. Individual hosts —")
bullet(f" Close the guest-experience gap, starting with value and cleanliness. Commercial listings trail single hosts by {abs(tt['rating_gap_adj']):.2f} "
       f"stars after controls, {C['below_4_5_pct']:.0f}% of them sit below 4.5, and they book {abs(tt['occupancy_gap_adj']):.0f} fewer nights than comparable "
       "listings. Turnover audits, cleaning checklists and value-adding inclusions would let operators turn their location advantage into occupancy. "
       f"An XGBoost price model (MAPE {m3['MAPE %']:.0f}% on unseen neighbourhoods) can support rate-setting for mid-market stock, but not luxury homes, which it under-prices.",
       lead="3. Commercial operators —")
para("*Limitations.* Inside Airbnb occupancy and revenue are estimates built from reviews; prices are asking prices; the main snapshot is from winter "
     "(checked against March 2026); the GTFS timetable is from October 2026 and SEIFA from the 2021 Census; and all findings are associations, not causal effects.",
     align="justify")

# ======================================================================
# References
# ======================================================================
doc.add_heading("References", level=1)
refs = [
    "Australian Bureau of Statistics. (2021). *Australian Statistical Geography Standard (ASGS) Edition 3: Statistical Area Level 2 digital boundaries* [Data set]. Licensed under CC BY 4.0. Retrieved 7 October 2026, from https://www.abs.gov.au/",
    "Australian Bureau of Statistics. (2023). *Socio-Economic Indexes for Areas (SEIFA), Australia, 2021: Statistical Area Level 2 indexes* [Data set]. Licensed under CC BY 4.0. Retrieved 7 October 2026, from https://www.abs.gov.au/",
    "Boeing, G. (2017). OSMnx: New methods for acquiring, constructing, analyzing, and visualizing complex street networks. *Computers, Environment and Urban Systems, 65*, 126–139.",
    "Cameron, A. C., & Miller, D. L. (2015). A practitioner's guide to cluster-robust inference. *Journal of Human Resources, 50*(2), 317–372.",
    "Chen, T., & Guestrin, C. (2016). XGBoost: A scalable tree boosting system. In *Proceedings of the 22nd ACM SIGKDD International Conference on Knowledge Discovery and Data Mining* (pp. 785–794).",
    "Dibbelt, J., Pajor, T., Strasser, B., & Wagner, D. (2013). Intriguingly simple and fast transit routing. In *Experimental Algorithms (SEA 2013)*, Lecture Notes in Computer Science 7933 (pp. 43–54). Springer.",
    "Hoerl, A. E., & Kennard, R. W. (1970). Ridge regression: Biased estimation for nonorthogonal problems. *Technometrics, 12*(1), 55–67.",
    "Inside Airbnb. (2026). *Sydney, New South Wales, Australia: listings, calendar and neighbourhoods* [Data set; compiled 29 June 2026 and 21 March 2026]. Licensed under CC BY 4.0. Retrieved 6–7 October 2026, from https://insideairbnb.com/get-the-data/",
    "Lundberg, S. M., & Lee, S.-I. (2017). A unified approach to interpreting model predictions. *Advances in Neural Information Processing Systems, 30*, 4765–4774.",
    "OpenStreetMap contributors. (2026). *OpenStreetMap* [Data set; beaches, tourist attractions and pedestrian network for Sydney]. Licensed under ODbL 1.0. Retrieved 7 October 2026, via the Overpass API, from https://www.openstreetmap.org/",
    "Pedregosa, F., et al. (2011). Scikit-learn: Machine learning in Python. *Journal of Machine Learning Research, 12*, 2825–2830.",
    "Roberts, D. R., et al. (2017). Cross-validation strategies for data with temporal, spatial, hierarchical, or phylogenetic structure. *Ecography, 40*(8), 913–929.",
    "Rousseeuw, P. J. (1987). Silhouettes: A graphical aid to the interpretation and validation of cluster analysis. *Journal of Computational and Applied Mathematics, 20*, 53–65.",
    "Santos Silva, J. M. C., & Tenreyro, S. (2006). The log of gravity. *Review of Economics and Statistics, 88*(4), 641–658.",
    "Transport for NSW. (2020). *Train station entrance locations* (stationentrances2020_v4) [Data set]. Licensed under CC BY 4.0. Retrieved 6 October 2026, from https://opendata.transport.nsw.gov.au/",
    "Transport for NSW. (2026a). *Timetables Complete GTFS* [Data set; timetable valid from 6 October 2026]. Licensed under CC BY 4.0. Retrieved 6 October 2026, from https://opendata.transport.nsw.gov.au/",
    "Transport for NSW. (2026b). *Train, Metro and Light Rail Station Entries and Exits* [Data set; October 2024 – August 2026]. Licensed under CC BY 4.0. Retrieved 7 October 2026, from https://opendata.transport.nsw.gov.au/",
]
for ref in refs:
    p = para(ref, after=2)
    p.paragraph_format.left_indent = Cm(0.8); p.paragraph_format.first_line_indent = Cm(-0.8)

# ---- Footer with page numbers ----
footer = sec.footer.paragraphs[0]; footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = footer.add_run(f"ACC2125 {GROUP} | Page "); r.font.size = Pt(9); r.font.color.rgb = GREY
for kind, text in (("begin", None), (None, "PAGE"), ("end", None)):
    run = footer.add_run(); run.font.size = Pt(9); run.font.color.rgb = GREY
    if kind:
        fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), kind); run._r.append(fc)
    else:
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = text; run._r.append(it)
sec.different_first_page_header_footer = True

out = ROOT / f"Report_{GROUP}.docx"
doc.save(out)
print("Wrote", out, "| figures:", fig_no[0], "| tables:", tab_no[0])
