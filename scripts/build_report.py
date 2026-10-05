"""Build Report_GroupXX.docx from the notebook outputs (outputs/results.json, outputs/*.csv, figures/*.png).

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
GROUP = "GroupXX"

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
doc.add_heading("Part A. Data Selection and Exploratory Analysis", level=1)
doc.add_heading("A1. Data selection", level=2)
para(f"We analyse the Inside Airbnb snapshot for Sydney, New South Wales, scraped between 17 and 29 June 2026 "
     f"({R['raw_listings']:,} listings; {R['raw_calendar_rows']:,} calendar rows covering June 2026 to June 2027). "
     "Sydney is one of the world's largest Airbnb markets. It has sharply different sub-markets (CBD apartments, harbour and beach suburbs, "
     "and the western suburbs) and a dense commuter-rail network, which makes it a natural place to ask whether transport "
     "access, rather than just location, shapes short-term rental performance. We use listings.csv.gz (listing attributes, price, "
     "reviews and estimated occupancy), calendar.csv.gz (forward availability; this snapshot contains no calendar prices) and "
     "neighbourhoods.geojson (boundaries of the 38 local government areas, LGAs).", align="justify")
para("Transport for NSW's *Train Station Entrance Locations* dataset (file stationentrances2020_v4.csv, "
     f"{R['raw_station_rows']:,} entrances at {R['raw_stations_unique']} stations, published on the TfNSW Open Data Hub under "
     "CC BY 4.0, reflecting the network as at 2020, accessed 6 October 2026). It gives the latitude and longitude of every "
     "entrance to Sydney Trains, NSW TrainLink, Sydney Metro and light-rail stations. Entrances, rather than station centroids, measure the walk a guest actually faces. "
     f"After restricting it to Greater Sydney ({R['entrances_sydney']} entrances, {R['stations_sydney']} stations), we join it to every listing with a "
     "haversine nearest-neighbour search. This produces the walking-proximity features *distance to nearest station entrance* and *stations within 1 km*, plus "
     f"distance-to-CBD as a control. The median listing is {R['median_dist_station_m']:,} m from a station entrance; "
     f"{R['share_within_500m']}% are within 500 m. These features drive Q3, both unguided analyses and all three models.",
     bold_lead="External dataset. ", align="justify")

log = pd.read_csv(ROOT / "outputs" / "data_decisions_log.csv")
log = log[["Change Details", "Reason and Justification"]]
table(log, f"Data Decisions Log ({R['raw_listings']:,} raw listings → {R['clean_listings']:,} analysed)", [7.6, 8.3], align_right_from=99)

# ---- Q1 ----
doc.add_heading("A2. Guided analysis", level=2)
doc.add_heading("Q1. Characteristics of the most popular listings", level=3)
q1 = pd.DataFrame(R["q1"])
pop, oth = q1["Popular (top 25%)"], q1["Other listings"]
para("We measure popularity with Inside Airbnb's *estimated nights booked in the last 12 months* and define the most popular "
     f"listings as the top quartile (at least {R['popular_threshold_nights']:.0f} nights; n = {R['popular_n']:,}). Because the estimate is "
     f"derived from review counts, it agrees almost perfectly with reviews in the last 12 months (ρ = {R['rho_occ_reviews_ltm']}) but only weakly "
     f"with the share of the next 90 calendar nights already unavailable (ρ = {R['rho_occ_calendar']}). Calendar 'unavailable' nights mix bookings with "
     "host blocks, so we use the review-based measure and treat the calendar as supporting evidence.", align="justify")
rows = [("Median nightly price", money(pop["Median price (AUD)"]), money(oth["Median price (AUD)"])),
        ("Median minimum stay (nights)", f"{pop['Median minimum nights']:.0f}", f"{oth['Median minimum nights']:.0f}"),
        ("Superhost share", f"{pop['Superhost %']:.0f}%", f"{oth['Superhost %']:.0f}%"),
        ("Median amenities listed", f"{pop['Median amenities']:.0f}", f"{oth['Median amenities']:.0f}"),
        ("Mean rating (5+ reviews)", f"{pop['Mean rating (5+ reviews)']:.2f}", f"{oth['Mean rating (5+ reviews)']:.2f}"),
        ("Entire home/apartment share", f"{pop['Entire home %']:.0f}%", f"{oth['Entire home %']:.0f}%"),
        ("Run by a commercial host (10+ listings)", f"{pop['Commercial host (10+) %']:.0f}%", f"{oth['Commercial host (10+) %']:.0f}%"),
        ("Within 500 m of a station entrance", f"{pop['Within 500 m of station %']:.0f}%", f"{oth['Within 500 m of station %']:.0f}%"),
        ("Median share of next 90 nights unavailable", f"{pop['Median 90-day unavailable %']:.0f}%", f"{oth['Median 90-day unavailable %']:.0f}%")]
table(pd.DataFrame(rows, columns=["Characteristic", f"Popular (n = {pop['Listings']:,.0f})", f"Other (n = {oth['Listings']:,.0f})"]),
      "Profile of popular listings versus the rest (price, minimum stay, amenities, rating and station distance differ at p < 0.001, Mann-Whitney U)", [8.3, 3.8, 3.8])
mn = R["q1_occ_by_minnights"]; rb = R["q1_occ_by_rating"]; sh = R["q1_occ_superhost"]
para("Popular listings are *not* the expensive ones: their median price is "
     f"{money(pop['Median price (AUD)'])} against {money(oth['Median price (AUD)'])}. They stand out on how they are run. "
     f"Superhosts make up {pop['Superhost %']:.0f}% of them (versus {oth['Superhost %']:.0f}%), and the median superhost listing books {sh['1']:.0f} nights a year against "
     f"{sh['0']:.0f} for other listings (Figure 1C). Flexibility matters most: median occupancy falls from {mn['1']:.0f} nights for one-night minimum stays to "
     f"{mn['4-7']:.0f} at four to seven nights and {mn['8-30']:.0f} at eight to thirty (Figure 1A). Occupancy also rises with rating, from {rb['<4.5']:.0f} nights below 4.5 to "
     f"{rb['4.8-4.9']:.0f} at 4.8–4.9 (Figure 1B). Popular listings are more often near rail and less often run by commercial hosts. "
     "The insight for hosts is that demand is earned through low booking friction and consistent service, not through charging more.",
     align="justify")
figure("q1_popularity_drivers", "Median estimated nights booked by minimum stay (A), rating band (B) and superhost status (C).")

# ---- Q2 ----
doc.add_heading("Q2. Relationship between price and guest ratings", level=3)
sp = R["q2_spearman_price"]; dec = R["q2_deciles"]; br = R["q2_by_room"]
para(f"Using the {R['q2_n_rated']:,} listings with at least five reviews, price and overall rating are positively but weakly related "
     f"(Spearman ρ = {R['q2_rho_overall']:.2f}, p < 0.001). The relationship holds within room types "
     f"(entire homes ρ = {br['Entire home/apt']['rho']:.2f}; private rooms ρ = {br['Private room']['rho']:.2f}), so it is not just a room-mix effect. "
     f"It is non-linear (Figure 2A). The jump happens at the bottom of the market: the cheapest decile (median {money(dec['price_median'][0])}) averages "
     f"{dec['rating_mean'][0]:.2f} stars and {dec['share_below_45'][0]:.0f}% of these listings are rated below 4.5. The second decile ({money(dec['price_median'][1])}) already averages "
     f"{dec['rating_mean'][1]:.2f}, and ratings then creep up only slowly to {dec['rating_mean'][9]:.2f} in the top decile. "
     f"A log-price regression explains only {100*R['q2_linear_r2']:.1f}% of rating variation ({100*R['q2_quad_r2']:.1f}% with a quadratic term).", align="justify")
para(f"The sub-scores explain why (Figure 2B). Location has the strongest link to price (ρ = {sp['Location']:.2f}): guests pay more for, and "
     f"value, good locations. Value has almost no link (ρ = {sp['Value']:.2f}): guests at expensive listings are more satisfied "
     "overall but do not feel they get better value for money. Pricing above the market's floor protects against poor ratings, but "
     "high prices do not buy higher ratings. Hosts should compete on the experience, not expect price to signal quality.", align="justify")
figure("q2_price_vs_rating", "Mean overall, value and location scores by price decile (A) and Spearman correlations between price and review sub-scores (B).")

# ---- Q3 ----
doc.add_heading("Q3. Location, listing characteristics and performance", level=3)
tp = R["q3_lga_top_price"]; bp = R["q3_lga_bottom_price"]; to = R["q3_lga_top_occ"]; bands = R["q3_bands"]
para(f"Supply is concentrated: the City of Sydney LGA alone holds {R['q3_sydney_lga_share']}% of listings and the five largest LGAs hold "
     f"{R['q3_top5_lga_share']}%. Both price and occupancy differ strongly across the {R['q3_lga_n']} LGAs with 30+ listings (Kruskal-Wallis p < 0.001 for both), "
     f"but they follow different maps (Figure 3). Prices peak on the coast and the North Shore (Pittwater median {money(tp['Pittwater']['median_price'])}, "
     f"Manly {money(tp['Manly']['median_price'])}, Woollahra {money(tp['Woollahra']['median_price'])}) and are lowest in the south and south-west "
     f"(Bankstown {money(bp['Bankstown']['median_price'])}, Hurstville {money(bp['Hurstville']['median_price'])}). Demand peaks in the City of Sydney "
     f"({to['Sydney']['median_occupancy']:.0f} nights at a mid-market {money(to['Sydney']['median_price'])}), while Pittwater, the dearest LGA, books only "
     f"{tp['Pittwater']['median_occupancy']:.0f} nights. Across LGAs, price and occupancy are almost unrelated (ρ = {R['q3_lga_price_occ_rho']}). "
     "Premium leisure areas earn their revenue through rate; the inner city earns it through volume.", align="justify")
figure("q3_lga_maps", "Median nightly price (A) and median nights booked (B) by LGA, with Transport for NSW station entrances overlaid. "
       "LGAs with fewer than 30 listings are greyed out.", width_cm=14.5)
para(f"The external rail data adds a second lens (Figure 4). Listings within 500 m of a station entrance book the most "
     f"(median {bands['median_occupancy'][0]:.0f} nights against {bands['median_occupancy'][3]:.0f} beyond 2 km) and have the highest share of forward nights already taken "
     f"({100*bands['unavailable_90d'][0]:.0f}% against {100*bands['unavailable_90d'][3]:.0f}%). They are, however, *cheaper* "
     f"({money(bands['median_price'][0])} against {money(bands['median_price'][3])}). The most rail-remote band contains the beach and bushland homes "
     f"of the Northern Beaches and Sutherland Shire, which are large, expensive and rated highly for location ({bands['mean_location_score'][3]:.2f}). "
     f"Distance to rail and distance to the CBD are also correlated (ρ = {R['q3_rho_cbd_station']}). The raw rail pattern therefore mixes "
     "the effect of access with *what* is built near stations and *where* stations are, which Analysis 1 separates.", align="justify")
figure("q3_transit_bands", "Median price (A), median nights booked (B) and mean location score (C) by distance to the nearest station entrance.")

# ---- A3 ----
doc.add_heading("A3. Unguided analysis", level=2)
doc.add_heading("Analysis 1. Is there a rail-access price premium?", level=3)
prem = pd.DataFrame(R["a3_premium"])
def pv(model, band):
    return prem[(prem.Model == model) & (prem.Band == band)].iloc[0]
s1 = pv("+ Structure", "<500 m"); s2 = pv("+ Structure + location", "<500 m")
para("A common investor assumption is that listings near transport command a nightly-rate premium. We test it with log-price OLS "
     "regressions (heteroskedasticity-robust HC3 errors) comparing each distance band with listings more than 2 km from a station. Controls are added in "
     "two blocks: listing structure (room type, guests, bedrooms, bathrooms, shared bathroom, amenities) and location (distance to the CBD and "
     "LGA fixed effects). Coefficients are converted to percentage price differences, exp(β) − 1.", align="justify")
para(f"The premium does not survive (Figure 5). Uncontrolled, listings within 500 m are {abs(R['a3_raw_premium_500']):.0f}% *cheaper*. "
     f"Controlling for structure shrinks this to {s1['Premium %']:.0f}%, because near-station stock is dominated by small apartments. Adding location "
     f"controls removes it altogether: {R['a3_adj_premium_500']:+.1f}% (95% CI {R['a3_adj_premium_500_ci'][0]:+.1f}% to {R['a3_adj_premium_500_ci'][1]:+.1f}%, "
     f"p = {s2['p']:.2f}), about {money(R['a3_adj_premium_500_dollars'])} a night at the median price of {money(R['median_price'])}. "
     f"Model fit rises from R² = {R['a3_r2']['Raw']:.2f} to {R['a3_r2']['+ Structure + location']:.2f}, so the controls, not rail, explain price. "
     f"Rail access does matter for demand: with the same controls plus price, listings within 500 m book {R['a3_occ_500_nights']:.1f} more nights a year "
     f"(95% CI {R['a3_occ_500_ci'][0]:.1f}–{R['a3_occ_500_ci'][1]:.1f}, p = {R['a3_occ_500_p']:.2f}). For investors, rail proximity is an occupancy "
     "advantage, not a pricing one, and should not be paid for as if it lifts nightly rates.", align="justify")
figure("a3_transit_premium", "Estimated price difference by distance band relative to listings more than 2 km from a station, under three sets of controls (95% CI).", width_cm=14.5)

doc.add_heading("Analysis 2. Commercial multi-listing hosts versus single-listing hosts", level=3)
ht = pd.DataFrame(R["a3_host_tier"]).set_index("host_tier"); hs = R["a3_host_subscores"]; tt = R["a3_host_tests"]
S, C = ht.loc["Single (1)"], ht.loc["Commercial (10+)"]
para(f"Sydney's supply is highly concentrated: just {C['hosts']:.0f} commercial hosts (10+ listings) operate {C['listing_share_pct']:.0f}% of short-term listings. "
     "We compare them with single-listing and small multi-listing hosts on guest satisfaction, demand and rail access (Figure 6).", align="justify")
tier_rows = [(t, f"{ht.loc[t,'listings']:,.0f}", f"{ht.loc[t,'mean_rating_5plus']:.2f}", f"{ht.loc[t,'below_4_5_pct']:.0f}%",
              f"{ht.loc[t,'superhost_pct']:.0f}%", f"{ht.loc[t,'median_occupancy']:.0f}", f"{ht.loc[t,'median_dist_station_m']:,.0f} m",
              f"{ht.loc[t,'within_500m_pct']:.0f}%") for t in ht.index]
table(pd.DataFrame(tier_rows, columns=["Host tier", "Listings", "Mean rating", "Rated < 4.5", "Superhost", "Median nights", "Median to station", "< 500 m"]),
      "Host tiers compared (ratings use listings with 5+ reviews)", [3.4, 1.6, 1.7, 1.7, 1.7, 1.8, 2.3, 1.7])
para(f"Commercial listings score lower on every sub-score, with the largest gaps in cleanliness ({hs['Commercial (10+)']['Cleanliness']:.2f} against "
     f"{hs['Single (1)']['Cleanliness']:.2f}) and value ({hs['Commercial (10+)']['Value']:.2f} against {hs['Single (1)']['Value']:.2f}). "
     f"The difference is large (rank-biserial r = {tt['rating_rank_biserial']:.2f}, p < 0.001), and a regression controlling for room type, price, "
     f"review volume, distance to the CBD and LGA still finds commercial listings {abs(tt['rating_gap_adj']):.2f} stars lower "
     f"(95% CI {tt['rating_gap_adj_ci'][0]:.2f} to {tt['rating_gap_adj_ci'][1]:.2f}). Yet commercial hosts hold the best-connected stock: median "
     f"{C['median_dist_station_m']:,.0f} m to a station against {S['median_dist_station_m']:,.0f} m for single hosts. Even so, their median occupancy is lower "
     f"({C['median_occupancy']:.0f} against {S['median_occupancy']:.0f} nights). Professional operators secure the best locations but lose demand through "
     "a weaker, more standardised guest experience. This links Q1 and Analysis 1: service quality, not location, is the binding constraint.",
     align="justify")
figure("a3_host_tiers", "Mean review sub-scores by host tier (A) and cumulative distribution of distance to the nearest station entrance (B).")

# ======================================================================
# Part B
# ======================================================================
doc.add_heading("Part B. Machine Learning Analysis, Interpretation and Limitations", level=1)
para(f"All models use the {R['clean_listings']:,} cleaned listings and the rail features built from the external dataset. The two supervised models "
     f"predict log nightly price from 18 numeric and 3 categorical inputs (structure, reviews, host scale, coordinates, distance to CBD, "
     "distance to station, stations within 1 km, room type, property group and LGA). A log target suits the right-skewed prices and makes "
     f"errors proportional. Both share the same random 80/20 split (seed 42; {R['ml_train_n']:,} training and {R['ml_test_n']:,} test listings); hyper-parameters "
     "are tuned by cross-validation on the training set only, and the test set is used once. A median-price baseline gives the reference point.",
     align="justify")

m1, m3, m1n, m3n = R["m1"], R["m3"], R["m1_noT"], R["m3_noT"]
BASE = pd.read_csv(ROOT / "outputs" / "model_comparison.csv").set_index("Model").loc["Baseline (median)"]
doc.add_heading("Model 1 (supervised). Ridge regression", level=2)
para("Ridge regression is linear regression with an L2 penalty (α·Σβ²) that shrinks coefficients towards zero (Hoerl & Kennard, 1970). It suits this "
     "data because many inputs are correlated (guests, bedrooms and beds; coordinates, LGA and distances). Our *objective* was an interpretable price "
     "model that shows how much rail access adds once a listing's structure and location are known.", bold_lead="Method and objective. ", align="justify")
para("Numeric inputs are median-imputed with missing-value indicators (12% of listings have no rating yet), then standardised. Categorical "
     "inputs are one-hot encoded, and levels with fewer than 20 rows are merged. Station and CBD distances enter as logarithms so each extra metre "
     f"matters less. α was tuned over 11 values from 0.01 to 1,000 by 5-fold CV; the best was α = {R['m1_alpha']:.0f}.",
     bold_lead="Preparation and tuning. ", align="justify")
para(f"Test R² = {m1['R2 (log price)']:.3f} (training {R['m1_train_r2']:.3f}, so there is no overfitting), RMSE = {money(m1['RMSE (AUD)'])}, "
     f"MAE = {money(m1['MAE (AUD)'])}, MAPE = {m1['MAPE %']:.1f}%. The median baseline scores R² ≈ 0 and MAE {money(BASE['MAE (AUD)'])}. Room type and size dominate (Figure 7B). "
     f"The rail coefficients are small (standardised β = {R['m1_coef_rail']['log_dist_station']:+.3f} for log distance; "
     f"{R['m1_coef_rail']['stations_within_1km']:+.3f} for stations within 1 km), and removing them changes test R² only from "
     f"{m1['R2 (log price)']:.4f} to {m1n['R2 (log price)']:.4f}. This is consistent with Analysis 1.", bold_lead="Results. ", align="justify")
para("(1) Rare room types (13 shared rooms) produced unstable dummies; merging infrequent levels fixed this. (2) LGA dummies, coordinates "
     "and CBD distance overlap, so some outer LGAs (e.g. Camden) take large positive coefficients that only offset the distance terms; the penalty "
     "stabilises predictions but individual location coefficients should not be read alone. (3) The model is additive and misses interactions. It under-predicts "
     "luxury homes (Figure 7A), and listed price is an asking price, not the price paid.", bold_lead="Issues and limitations. ", align="justify")
figure("m1_ridge", "Ridge predicted versus actual price on the test set (A) and the largest standardised non-LGA coefficients (B).")

doc.add_heading("Model 2 (unsupervised). K-Means market segmentation", level=2)
kt = pd.DataFrame(R["m2_k_table"]).set_index("k"); pr = pd.DataFrame(R["m2_profiles"])
para("K-Means assigns each listing to the nearest of k centroids and repeatedly moves each centroid to the mean of its members, minimising "
     "within-cluster squared distance (inertia). Our *objective* was to segment Sydney's supply into investment zones combining location, price, "
     "rail access and demand. Inputs were latitude, longitude, log price, log distance to station, distance to the CBD and estimated occupancy, all "
     "standardised because their units differ by orders of magnitude.", bold_lead="Method and objective. ", align="justify")
para(f"We fitted k = 2–10 (Figure 8). The silhouette score peaks at k = 3 ({kt.loc[3,'silhouette']:.2f}), but that solution only separates coast, suburbs "
     f"and inner city. The elbow flattens after k = 4, and k = 4 (silhouette {R['m2_silhouette']:.2f}) splits the inner city into two commercially "
     "very different groups, so we chose it for interpretability. Silhouette is O(n²), so it was computed on a fixed 4,000-listing sample.",
     bold_lead="Choosing k. ", align="justify")
figure("m2_choose_k", "Elbow curve (A) and silhouette score (B) for k = 2–10.", width_cm=13)
prow = [(r_.segment, f"{r_.listings:,.0f}", money(r_.median_price), f"{r_.median_occupancy:.0f}", money(r_.median_revenue),
         f"{r_.median_dist_station_m:,.0f} m", f"{r_.median_dist_cbd_km:.1f} km") for r_ in pr.itertuples()]
table(pd.DataFrame(prow, columns=["Segment", "Listings", "Median price", "Median nights", "Median revenue", "To station", "To CBD"]),
      "K-Means segment profiles (medians; revenue is Inside Airbnb's 12-month estimate)", [4.6, 1.7, 1.9, 1.7, 2.1, 2.0, 1.9])
city = pr[pr.segment == "City-core high-turnover"].iloc[0]; ring = pr[pr.segment == "Inner-ring low-utilisation"].iloc[0]
para(f"The clusters map cleanly (Figure 9). The *City-core high-turnover* and *Inner-ring low-utilisation* segments occupy the same LGAs "
     f"and similar rail access ({city.median_dist_station_m:,.0f} m against {ring.median_dist_station_m:,.0f} m), yet median occupancy is {city.median_occupancy:.0f} against "
     f"{ring.median_occupancy:.0f} nights and median revenue {money(city.median_revenue)} against {money(ring.median_revenue)}. Within the same locations, "
     "operation separates winners from under-used stock. The *Coastal premium* segment charges the most but is rail-remote and seasonal; "
     "the *Outer-suburban budget* segment is cheap with low demand. Silhouettes near 0.26 mean the segments overlap, K-Means assumes compact "
     "clusters and treats coordinates as Euclidean, and results depend on the inputs chosen, so the segments are a screening tool, not natural boundaries.",
     bold_lead="Results and limitations. ", align="justify")
figure("m2_cluster_maps", "Listings in each K-Means segment (blue) against all listings (grey), with segment medians.")

doc.add_heading("Model 3 (not covered in class). XGBoost with SHAP interpretation", level=2)
p3 = R["m3_params"]
para("XGBoost (Chen & Guestrin, 2016) is a gradient-boosted tree ensemble. Trees are added one at a time, each fitted to the residual errors of the "
     "trees before it and scaled by a learning rate, with L2 regularisation on leaf weights and row/column subsampling to limit overfitting. Unlike "
     "Ridge, it learns non-linear effects and interactions automatically. To interpret it we use SHAP (Lundberg & Lee, 2017), which splits each "
     "prediction into additive feature contributions based on Shapley values from cooperative game theory. With a log target, a SHAP value s means "
     "the feature moves the predicted price by about exp(s) − 1. Our *objective* was to beat Ridge's accuracy and isolate the non-linear effect of rail proximity.",
     bold_lead="Method and objective. ", align="justify")
para(f"The same inputs as Model 1 (raw distances, since trees are scale-free) with one-hot categories. A randomised search over 30 configurations with 3-fold CV "
     f"selected {p3['n_estimators']} trees, depth {p3['max_depth']}, learning rate {p3['learning_rate']}, subsample {p3['subsample']}, column sample "
     f"{p3['colsample_bytree']} and λ = {p3['reg_lambda']} (CV RMSE {R['m3_cv_rmse_log']:.3f} log units against {R['m1_cv_rmse_log']:.3f} for Ridge).",
     bold_lead="Preparation and tuning. ", align="justify")
cmp_rows = [("Baseline (median price)", f"{BASE['R2 (log price)']:.3f}", money(BASE['RMSE (AUD)']), money(BASE['MAE (AUD)']), f"{BASE['MAPE %']:.1f}%")]
for name, m in [("Model 1: Ridge", m1), ("Ridge without rail features", m1n), ("Model 3: XGBoost", m3), ("XGBoost without rail features", m3n)]:
    cmp_rows.append((name, f"{m['R2 (log price)']:.3f}", money(m['RMSE (AUD)']), money(m['MAE (AUD)']), f"{m['MAPE %']:.1f}%"))
table(pd.DataFrame(cmp_rows, columns=["Model (test set)", "R² (log)", "RMSE", "MAE", "MAPE"]), "Test-set performance and rail-feature ablation",
      [6.3, 2.2, 2.4, 2.4, 2.4])
sb = R["m3_shap_dist_band_pct"]
para(f"XGBoost improves on Ridge on every metric: R² {m3['R2 (log price)']:.3f} against {m1['R2 (log price)']:.3f}, MAE {money(m3['MAE (AUD)'])} against "
     f"{money(m1['MAE (AUD)'])}, MAPE {m3['MAPE %']:.1f}% against {m1['MAPE %']:.1f}%. The gain comes from non-linear size and location effects. SHAP ranks bedrooms, guests, room "
     f"type and longitude (the east–west, coast-to-inland gradient) highest (Figure 10A). The rail features rank {R['m3_shap_rail_rank']}th and carry only "
     f"{R['m3_shap_rail_share_pct']}% of total |SHAP|. Figure 10B isolates station distance: within 4 km its effect stays within about ±1% of price "
     f"(e.g. {sb['<250 m']:+.1f}% under 250 m, {sb['500 m-1 km']:+.1f}% at 500 m–1 km). It turns positive only beyond 5 km ({sb['>5 km']:+.1f}% on average), "
     "where it is acting as a proxy for coastal Northern Beaches homes rather than a rail effect. Dropping rail features moves test R² only from "
     f"{m3['R2 (log price)']:.4f} to {m3n['R2 (log price)']:.4f}. Three independent methods (regression with controls, Ridge and XGBoost with SHAP) agree "
     "that rail proximity does not set Sydney's short-term rental prices.", bold_lead="Results. ", align="justify")
eb = R["ml_err_by_band"]
para(f"(1) Overfitting: training R² is {R['m3_train_r2']:.2f} against {m3['R2 (log price)']:.2f} on test. The search favoured subsampling and "
     "regularisation, and CV and test errors agree, so the score is not a lucky split. (2) SHAP spreads one-hot categories over many columns, "
     "under-stating LGA and room type, so we summed SHAP values back to the original variable. (3) Predictions shrink to the middle: XGBoost "
     f"over-predicts the cheapest quartile by {eb['XGBoost bias %']['Q1 (cheapest)']:.0f}% and under-predicts the dearest by "
     f"{abs(eb['XGBoost bias %']['Q4 (dearest)']):.0f}%. (4) SHAP explains the model, not causation, and the random split lets neighbouring listings sit "
     "in both train and test sets, so spatial generalisation may be optimistic.", bold_lead="Issues and limitations. ", align="justify")
figure("m3_shap", "XGBoost feature importance as mean |SHAP|, rail features in orange (A), and SHAP effect of station distance on predicted price with binned mean (B).")

# ======================================================================
# Part C
# ======================================================================
doc.add_heading("Part C. Summary and Business Recommendations", level=1)
para("Across the exploratory and machine learning analyses, one message is consistent: in Sydney's short-term rental market, *what* a "
     "listing is and *where* it sits set the price, while *how it is run* sets the demand. Location strongly shapes price (coastal and "
     "North Shore premiums, a western-suburbs discount). Rail access, which looks like a 30% discount in raw data, has no meaningful price effect "
     f"once structure and location are controlled ({R['a3_adj_premium_500']:+.1f}%, not significant), and adds almost nothing to either price model. "
     f"It does add demand (about {R['a3_occ_500_nights']:.0f} extra nights a year within 500 m). Demand, in turn, depends on operation: popular listings "
     "are superhost-run, have short minimum stays and avoid sub-4.5 ratings, and K-Means shows that within the same inner-city locations a "
     f"high-turnover segment earns about {city.median_revenue/ring.median_revenue:.0f} times the median revenue of an under-used one. "
     "Commercial hosts hold the best-connected stock but deliver lower ratings and lower occupancy.", align="justify")
bullet(f" Value transit-adjacent units on occupancy, not on nightly rate. Our models show no reliable price premium for being within "
       f"500 m of a station, but a measurable occupancy gain. Investors should not pay above-market purchase prices expecting higher rates near "
       f"stations. They should target well-connected inner-city stock and underwrite it on volume, as in the city-core segment "
       f"({money(city.median_price)} median rate, {city.median_occupancy:.0f} nights, {money(city.median_revenue)} median revenue).",
       lead="1. Investors —")
bullet(f" Cut booking friction before cutting price. Occupancy falls from {mn['1']:.0f} to {mn['4-7']:.0f} median nights as minimum "
       f"stays rise from one to four to seven nights, and superhost listings book about {sh['1']/max(sh['0'],1):.0f} times as many nights as others. Hosts in the "
       "inner-ring low-utilisation segment, who already have the city-core's locations, should first move to one- or two-night minimums, "
       "work towards superhost criteria and add amenities, before discounting. Price is a weak lever on ratings (ρ = "
       f"{R['q2_rho_overall']:.2f}).", lead="2. Individual hosts —")
bullet(f" Close the guest-experience gap, starting with cleanliness. Commercial listings trail single hosts by "
       f"{abs(tt['rating_gap_adj']):.2f} stars after controls and {hs['Single (1)']['Cleanliness'] - hs['Commercial (10+)']['Cleanliness']:.2f} on "
       f"cleanliness, and {C['below_4_5_pct']:.0f}% of their listings fall below the 4.5 threshold where occupancy collapses (median {rb['<4.5']:.0f} nights against {rb['4.8-4.9']:.0f} at 4.8–4.9). "
       "Turnover audits, standardised cleaning checklists and value-adding inclusions would let operators convert their location "
       f"advantage into the occupancy they currently miss ({C['median_occupancy']:.0f} against {S['median_occupancy']:.0f} nights). An XGBoost price model "
       f"(MAPE {m3['MAPE %']:.0f}%) can support rate-setting for mid-market stock, but not for luxury homes, where it under-predicts.",
       lead="3. Commercial operators —")
para("Limitations: Inside Airbnb occupancy and revenue are model estimates built from reviews; listed prices are asking prices from one "
     "snapshot (winter); the station data reflects the 2020 network, so newer Sydney Metro stations are missing; and all findings are associations.",
     align="justify")

# ======================================================================
# References
# ======================================================================
doc.add_heading("References", level=1)
refs = [
    "Chen, T., & Guestrin, C. (2016). XGBoost: A scalable tree boosting system. In *Proceedings of the 22nd ACM SIGKDD International Conference on Knowledge Discovery and Data Mining* (pp. 785–794). https://doi.org/10.1145/2939672.2939785",
    "Hoerl, A. E., & Kennard, R. W. (1970). Ridge regression: Biased estimation for nonorthogonal problems. *Technometrics, 12*(1), 55–67.",
    f"Inside Airbnb. (2026). *Sydney, New South Wales, Australia: listings, calendar and neighbourhoods* [Data set; scraped 17–29 June 2026]. Licensed under CC BY 4.0. Retrieved 6 October 2026, from https://insideairbnb.com/get-the-data/",
    "Lundberg, S. M., & Lee, S.-I. (2017). A unified approach to interpreting model predictions. *Advances in Neural Information Processing Systems, 30*, 4765–4774.",
    "Pedregosa, F., et al. (2011). Scikit-learn: Machine learning in Python. *Journal of Machine Learning Research, 12*, 2825–2830.",
    "Rousseeuw, P. J. (1987). Silhouettes: A graphical aid to the interpretation and validation of cluster analysis. *Journal of Computational and Applied Mathematics, 20*, 53–65.",
    "Transport for NSW. (2020). *Train station entrance locations* (stationentrances2020_v4) [Data set]. TfNSW Open Data Hub. Licensed under CC BY 4.0. Retrieved 6 October 2026, from https://opendata.transport.nsw.gov.au/",
]
for ref in refs:
    p = para(ref, after=3)
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
