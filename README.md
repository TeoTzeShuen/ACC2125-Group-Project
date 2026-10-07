# Does Rail Access Pay? Machine Learning on Sydney's Short-Term Rental Market

## Goal of the analysis

Investors often assume that short-term rental listings near train stations earn more. This project tests that assumption on Sydney, Australia. It joins the Inside Airbnb listings for Sydney to the current Transport for NSW rail, Metro and light-rail network (station entrances, frequency, travel time to the CBD and airport, patronage), measures real walking distance on the OpenStreetMap street network, and controls for beaches, attractions and neighbourhood status (ABS SEIFA by SA2). It then asks:

1. **Who gets booked?** What characterises the most popular listings?
2. **Does price buy ratings?** How are price and guest ratings related?
3. **Where does location matter?** How do neighbourhood and rail access relate to price, demand and reviews?
4. **Is there a rail-access price premium?** Does it survive controls for listing type and location?
5. **Do commercial hosts differ?** Do hosts with 10+ listings differ from single-listing hosts in ratings and rail access?
6. **Can we predict price, and what does rail access add?** Three machine learning models: Ridge regression, K-Means market segmentation, and XGBoost with SHAP interpretation.

### Headline findings

- Listings within a 500 m walk of a station look about 26% cheaper in raw data. After controlling for listing structure, beaches, attractions and SA2 neighbourhood (standard errors clustered by host), the difference is +1.7% (95% CI −3.0% to +6.6%), not significant. Rail proximity also adds no measurable demand (−1.1 nights a year, 95% CI −8.8 to +6.7) or revenue (+2.2%, 95% CI −9.2% to +14.9%).
- The average hides a split: near-station listings earn a premium in the middle ring (+4.6%) and outer suburbs (+6.8%), but not in the walkable inner city (−3.3%).
- The first draft's measure (straight-line distance to the outdated 2020 station file) produced a spurious significant +5.8% premium in the same model.
- Popular listings are not the expensive ones: they are more often run by superhosts and have shorter minimum stays, although part of the superhost link is mechanical (both measures count reviews) and disappears under a calendar-based definition.
- Commercial hosts hold the best-connected stock but rate about 0.16 stars lower and book about 21 fewer nights a year than comparable single-host listings in the same SA2.
- On held-out SA2 neighbourhoods XGBoost (R² 0.79) beats Ridge (0.72) at predicting log price; a random split overstates both (0.81 and 0.73). Rail features add almost nothing (Ridge +0.010 R², XGBoost none).

## Repository contents

| Path | Description |
|---|---|
| `Code_Group09.ipynb` | The full analysis notebook (Parts A1–A3 and B). Includes the Data Decisions Log. |
| `External01_Group09.csv` | Rail, Metro and light-rail station entrances with mode, off-peak frequency and rail travel time to the CBD and airport, derived from the TfNSW GTFS timetable by the notebook. |
| `External02_Group09.csv` | TfNSW station entries and exits (monthly, Oct 2024 – Aug 2026). |
| `External03_Group09.csv` | ABS SEIFA 2021 indexes for NSW SA2s (extracted by the notebook from the ABS workbook). |
| `External04_Group09.csv` | OpenStreetMap beaches and tourist attractions in the Sydney window. |
| `External05_Group09.csv` | TfNSW train station entrances (2020, v4), kept as a comparison with the current network. |
| `Data01_Group09.csv` | Cleaned listings with engineered rail features (written by the notebook). |
| `Data02_Group09.csv` | 90-day forward calendar availability per listing (written by the notebook). |
| `figures/` | Charts at 300 DPI (written by the notebook). |
| `outputs/` | Data Decisions Log, model comparison and `results.json` with every headline number (written by the notebook). |
| `scripts/` | Build scripts. `build_notebook.py` regenerates and executes the notebook from `notebook_source.txt`; `fetch_osm.py` downloads the OpenStreetMap inputs; `build_report.py` and `build_slides.py` build the Word report and slides from the notebook outputs (their text still describes the first draft and needs updating). |

## Data you need to download

The notebook reads raw files from a `data/` folder in the project root, plus the TfNSW GTFS feed in `tfNSW GTFS/`. None of these are included in this repository (they are large or redistributable only from source). The submitted `External0X_Group09.csv` files are included, and once they exist the notebook reads them instead of rebuilding them, so the GTFS feed and the SEIFA workbook are needed only to rebuild External01 and External03.

### 1. Inside Airbnb: Sydney (required)

- **Source:** <https://insideairbnb.com/get-the-data/> (Australia → New South Wales → Sydney)
- **Snapshot used:** listings scraped between 17 and 29 June 2026
- **Files to download and place in `data/`:**

| File | Used for |
|---|---|
| `listings.csv.gz` | Listing attributes, price, reviews, host and location fields (the detailed file, not the summary `listings.csv`) |
| `calendar.csv.gz` | Forward availability (90 days after the snapshot) |
| `neighbourhoods.geojson` | LGA boundaries for the maps |

`reviews.csv.gz`, `reviews.csv`, `listings.csv` (summary) and `neighbourhoods.csv` are not used by the notebook.

For the cross-season check, also download the **21 March 2026** `listings.csv.gz` and save it as `data/snapshots/listings_2026-03-21.csv.gz`. (The September and December 2025 Sydney snapshots have no prices and are not used.)

Inside Airbnb publishes new snapshots over time and removes older ones from its main page, so the June 2026 snapshot may not be the one currently listed. If you download a different snapshot, the numbers, figures and findings will change. Check the data archive on the same site for earlier snapshots.

### 2. Transport for NSW: GTFS timetable (only to rebuild External01)

*Timetables Complete GTFS* from the TfNSW Open Data Hub (<https://opendata.transport.nsw.gov.au/data/dataset/timetables-complete-gtfs>; free API key required). Unzip into `tfNSW GTFS/`. The notebook uses the timetable for Wednesday 7 October 2026.

### 3. ABS: SA2 boundaries and SEIFA (required / only to rebuild External03)

- SA2 boundaries, ASGS Edition 3 (2021), GDA2020 shapefile `SA2_2021_AUST_SHP_GDA2020.zip`, unzipped into `data/abs/`: <https://www.abs.gov.au/statistics/standards/australian-statistical-geography-standard-asgs-edition-3/jul2021-jun2026/access-and-downloads/digital-boundary-files>
- *Statistical Area Level 2, Indexes, SEIFA 2021.xlsx* in `data/abs/`: <https://www.abs.gov.au/statistics/people/people-and-communities/socio-economic-indexes-areas-seifa-australia/latest-release>

### 4. OpenStreetMap (required)

Run `python scripts/fetch_osm.py` once from the project root. It downloads beaches and attractions (`data/osm/osm_pois_raw.csv`) and the pedestrian network within 2.5 km of every station (`data/osm/sydney_walk.graphml`, about 300 MB) through the Overpass API, which takes roughly 10 minutes.

### Expected folder layout

```
.
├── Code_Group09.ipynb
├── External01_Group09.csv … External05_Group09.csv
├── tfNSW GTFS/                      (only to rebuild External01)
└── data/
    ├── listings.csv.gz
    ├── calendar.csv.gz
    ├── neighbourhoods.geojson
    ├── snapshots/listings_2026-03-21.csv.gz
    ├── abs/SA2_2021_AUST_GDA2020.shp (+ .dbf/.shx/.prj)
    ├── abs/Statistical Area Level 2, Indexes, SEIFA 2021.xlsx
    └── osm/ (written by scripts/fetch_osm.py)
```

## Running the notebook

Tested with **Python 3.14** and the following libraries:

```
pip install pandas numpy scipy matplotlib seaborn geopandas statsmodels scikit-learn xgboost shap ipykernel openpyxl networkx osmnx
```

(`osmnx` is only needed for `scripts/fetch_osm.py`; `python-docx`, `python-pptx`, `nbformat` and `nbclient` only for the other build scripts.)

Open `Code_Group09.ipynb` in Jupyter or VS Code with a Python 3.14 kernel and run all cells from the project folder. A full run takes about 6 minutes (the two XGBoost searches are the slowest step; the first run also converts the walking network to a compact cache). The notebook creates `figures/` and `outputs/` and overwrites `Data01_Group09.csv` and `Data02_Group09.csv`. Random seeds are fixed (42) so results are reproducible with the same data and library versions.

## Attributions and licences

- **Inside Airbnb.** Sydney, New South Wales, Australia: listings, calendar and neighbourhoods. Compiled from public Airbnb information by Inside Airbnb, <https://insideairbnb.com/>. Licensed under [Creative Commons Attribution 4.0 International (CC BY 4.0)](https://creativecommons.org/licenses/by/4.0/). Inside Airbnb data is independent of, and not affiliated with, Airbnb.
- **Transport for NSW.** *Timetables Complete GTFS* (timetable valid from 6 October 2026); *Train, Metro and Light Rail Station Entries and Exits* (October 2024 – August 2026); *Train station entrance locations* (2020, v4). TfNSW Open Data Hub, <https://opendata.transport.nsw.gov.au/>. Licensed under CC BY 4.0.
- **Australian Bureau of Statistics.** *Socio-Economic Indexes for Areas (SEIFA), Australia, 2021* (SA2 indexes, released 27 April 2023) and *Australian Statistical Geography Standard (ASGS) Edition 3*, SA2 2021 digital boundaries. Licensed under CC BY 4.0.
- **OpenStreetMap.** Beaches, tourist attractions and the pedestrian network © OpenStreetMap contributors, available under the Open Database Licence (ODbL 1.0), <https://www.openstreetmap.org/copyright>; downloaded via the Overpass API with OSMnx (Boeing, 2017) in October 2026.
- Methods: Dibbelt et al. (2013), connection scan algorithm; Santos Silva & Tenreyro (2006), PPML; Chen & Guestrin (2016), XGBoost; Lundberg & Lee (2017), SHAP; Hoerl & Kennard (1970), Ridge regression; Rousseeuw (1987), silhouette score; Pedregosa et al. (2011), scikit-learn.

## Limitations

Inside Airbnb's occupancy and revenue figures are model estimates built from review counts, and the calendar mixes bookings with host-blocked nights. Prices are asking prices; the main snapshot is from winter (June 2026), with March 2026 used as a cross-season check. The GTFS timetable (October 2026) is a few months later than the listings snapshot. SEIFA is from the 2021 Census. All findings are associations, not causal effects.
