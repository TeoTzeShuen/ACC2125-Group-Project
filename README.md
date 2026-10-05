# Does Rail Access Pay? Machine Learning on Sydney's Short-Term Rental Market

Group project for **ACC2125 Data Analytics and Machine Learning** (Singapore Institute of Technology, AY2026/27 Trimester 1).

## Goal of the analysis

Investors often assume that short-term rental listings near train stations earn more. This project tests that assumption on Sydney, Australia. It joins the Inside Airbnb listings for Sydney to the location of every Transport for NSW train station entrance, then asks:

1. **Who gets booked?** What characterises the most popular listings?
2. **Does price buy ratings?** How are price and guest ratings related?
3. **Where does location matter?** How do neighbourhood and rail access relate to price, demand and reviews?
4. **Is there a rail-access price premium?** Does it survive controls for listing type and location?
5. **Do commercial hosts differ?** Do hosts with 10+ listings differ from single-listing hosts in ratings and rail access?
6. **Can we predict price, and what does rail access add?** Three machine learning models: Ridge regression, K-Means market segmentation, and XGBoost with SHAP interpretation.

### Headline findings

- Listings within 500 m of a station look about 30% cheaper in raw data. After controlling for room type, size, distance to the CBD and local government area (LGA), the difference is +2.3% (95% CI −0.3% to +5.0%) and is not statistically significant.
- Rail proximity is associated with about 4.8 more booked nights a year (95% CI 0.3–9.3).
- Popular listings are not the expensive ones: they are more often run by superhosts and have shorter minimum stays.
- Commercial hosts hold the best-connected stock but rate about 0.17 stars lower after controls, and book fewer nights.
- XGBoost (test R² 0.81) beats Ridge (0.73) at predicting log price. Removing the rail features changes R² by less than 0.002 in both.

## Repository contents

| Path | Description |
|---|---|
| `Code_GroupXX.ipynb` | The full analysis notebook (Parts A1–A3 and B). Includes the Data Decisions Log. |
| `External_GroupXX.csv` | Transport for NSW train station entrances (see attributions). |
| `Data01_GroupXX.csv` | Cleaned listings with engineered rail features (written by the notebook). |
| `Data02_GroupXX.csv` | 90-day forward calendar availability per listing (written by the notebook). |
| `figures/` | Charts at 300 DPI (written by the notebook). |
| `outputs/` | Data Decisions Log, model comparison and `results.json` with every headline number (written by the notebook). |
| `scripts/` | Build scripts. `build_notebook.py` regenerates and executes the notebook from `notebook_source.txt`; `build_report.py` and `build_slides.py` build the Word report and slides from the notebook outputs. |

## Data you need to download

The notebook reads three raw files from a `data/` folder in the project root. The raw Inside Airbnb files are **not** included in this repository.

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

Inside Airbnb publishes new snapshots over time and removes older ones from its main page, so the June 2026 snapshot may not be the one currently listed. If you download a different snapshot, the numbers, figures and findings will change. Check the data archive on the same site for earlier snapshots.

### 2. Transport for NSW: train station entrances (included)

`External_GroupXX.csv` is included in this repository, so no download is needed to run the notebook. It is the *Train Station Entrance Locations* dataset (file `stationentrances2020_v4.csv`) from the Transport for NSW Open Data Hub: <https://opendata.transport.nsw.gov.au/>. It lists the latitude and longitude of every entrance to Sydney Trains, NSW TrainLink, Sydney Metro and light-rail stations, as at 2020. Newer stations opened since 2020 are not included.

### Expected folder layout

```
.
├── Code_GroupXX.ipynb
├── External_GroupXX.csv
└── data/
    ├── listings.csv.gz
    ├── calendar.csv.gz
    └── neighbourhoods.geojson
```

## Running the notebook

Tested with **Python 3.14** and the following libraries:

```
pip install pandas numpy scipy matplotlib seaborn geopandas statsmodels scikit-learn xgboost shap ipykernel
```

(`python-docx`, `python-pptx`, `nbformat` and `nbclient` are only needed to run the build scripts in `scripts/`.)

Open `Code_GroupXX.ipynb` in Jupyter or VS Code with a Python 3.14 kernel and run all cells from the project folder. A full run takes about 3–4 minutes (the XGBoost search is the slowest step). The notebook creates `figures/` and `outputs/` and overwrites `Data01_GroupXX.csv` and `Data02_GroupXX.csv`. Random seeds are fixed (42) so results are reproducible with the same data and library versions.

## Attributions and licences

- **Inside Airbnb.** Sydney, New South Wales, Australia: listings, calendar and neighbourhoods. Compiled from public Airbnb information by Inside Airbnb, <https://insideairbnb.com/>. Licensed under [Creative Commons Attribution 4.0 International (CC BY 4.0)](https://creativecommons.org/licenses/by/4.0/). Inside Airbnb data is independent of, and not affiliated with, Airbnb.
- **Transport for NSW.** *Train station entrance locations* (2020, v4), TfNSW Open Data Hub, <https://opendata.transport.nsw.gov.au/>. Licensed under CC BY 4.0.
- Methods: Chen & Guestrin (2016), XGBoost; Lundberg & Lee (2017), SHAP; Hoerl & Kennard (1970), Ridge regression; Rousseeuw (1987), silhouette score; Pedregosa et al. (2011), scikit-learn.

## Limitations

Inside Airbnb's occupancy and revenue figures are model estimates built from review counts, and the calendar mixes bookings with host-blocked nights. Prices are asking prices from a single (winter) snapshot. The station data reflects the 2020 network. All findings are associations, not causal effects.

## Academic use

This was prepared for coursework at the Singapore Institute of Technology. If you are a student on this module, follow your institution's academic integrity policy before reusing any of it.
