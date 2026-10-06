"""Download the OpenStreetMap inputs the notebook needs (run once from the project root).

1. Points of interest (beaches and tourist attractions) inside the Sydney listings window
   -> data/osm/osm_pois_raw.csv (cleaned into External04_Group09.csv by the notebook)
2. The pedestrian street network within 2.5 km of every rail station
   -> data/osm/sydney_walk.graphml (used for walking distances to station entrances)

OpenStreetMap data (c) OpenStreetMap contributors, ODbL 1.0. Queried through the Overpass API with OSMnx.
"""
from pathlib import Path

import geopandas as gpd
import numpy as np
import osmnx as ox
import pandas as pd
from shapely.geometry import box

ROOT = Path(__file__).resolve().parent.parent
OSM = ROOT / "data" / "osm"
OSM.mkdir(parents=True, exist_ok=True)
ox.settings.use_cache = True
ox.settings.cache_folder = str(OSM / "cache")
ox.settings.requests_timeout = 600
ox.settings.max_query_area_size = 2_500_000_000  # fewer, larger Overpass queries

listings = pd.read_csv(ROOT / "data" / "listings.csv.gz", usecols=["latitude", "longitude"])
window = box(listings["longitude"].min() - 0.05, listings["latitude"].min() - 0.05,
             listings["longitude"].max() + 0.05, listings["latitude"].max() + 0.05)

# ---- 1. Points of interest ----
poi_path = OSM / "osm_pois_raw.csv"
if not poi_path.exists():
    tags = {"natural": "beach", "tourism": ["attraction", "museum", "zoo", "aquarium", "theme_park", "gallery"]}
    feats = ox.features_from_polygon(window, tags)
    feats = feats.reset_index()
    pts = feats.to_crs(7856)
    pts["geometry"] = pts.geometry.representative_point()
    pts = pts.to_crs(4326)
    out = pd.DataFrame({
        "osm_type": feats["element"] if "element" in feats else feats.get("element_type"),
        "osm_id": feats["id"] if "id" in feats else feats.get("osmid"),
        "name": feats.get("name"),
        "natural": feats.get("natural"),
        "tourism": feats.get("tourism"),
        "wikidata": feats.get("wikidata"),
        "lat": pts.geometry.y.values,
        "lon": pts.geometry.x.values,
    })
    out.to_csv(poi_path, index=False)
    print("POIs:", len(out))

# ---- 2. Walking network around stations ----
graph_path = OSM / "sydney_walk.graphml"
if not graph_path.exists():
    stops = pd.read_csv(ROOT / "tfNSW GTFS" / "stops.txt", dtype=str)
    stops = stops[stops["location_type"] == "1"].copy()
    stops = stops[stops["stop_name"].str.contains(r"Station$|Light Rail$", regex=True)]
    stops[["stop_lat", "stop_lon"]] = stops[["stop_lat", "stop_lon"]].astype(float)
    g = gpd.GeoDataFrame(stops, geometry=gpd.points_from_xy(stops["stop_lon"], stops["stop_lat"]), crs=4326)
    g = g[g.within(window)]
    area = g.to_crs(7856).buffer(2500).union_all()
    area = gpd.GeoSeries([area], crs=7856).to_crs(4326).iloc[0].intersection(window)
    print("Stations in window:", len(g))
    G = ox.graph_from_polygon(area, network_type="walk", simplify=True, retain_all=False)
    ox.save_graphml(G, graph_path)
    print("Walk graph:", G.number_of_nodes(), "nodes,", G.number_of_edges(), "edges")
