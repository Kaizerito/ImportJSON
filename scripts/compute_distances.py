"""Compute driving distances from Arc-et-Senans to a list of communes.

This script uses the OpenRouteService directions API to geocode and compute
routes to a predefined list of communes around Arc-et-Senans, France. The
results are exported to both Excel and CSV files in the current working
directory.

Usage:
    python scripts/compute_distances.py --api-key <ORS_API_KEY>

The API key can also be provided through the environment variable
``ORS_API_KEY``. If neither the flag nor the environment variable is set, the
script will exit with an error message.

The script rate-limits requests to avoid exceeding the OpenRouteService usage
limits. Depending on the response times of the external service, the full run
can take several minutes.
"""

from __future__ import annotations

import argparse
import os
import time
from typing import Iterable, List, Tuple

import openrouteservice
import pandas as pd

# Arc-et-Senans coordinates (lat, lon)
ARC_LAT = 47.03305410997385
ARC_LON = 5.777678186344349

# (Postal code, Commune name)
COMMUNES: List[Tuple[str, str]] = [
    ("25001", "Abbans-Dessous"),
    ("25084", "Boussières"),
    ("25103", "Busy"),
    ("25105", "Byans-sur-Doubs"),
    ("25147", "Chemaudin et Vaux"),
    ("25287", "Grandfontaine"),
    ("25397", "Montferrand-le-Château"),
    ("25438", "Osselle-Routelle"),
    ("25477", "Rancenay"),
    ("25561", "Thoraise"),
    ("25564", "Torpes"),
    ("25631", "Vorges-les-Pins"),
    ("25002", "Abbans-Dessus"),
    ("25044", "Bartherans"),
    ("25090", "Brères"),
    ("25098", "Buffard"),
    ("25104", "By"),
    ("25109", "Cessey"),
    ("25126", "Charnay"),
    ("25143", "Chay"),
    ("25149", "Chenecey-Buillon"),
    ("25154", "Chouzelot"),
    ("25171", "Courcelles"),
    ("25185", "Cussey-sur-Lison"),
    ("25209", "Fourg"),
    ("25253", "Goux-sous-Landet"),
    ("25283", "Lavans-Quingey"),
    ("25330", "Le Val"),
    ("25336", "Liesle"),
    ("25340", "Lombard"),
    ("25379", "Mesmay"),
    ("25416", "Myon"),
    ("25443", "Palantine"),
    ("25445", "Paroy"),
    ("25450", "Pessans"),
    ("25460", "Quingey"),
    ("25475", "Rennes-sur-Loue"),
    ("25488", "Ronchaux"),
    ("25500", "Rouhe"),
    ("25507", "Samson"),
    ("25528", "Échay"),
    ("25021", "Arc-et-Senans"),
    ("39370", "Mouchard"),
    ("39403", "Pagnoz"),
    ("39026", "Augerans"),
    ("39037", "Bans"),
    ("39048", "Belmont"),
    ("39093", "Chamblay"),
    ("39117", "Chatelay"),
    ("39149", "Chissey-sur-Loue"),
    ("39249", "Germigney"),
    ("39305", "La Loye"),
    ("39350", "La Vieille-Loye"),
    ("39365", "Mont-sous-Vaudrey"),
    ("39387", "Montbarrey"),
    ("39399", "Nevy-lès-Dole"),
    ("39502", "Ounans"),
    ("39520", "Santans"),
    ("39546", "Souvans"),
    ("39559", "Vaudrey"),
    ("39002", "Abergement-le-Grand"),
    ("39013", "Arbois"),
    ("39019", "Champagne-sur-Loue"),
    ("39095", "Cramans"),
    ("39116", "Grange-de-Vaivre"),
    ("39176", "La Châtelaine"),
    ("39206", "La Ferté"),
    ("39223", "Les Arsures"),
    ("39259", "Les Planches-près-Arbois"),
    ("39319", "Mathenay"),
    ("39325", "Mesnay"),
    ("39337", "Molamboz"),
    ("39355", "Montigny-lès-Arsures"),
    ("39425", "Port-Lesney"),
    ("39439", "Pupillin"),
    ("39446", "Saint-Cyr-Montmalin"),
    ("39479", "Vadans"),
    ("39539", "Villeneuve-d'Aval"),
    ("39565", "Villers-Farlay"),
    ("39569", "Villette-lès-Arbois"),
    ("39572", "Écleux"),
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--api-key",
        dest="api_key",
        default=None,
        help="OpenRouteService API key. Overrides the ORS_API_KEY environment variable.",
    )
    parser.add_argument(
        "--sleep",
        dest="sleep",
        type=float,
        default=0.8,
        help="Delay (in seconds) between API calls to avoid hitting rate limits.",
    )
    parser.add_argument(
        "--output-prefix",
        dest="output_prefix",
        default="distances_arc_et_senans",
        help="Base filename for the generated CSV and XLSX files.",
    )
    return parser.parse_args()


def resolve_api_key(cli_key: str | None) -> str:
    api_key = cli_key or os.getenv("ORS_API_KEY")
    if not api_key:
        raise SystemExit(
            "An OpenRouteService API key is required. Provide it via --api-key or the ORS_API_KEY environment variable."
        )
    return api_key


def init_client(api_key: str) -> openrouteservice.Client:
    return openrouteservice.Client(key=api_key)


def geocode_commune(
    client: openrouteservice.Client, commune: Tuple[str, str]
) -> Tuple[str, str, float | None, float | None]:
    postal_code, name = commune
    query = f"{name}, {postal_code}, France"
    response = client.pelias_search(text=query)
    if response.get("features"):
        coordinates = response["features"][0]["geometry"]["coordinates"]
        return postal_code, name, coordinates[1], coordinates[0]
    return postal_code, name, None, None


def compute_route(
    client: openrouteservice.Client, lat: float, lon: float
) -> Tuple[float, float]:
    route = client.directions(
        coordinates=[[ARC_LON, ARC_LAT], [lon, lat]],
        profile="driving-car",
        format="json",
    )
    summary = route["routes"][0]["summary"]
    distance_km = summary["distance"] / 1000
    duration_min = summary["duration"] / 60
    return distance_km, duration_min


def collect_distances(
    client: openrouteservice.Client,
    communes: Iterable[Tuple[str, str]],
    sleep_seconds: float,
) -> List[dict]:
    results = []
    for postal_code, name in communes:
        print(f"Traitement : {name} ({postal_code})...")
        _, _, lat, lon = geocode_commune(client, (postal_code, name))
        if lat is None or lon is None:
            print("  → Impossible de trouver la localisation, IGNORÉ")
            continue
        distance_km, duration_min = compute_route(client, lat, lon)
        results.append(
            {
                "Code postal": postal_code,
                "Commune": name,
                "Distance Arc-et-Senans (km)": round(distance_km, 2),
                "Temps voiture (min)": round(duration_min, 1),
            }
        )
        time.sleep(max(0.0, sleep_seconds))
    return results


def export_results(data: List[dict], output_prefix: str) -> None:
    df = pd.DataFrame(data)
    csv_path = f"{output_prefix}.csv"
    xlsx_path = f"{output_prefix}.xlsx"
    df.to_csv(csv_path, index=False)
    df.to_excel(xlsx_path, index=False)
    print(f"Fichiers générés : {xlsx_path} + {csv_path}")


def main() -> None:
    args = parse_args()
    api_key = resolve_api_key(args.api_key)
    client = init_client(api_key)
    results = collect_distances(client, COMMUNES, args.sleep)
    if not results:
        print("Aucun résultat calculé.")
        return
    export_results(results, args.output_prefix)


if __name__ == "__main__":
    main()
