"""Compute driving distances from Arc-et-Senans to a list of communes.

This script uses the OpenRouteService directions API to geocode and compute
routes to a predefined list of communes around Arc-et-Senans, France. The
results are exported to both Excel and CSV files in the current working
directory.

Quick start:
1. Installez les dépendances : ``pip install openrouteservice pandas``.
2. Renseignez votre clé dans l'une des options suivantes :
   - ``python scripts/compute_distances.py --api-key VOTRE_CLE``
   - ``export ORS_API_KEY=VOTRE_CLE`` puis exécutez le script
   - placez la clé dans un fichier texte (ou ``.env`` contenant ``ORS_API_KEY=VOTRE_CLE``)
     et utilisez ``--api-key-file chemin/vers/fichier``

Si aucune clé n'est trouvée via ces méthodes, le script affichera un message
explicatif et s'arrêtera.

The script rate-limits requests to avoid exceeding the OpenRouteService usage
limits. Depending on the response times of the external service, the full run
can take several minutes.
"""

from __future__ import annotations

import argparse
import os
import time
from pathlib import Path
from typing import Iterable, List, Tuple, TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover - uniquement pour l'analyse statique
    import openrouteservice

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
        "--api-key-file",
        dest="api_key_file",
        default=None,
        help="Path to a text file (or .env file) that contains an ORS_API_KEY entry.",
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


def resolve_api_key(cli_key: str | None, api_key_file: str | None) -> str:
    """Resolve the OpenRouteService API key from multiple locations."""

    if cli_key:
        return cli_key.strip()

    if api_key_file:
        api_key = read_api_key_from_file(Path(api_key_file))
        if api_key:
            return api_key

    env_key = os.getenv("ORS_API_KEY")
    if env_key:
        return env_key.strip()

    dotenv_key = read_api_key_from_dotenv(Path(".env"))
    if dotenv_key:
        return dotenv_key

    raise SystemExit(
        "An OpenRouteService API key is required. Provide it via --api-key, --api-key-file, "
        "the ORS_API_KEY environment variable, or by adding ORS_API_KEY=<votre_cle> to a local .env file."
    )


def read_api_key_from_file(path: Path) -> str | None:
    """Return the API key stored in *path* if it exists and is readable."""

    try:
        contents = path.read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        print(f"  → Fichier de clé API introuvable : {path}")
        return None
    except OSError as exc:
        raise SystemExit(f"Impossible de lire le fichier de clé API '{path}': {exc}") from exc

    if not contents:
        print(f"  → Le fichier de clé API '{path}' est vide.")
        return None
    if "=" not in contents:
        return contents

    # Autorise un fichier .env minimal contenant ORS_API_KEY=<valeur>
    for line in contents.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        if key.strip() == "ORS_API_KEY":
            return value.strip().strip('"').strip("'")
    return None


def read_api_key_from_dotenv(path: Path) -> str | None:
    """Extract ORS_API_KEY from a local .env file, if present."""

    if not path.exists():
        return None

    try:
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            if key.strip() == "ORS_API_KEY":
                return value.strip().strip('"').strip("'")
    except OSError as exc:
        raise SystemExit(f"Impossible de lire le fichier .env '{path}': {exc}") from exc
    return None


def init_client(api_key: str) -> "openrouteservice.Client":
    import openrouteservice

    return openrouteservice.Client(key=api_key)


def geocode_commune(
    client: "openrouteservice.Client", commune: Tuple[str, str]
) -> Tuple[str, str, float | None, float | None]:
    postal_code, name = commune
    query = f"{name}, {postal_code}, France"
    response = client.pelias_search(text=query)
    if response.get("features"):
        coordinates = response["features"][0]["geometry"]["coordinates"]
        return postal_code, name, coordinates[1], coordinates[0]
    return postal_code, name, None, None


def compute_route(
    client: "openrouteservice.Client", lat: float, lon: float
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
    import pandas as pd

    df = pd.DataFrame(data)
    csv_path = f"{output_prefix}.csv"
    xlsx_path = f"{output_prefix}.xlsx"
    df.to_csv(csv_path, index=False)
    df.to_excel(xlsx_path, index=False)
    print(f"Fichiers générés : {xlsx_path} + {csv_path}")


def main() -> None:
    args = parse_args()
    api_key = resolve_api_key(args.api_key, args.api_key_file)
    client = init_client(api_key)
    results = collect_distances(client, COMMUNES, args.sleep)
    if not results:
        print("Aucun résultat calculé.")
        return
    export_results(results, args.output_prefix)


if __name__ == "__main__":
    main()
