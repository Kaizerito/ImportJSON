"""Find communes reachable within a travel-time budget from a given address.

This command-line tool relies on the OpenRouteService (ORS) APIs to:

1. Géocoder une adresse de départ (ex. « Arc-et-Senans, France »).
2. Géocoder une liste pré-définie de communes dans le Doubs / Jura.
3. Calculer les itinéraires routiers voiture entre l'adresse et chaque commune.
4. Filtrer les communes dont le temps de trajet est inférieur ou égal au budget
   fourni (en minutes).

Les résultats peuvent être affichés dans le terminal et, si désiré, exportés
au format CSV / Excel.

Quick start:
1. Installez les dépendances : ``pip install openrouteservice pandas``.
2. Fournissez la clé ORS via ``--api-key``, ``--api-key-file``, la variable
   ``ORS_API_KEY`` ou un fichier ``.env`` (voir ``--help``).
3. Lancez :
   ``python scripts/compute_distances.py --address "Arc-et-Senans" --max-duration 20``
   (remplacez l'adresse / le temps par vos valeurs).

Le script temporise les appels API pour éviter de dépasser les limites
d'utilisation d'OpenRouteService. La durée totale dépendra du temps de réponse
des services externes.
"""

from __future__ import annotations

import argparse
import importlib.util
import os
import sys
import time
from pathlib import Path
from typing import Iterable, List, Tuple, TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover - uniquement pour l'analyse statique
    import openrouteservice

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
        "--address",
        dest="address",
        required=True,
        help="Adresse de départ (ex: 'Arc-et-Senans, France').",
    )
    parser.add_argument(
        "--max-duration",
        dest="max_duration",
        type=float,
        required=True,
        help="Temps de trajet maximum en minutes.",
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
        default=None,
        help="Nom de base pour exporter les résultats (CSV + XLSX). Aucun export si omis.",
    )
    parser.add_argument(
        "--include-out-of-range",
        dest="include_out_of_range",
        action="store_true",
        help="Inclut les communes hors budget dans l'export/affichage.",
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


def ensure_dependency(module_name: str, install_hint: str) -> None:
    """Exit with a friendly hint if *module_name* is not importable."""

    if importlib.util.find_spec(module_name) is None:
        raise SystemExit(
            f"Le module Python '{module_name}' est requis. Installez-le avec : {install_hint}"
        )


def init_client(api_key: str) -> "openrouteservice.Client":
    ensure_dependency("openrouteservice", "pip install openrouteservice")
    import openrouteservice

    return openrouteservice.Client(key=api_key)


def geocode(
    client: "openrouteservice.Client", query: str
) -> Tuple[float | None, float | None]:
    response = client.pelias_search(text=query)
    if response.get("features"):
        coordinates = response["features"][0]["geometry"]["coordinates"]
        return coordinates[1], coordinates[0]
    return None, None


def geocode_commune(
    client: "openrouteservice.Client", commune: Tuple[str, str]
) -> Tuple[str, str, float | None, float | None]:
    postal_code, name = commune
    lat, lon = geocode(client, f"{name}, {postal_code}, France")
    return postal_code, name, lat, lon


def compute_route(
    client: "openrouteservice.Client",
    start_lat: float,
    start_lon: float,
    lat: float,
    lon: float,
) -> Tuple[float, float]:
    route = client.directions(
        coordinates=[[start_lon, start_lat], [lon, lat]],
        profile="driving-car",
        format="json",
    )
    summary = route["routes"][0]["summary"]
    distance_km = summary["distance"] / 1000
    duration_min = summary["duration"] / 60
    return distance_km, duration_min


def collect_distances(
    client: "openrouteservice.Client",
    start_lat: float,
    start_lon: float,
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
        distance_km, duration_min = compute_route(client, start_lat, start_lon, lat, lon)
        results.append(
            {
                "Code postal": postal_code,
                "Commune": name,
                "Distance (km)": round(distance_km, 2),
                "Temps voiture (min)": round(duration_min, 1),
            }
        )
        time.sleep(max(0.0, sleep_seconds))
    return results


def export_results(data: List[dict], output_prefix: str) -> None:
    ensure_dependency("pandas", "pip install pandas")
    import pandas as pd

    df = pd.DataFrame(data)
    csv_path = f"{output_prefix}.csv"
    xlsx_path = f"{output_prefix}.xlsx"
    df.to_csv(csv_path, index=False)
    df.to_excel(xlsx_path, index=False)
    print(f"Fichiers générés : {xlsx_path} + {csv_path}")


def print_results_table(results: List[dict]) -> None:
    if not results:
        print("Aucune commune ne correspond au budget de temps indiqué.")
        return

    headers = ["Code postal", "Commune", "Temps voiture (min)", "Distance (km)"]
    col_widths = {header: len(header) for header in headers}
    for row in results:
        for header in headers:
            col_widths[header] = max(col_widths[header], len(str(row.get(header, ""))))

    def format_row(row: dict | None) -> str:
        if row is None:
            return " | ".join(header.ljust(col_widths[header]) for header in headers)
        return " | ".join(str(row.get(header, "")).ljust(col_widths[header]) for header in headers)

    separator = "-+-".join("-" * col_widths[header] for header in headers)
    print(format_row(None))
    print(separator)
    for row in results:
        print(format_row(row))


def main() -> None:
    args = parse_args()
    api_key = resolve_api_key(args.api_key, args.api_key_file)
    client = init_client(api_key)
    start_lat, start_lon = geocode(client, args.address)
    if start_lat is None or start_lon is None:
        raise SystemExit(
            "Adresse de départ introuvable via ORS. Vérifiez l'orthographe ou précisez davantage."
        )

    all_results = collect_distances(client, start_lat, start_lon, COMMUNES, args.sleep)
    if not all_results:
        print("Aucun résultat calculé.")
        return

    all_results.sort(key=lambda item: item["Temps voiture (min)"])
    in_range = [
        row for row in all_results if row["Temps voiture (min)"] <= round(args.max_duration, 10)
    ]

    print()
    print(
        f"Communes atteignables en ≤ {args.max_duration} min depuis '{args.address}':"
    )
    print_results_table(in_range)

    if args.include_out_of_range:
        print()
        print("Communes hors budget de temps :")
        out_of_range = [row for row in all_results if row not in in_range]
        print_results_table(out_of_range)

    if args.output_prefix:
        try:
            export_results(
                all_results if args.include_out_of_range else in_range,
                args.output_prefix,
            )
        except SystemExit:
            raise
        except Exception as exc:  # pragma: no cover - affichage utilisateur
            print(f"Erreur lors de l'export : {exc}", file=sys.stderr)
            raise


if __name__ == "__main__":
    main()
