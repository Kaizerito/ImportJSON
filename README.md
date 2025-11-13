# ImportJSON

Import JSON from any URL directly into your Google Sheets. `ImportJSON.gs` adds an `=ImportJSON()` function to your spreadsheet, allowing quick and easy JSON importing. To use go to `Tools` > `Script Editor` and add the `ImportJSON.gs` file. Now in your spreadsheet you can access the `ImportJSON()` function. Use it like this:

    =ImportJSON("https://mysafeinfo.com/api/data?list=bestnovels&format=json&rows=20&alias=cnt=count,avg=average_rank,tt=title,au=author,yr=year", "/title")

Here are all the functions available:

| Function                |  Description                                                                      |
|-------------------------|-----------------------------------------------------------------------------------|
| **ImportJSON**          | For use by end users to import a JSON feed from a URL                             |
| **ImportJSONFromSheet** | For use by end users to import JSON from one of the Sheets                        |
| **ImportJSONViaPost**   | For use by end users to import a JSON feed from a URL using POST parameters       |
| **ImportJSONBasicAuth** | For use by end users to import a JSON feed from a URL with HTTP Basic Auth        |
| **ImportJSONAdvanced**  | For use by script developers to easily extend the functionality of this library   |

Review `ImportJSON.gs` for more info on how to use these in detail.

## Version
- v1.6.0 (June 2, 2019) Fixed null values (thanks @gdesmedt1)
- v1.5.0 (January 11, 2019) Adds ability to include all headers in a fixed order even when no data is present for a given header in some or all rows.
- v1.4.0 (July 23, 2017) - Project transferred to Brad Jasper. Fixed off-by-one array bug. Fixed previous value bug. Added custom annotations. Added ImportJSONFromSheet and ImportJSONBasicAuth.
- v1.3.0 - Adds ability to import the text from a set of rows containing the text to parse. All cells are concatenated
- v1.2.1 - Fixed a bug with how nested arrays are handled. The rowIndex counter wasn't incrementing properly when parsing.
- v1.2.0 - Added ImportJSONViaPost and support for fetchOptions to ImportJSONAdvanced
- v1.1.1 - Added a version number using Google Scripts Versioning so other developers can use the library
- v1.1.0 - Added support for the noHeaders option
- v1.0.0 - Initial release

## How can you help?
- Found a bug? Report it! https://github.com/bradjasper/ImportJSON/issues
- Want to contribute? Submit an <a href="https://github.com/bradjasper/ImportJSON/issues?q=is%3Aissue+is%3Aopen+label%3Aenhancement">enhancement</a>

## Website archive
This code base used to be hosted at http://blog.fastfedora.com/projects/import-json and contained a lot of useful information. It has been archived at https://rawgit.com/bradjasper/ImportJSON/master/archive/blog.fastfedora.com/projects/import-json.html

## Alternatives
Some of this if possible internally with Google App Scripts External APIs, like UrlFetch: https://developers.google.com/apps-script/guides/services/external

These require a Google account and an explicit permission, but in some cases may be a good fit.

## Script de calcul des distances Arc-et-Senans

Le dépôt contient également `scripts/compute_distances.py`, un utilitaire Python qui géocode
une liste de communes autour d'Arc-et-Senans puis calcule les distances/temps de trajet via
l'API OpenRouteService.

### Pré-requis

1. **Python 3.9+** installé sur votre machine.
2. Créez (optionnel mais recommandé) un environnement virtuel :

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # sous Windows : .venv\Scripts\activate
   ```

3. Installez les dépendances nécessaires :

   ```bash
   pip install openrouteservice pandas
   ```

4. Récupérez une clé API personnelle sur <https://openrouteservice.org/sign-up/>.

### Fournir la clé API

Choisissez l'une des méthodes suivantes :

* Via l'argument CLI : `python scripts/compute_distances.py --api-key VOTRE_CLE`
* Via une variable d'environnement :

  ```bash
  export ORS_API_KEY="VOTRE_CLE"  # Windows PowerShell : $env:ORS_API_KEY="VOTRE_CLE"
  ```

* Via un fichier texte ou `.env` contenant une ligne `ORS_API_KEY=VOTRE_CLE`, puis
  exécutez `python scripts/compute_distances.py --api-key-file chemin/vers/.env`

### Lancement du script

Ensuite, lancez simplement :

```bash
python scripts/compute_distances.py
```

Le script affichera l'avancement pour chaque commune, puis générera deux fichiers dans le
répertoire courant : `distances_arc_et_senans.csv` et `distances_arc_et_senans.xlsx`.

