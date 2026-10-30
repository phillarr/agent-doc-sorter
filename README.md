# agent-doc-sorter

Agent qui range automatiquement les fichiers téléchargés.

## Fonctionnement

- `doc_sorter.py` surveille `~/Downloads` (watchdog). À chaque nouveau fichier, un agent Claude (`claude-haiku-5-5`, via `tool_runner`) choisit où le ranger grâce à 3 outils :
  - `lister_dossiers` : dossiers existants sous Documents, Pictures, Music, Videos (2 niveaux) ;
  - `lire_debut` : début du contenu (PDF, Word, Excel, PowerPoint, texte ; les fichiers binaires ne sont pas lus) ;
  - `deplacer` : déplace vers un dossier existant et autorisé, ou vers `Downloads/A_trier` en cas de doute.
- Chaque action est écrite dans `journal.log` (à côté du script, ignoré par git).
- Dépendances : `anthropic`, `watchdog`, `pypdf`, `python-docx`, `openpyxl`, `python-pptx`. La clé API est lue dans la variable d'environnement Utilisateur `DOC_SORTER_API_KEY` (clé « agent-doc-sorter » dans la console Anthropic, distincte de celle des autres agents).

## Exécution en arrière-plan

L'agent tourne en arrière-plan via `pythonw` (sans fenêtre), lancé au démarrage de Windows par :

`C:\Users\<nom>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\agent-doc-sorter.bat`

## Après chaque modification du code

Le processus en cours garde l'ancienne version : il faut l'arrêter puis le relancer.

Au démarrage, l'agent écrit son numéro de processus dans `agent.pid` (ignoré par git), ce qui permet de n'arrêter que lui :

```powershell
Get-Process -Id (Get-Content C:\Users\<nom>\agent-doc-sorter\agent.pid) -ErrorAction SilentlyContinue | Where-Object Name -eq pythonw | Stop-Process -Force
explorer.exe "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\agent-doc-sorter.bat"
```

Le filtre `pythonw` évite d'arrêter un autre programme si le numéro a été réattribué. Passer par `explorer.exe` garantit que l'agent hérite de la variable utilisateur `DOC_SORTER_API_KEY`.

Vérifier ensuite dans `journal.log` qu'une ligne `DÉMARRAGE` récente apparaît, puis qu'un fichier test est bien `RANGÉ` (une erreur d'authentification n'apparaît qu'au premier fichier).
