# agent-doc-sorter

Agent qui range automatiquement les fichiers téléchargés.

## Fonctionnement

- `doc_sorter.py` surveille `~/Downloads` (watchdog). À chaque nouveau fichier, un agent Claude (`claude-haiku-5-5`, via `tool_runner`) choisit où le ranger grâce à 4 outils :
  - `lister_dossiers` : dossiers existants sous Documents, Pictures, Music, Videos (4 niveaux) ;
  - `chercher_similaires` : fichiers déjà rangés dont le nom partage au moins 2 mots avec le nouveau (pour ranger une nouvelle version à côté de l'ancienne) ;
  - `lire_debut` : début du contenu (PDF, Word, Excel, PowerPoint, texte ; les autres fichiers binaires ne sont pas lus). Les images (réduites à 1568 px) et la 1re page des PDF scannés sont envoyées telles quelles à Claude, qui les lit ;
  - `deplacer` : déplace vers un dossier existant et autorisé, ou vers `Downloads/A_trier` en cas de doute.
- Chaque action est écrite dans `journal.log` (à côté du script, ignoré par git).
- Au démarrage, l'agent range les fichiers de moins de 7 jours arrivés pendant qu'il était arrêté. Un fichier déjà traité (son nom figure dans le journal) et non modifié depuis est considéré comme laissé exprès dans Téléchargements, et n'est pas déplacé.
- Une notification Windows signale chaque fichier envoyé dans `A_trier` et chaque erreur.
- Dépendances : `anthropic`, `watchdog`, `pypdf`, `python-docx`, `openpyxl`, `python-pptx`, `pillow`. La clé API est lue dans la variable d'environnement Utilisateur `DOC_SORTER_API_KEY` (clé « agent-doc-sorter » dans la console Anthropic, distincte de celle des autres agents).

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
