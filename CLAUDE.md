# agent-doc-sorter

Agent qui range automatiquement les fichiers téléchargés.

## Fonctionnement

- `doc_sorter.py` surveille `~/Downloads` (watchdog). À chaque nouveau fichier, un agent Claude (`claude-haiku-5-5`, via `tool_runner`) choisit où le ranger grâce à 3 outils :
  - `lister_dossiers` : dossiers existants sous Documents, Pictures, Music, Videos (2 niveaux) ;
  - `lire_debut` : début du contenu (PDF, Word, Excel, texte) ;
  - `deplacer` : déplace vers un dossier existant et autorisé, ou vers `Downloads/A_trier` en cas de doute.
- Chaque action est écrite dans `journal.log` (à côté du script, ignoré par git).
- Dépendances : `anthropic`, `watchdog`, `pypdf`, `python-docx`, `openpyxl`. La clé `ANTHROPIC_API_KEY` doit être définie dans l'environnement.

## Exécution en arrière-plan

L'agent tourne en arrière-plan via `pythonw` (sans fenêtre), lancé au démarrage de Windows par :

`C:\Users\<nom>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\agent-doc-sorter.bat`

## Après chaque modification du code

Le processus en cours garde l'ancienne version : il faut l'arrêter puis le relancer.

```powershell
taskkill /im pythonw.exe /f
& "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\agent-doc-sorter.bat"
```

Attention : `taskkill /im pythonw.exe /f` arrête **tous** les processus `pythonw`, pas seulement cet agent.

Vérifier ensuite dans `journal.log` qu'une ligne `DÉMARRAGE` récente apparaît.
