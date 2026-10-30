# agent-doc-sorter

Agent IA qui range automatiquement les fichiers téléchargés sous Windows.

Un fichier arrive dans Téléchargements : l'agent regarde les dossiers existants, lit le début du fichier si besoin, puis le déplace au bon endroit dans Documents, Images, Musique ou Vidéos. En cas de doute, il le met dans `Téléchargements\A_trier` et prévient par une notification.

## Fonctionnement

- `doc_sorter.py` surveille `~/Downloads` (watchdog). À chaque nouveau fichier, un agent Claude (`claude-haiku-5-5`, via `tool_runner`) choisit où le ranger grâce à 4 outils :
  - `lister_dossiers` : dossiers existants sous Documents, Pictures, Music, Videos (4 niveaux) ;
  - `chercher_similaires` : fichiers déjà rangés dont le nom partage au moins 2 mots avec le nouveau (pour ranger une nouvelle version à côté de l'ancienne) ;
  - `lire_debut` : début du contenu (PDF, Word, Excel, PowerPoint, texte ; les autres fichiers binaires ne sont pas lus). Les images (réduites à 1568 px) et la 1re page des PDF scannés sont envoyées telles quelles à Claude, qui les lit ;
  - `deplacer` : déplace vers un dossier existant et autorisé, ou vers `Downloads/A_trier` en cas de doute.
- Garde-fous : l'agent ne touche qu'aux fichiers posés directement dans Téléchargements, ne range que dans des dossiers existants sous Documents, Pictures, Music et Videos, et n'écrase jamais un fichier (suffixe date-heure si le nom existe déjà ; une copie identique est supprimée).
- Chaque action est écrite dans `journal.log` (à côté du script, ignoré par git).
- Au démarrage, l'agent range les fichiers de moins de 7 jours arrivés pendant qu'il était arrêté. Un fichier déjà traité (son nom figure dans le journal) et non modifié depuis est considéré comme laissé exprès dans Téléchargements, et n'est pas déplacé.
- Une notification Windows signale chaque fichier envoyé dans `A_trier` et chaque erreur.

**Confidentialité** : les noms de vos dossiers, de quelques fichiers déjà rangés et le début du contenu des fichiers téléchargés sont envoyés à l'API Anthropic pour décider du rangement.

## Installation

Prérequis : Windows 10/11, Python 3.10 ou plus récent, un compte sur [console.anthropic.com](https://console.anthropic.com) (l'API est facturée à l'usage, séparément d'un abonnement Claude).

1. Cloner le dépôt et installer les dépendances :

   ```powershell
   git clone https://github.com/phillarr/agent-doc-sorter.git $HOME\agent-doc-sorter
   cd $HOME\agent-doc-sorter
   pip install -r requirements.txt
   ```

2. Créer une clé API sur la console Anthropic (Settings → API Keys), puis l'enregistrer comme variable d'environnement Utilisateur, dans PowerShell :

   ```powershell
   [Environment]::SetEnvironmentVariable("DOC_SORTER_API_KEY", "votre-clé", "User")
   ```

3. Adapter `UTILISATEUR` (votre prénom) en haut de `doc_sorter.py`.

4. Tester au premier plan (Ctrl + C pour arrêter), dans une nouvelle fenêtre PowerShell pour que la variable soit prise en compte :

   ```powershell
   python $HOME\agent-doc-sorter\doc_sorter.py
   ```

5. Lancer l'agent au démarrage de Windows : créer le fichier `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\agent-doc-sorter.bat` contenant :

   ```bat
   start "" pyw "%USERPROFILE%\agent-doc-sorter\doc_sorter.py"
   ```

   `pyw` lance Python sans fenêtre. S'il n'est pas installé, remplacer `pyw` par le chemin complet de `pythonw.exe`.

## Après chaque modification du code

Le processus en cours garde l'ancienne version : il faut l'arrêter puis le relancer.

Au démarrage, l'agent écrit son numéro de processus dans `agent.pid` (ignoré par git), ce qui permet de n'arrêter que lui :

```powershell
Get-Process -Id (Get-Content $HOME\agent-doc-sorter\agent.pid) -ErrorAction SilentlyContinue | Where-Object Name -eq pythonw | Stop-Process -Force
explorer.exe "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\agent-doc-sorter.bat"
```

Le filtre `pythonw` évite d'arrêter un autre programme si le numéro a été réattribué. Passer par `explorer.exe` garantit que l'agent hérite de la variable utilisateur `DOC_SORTER_API_KEY`.

Vérifier ensuite dans `journal.log` qu'une ligne `DÉMARRAGE` récente apparaît, puis qu'un fichier test est bien `RANGÉ` (une erreur d'authentification n'apparaît qu'au premier fichier).

## Licence

MIT, voir [LICENSE](LICENSE).
