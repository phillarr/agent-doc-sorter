# --- Bibliothèques ---
import time, shutil, anthropic                      # temps, déplacement de fichiers, API Claude
from datetime import datetime                       # date et heure pour le journal
from pathlib import Path                            # manipulation des chemins
from anthropic import beta_tool                     # transforme une fonction Python en outil pour l'IA
from pypdf import PdfReader                         # lecture des PDF
from docx import Document                           # lecture des Word
from openpyxl import load_workbook                  # lecture des Excel
from watchdog.observers import Observer             # surveille un dossier
from watchdog.events import FileSystemEventHandler  # réagit aux nouveaux fichiers

# --- Configuration ---
HOME = Path.home()                                  # C:\Users\<nom>
DL = HOME / "Downloads"                             # dossier surveillé
A_TRIER = DL / "A_trier"                            # destination en cas de doute
RACINES = [HOME / "Documents", HOME / "Pictures",   # seuls endroits autorisés (et leurs sous-dossiers)
           HOME / "Music", HOME / "Videos"]
JOURNAL = Path(__file__).parent / "journal.log"     # fichier de log, à côté du script
RATTRAPAGE_JOURS = 7                                # au démarrage, range les fichiers de moins de 7 jours

# --- Journal : écrit à l'écran ET dans le fichier ---
def journal(message):
    ligne = f"{datetime.now():%Y-%m-%d %H:%M:%S} | {message}"
    print(ligne)
    with open(JOURNAL, "a", encoding="utf-8") as f:                    # "a" = ajoute à la fin sans effacer
        f.write(ligne + "\n")

# --- Sécurité : l'IA ne peut toucher qu'aux fichiers posés directement dans Téléchargements ---
def dans_telechargements(nom):
    f = (DL / nom).resolve()
    if f.parent != DL.resolve() or not f.is_file():                     # bloque "../Documents/..." et les dossiers
        journal(f"  refusé : {nom} (hors Téléchargements)")
        return None
    return f

# --- Outil 1 : voir les dossiers existants ---
@beta_tool
def lister_dossiers() -> str:
    """Liste les dossiers et sous-dossiers existants où un fichier peut être rangé."""
    dossiers = []
    for racine in RACINES:
        dossiers.append(racine.name)                                    # ex : Documents
        for d in racine.glob("*/"):                                     # 1er niveau : Documents/Travail
            if not d.name.startswith("."):
                dossiers.append(f"{racine.name}/{d.name}")
                for s in d.glob("*/"):                                  # 2e niveau : Documents/Travail/Client
                    if not s.name.startswith("."):
                        dossiers.append(f"{racine.name}/{d.name}/{s.name}")
    return "\n".join(dossiers[:300])                                    # limite pour rester rapide

# --- Outil 2 : lire le début d'un fichier ---
@beta_tool
def lire_debut(nom: str) -> str:
    """Lit le début du contenu d'un fichier PDF, Word, Excel ou texte du dossier Téléchargements."""
    f = dans_telechargements(nom)
    if not f:
        return "Fichier introuvable dans Téléchargements"
    ext = f.suffix.lower()
    try:
        if ext == ".pdf":
            texte = PdfReader(f).pages[0].extract_text() or ""          # 1re page
        elif ext == ".docx":
            texte = "\n".join(p.text for p in Document(f).paragraphs)   # paragraphes
        elif ext in (".xlsx", ".xlsm"):
            feuille = load_workbook(f, read_only=True, data_only=True).active
            texte = "\n".join(" | ".join(str(c) for c in row if c is not None)
                              for row in feuille.iter_rows(max_row=20, values_only=True))  # 20 lignes
        else:
            texte = f.read_text(errors="ignore")                        # texte brut
        journal(f"  lecture du contenu de {nom}")
        return texte[:1500]                                             # 1500 caractères max
    except Exception as e:
        return f"Illisible : {e}"

# --- Outil 3 : déplacer un fichier ---
@beta_tool
def deplacer(nom: str, destination: str, raison: str) -> str:
    """Déplace un fichier vers un dossier existant (chemin tel que donné par lister_dossiers) ou vers 'A_trier'.
    raison : explique en une phrase courte pourquoi tu choisis ce dossier."""
    src = dans_telechargements(nom)
    if not src:                                                         # déjà déplacé ou hors Téléchargements
        return "Fichier introuvable dans Téléchargements"
    nom = src.name
    if destination == "A_trier":
        dossier = A_TRIER
        dossier.mkdir(exist_ok=True)                                    # seul dossier qu'il peut créer
    else:
        dossier = (HOME / destination).resolve()
        autorise = any(dossier == r or r in dossier.parents for r in RACINES)
        if not autorise or not dossier.is_dir():                        # sécurité : existant et autorisé uniquement
            journal(f"  refusé : {destination}")
            return "Destination non autorisée ou inexistante, choisis-en une de la liste"
    cible = dossier / nom
    if cible.exists():                                                  # évite d'écraser
        cible = dossier / f"{src.stem}_{int(time.time())}{src.suffix}"
    shutil.move(src, cible)                                             # déplacement réel
    journal(f"RANGÉ   | {nom} -> {cible} | raison : {raison}")
    return "déplacé"

# --- L'agent : l'IA décide (prompt structuré ROCIF) ---
def ranger(nom):
    journal(f"NOUVEAU | {nom}")
    fin = anthropic.Anthropic().beta.messages.tool_runner(
        model="claude-haiku-5-5", max_tokens=800,                       # modèle rapide et peu cher
        tools=[lister_dossiers, lire_debut, deplacer],                  # ses 3 outils
        messages=[{"role": "user", "content":
            "RÔLE : Tu es l'assistant de rangement de fichiers de Phil.\n"
            f"OBJECTIF : Ranger le fichier téléchargé « {nom} » au bon endroit.\n"
            "CONTEXTE : Phil classe ses fichiers par organisme ou projet (ex : Documents/Travail) "
            "avant de les classer par type. Les sous-dossiers précis sont toujours préférés aux dossiers génériques.\n"
            "INSTRUCTIONS :\n"
            "1. Liste les dossiers existants.\n"
            "2. Si le nom ne suffit pas, lis le début du fichier.\n"
            "3. Range-le dans le dossier existant le plus précis.\n"
            "4. S'il n'existe pas de dossier pour l'organisme ou le projet, utilise un dossier par type de document "
            "s'il y en a un (ex : Documents/Factures pour une facture). Mets-le dans A_trier seulement si aucun "
            "dossier ne convient, même par type. Ne devine jamais.\n"
            "FORMAT : Une seule action de déplacement, avec une raison d'une phrase."}]
    ).until_done()                                                      # tourne jusqu'à la fin
    if (DL / nom).exists():                                             # l'IA s'est arrêtée sans déplacer le fichier
        texte = " ".join(b.text for b in fin.content if b.type == "text").replace("\n", " ")
        journal(f"NON RANGÉ | {nom} | arrêt : {fin.stop_reason} | réponse : {texte[:300]}")
        deplacer(nom, "A_trier", "l'IA n'a pas choisi de dossier")      # filet de sécurité

# --- Le déclencheur ---
class Handler(FileSystemEventHandler):
    def on_created(self, e): self.check(e.src_path)                     # nouveau fichier
    def on_moved(self, e): self.check(e.dest_path)                      # fichier renommé (fin de téléchargement)
    def check(self, p):
        p = Path(p)
        try:
            if p.parent == DL and p.suffix.lower() not in (".crdownload", ".tmp", ".part"):  # ignore les téléchargements en cours
                time.sleep(2)                                           # laisse le fichier finir de s'écrire
                if p.is_file():
                    ranger(p.name)                                      # lance l'agent
        except Exception as e:
            journal(f"ERREUR  | {p.name} | {e}")                        # erreurs aussi dans le journal

# --- Démarrage ---
obs = Observer(); obs.schedule(Handler(), str(DL)); obs.start()         # surveille Downloads
journal(f"DÉMARRAGE | surveillance de {DL}")

# --- Rattrapage : fichiers arrivés pendant que l'agent était arrêté ---
limite = time.time() - RATTRAPAGE_JOURS * 86400
en_attente = [p for p in sorted(DL.iterdir()) if p.is_file() and p.stat().st_mtime > limite]
journal(f"RATTRAPAGE | {len(en_attente)} fichier(s) de moins de {RATTRAPAGE_JOURS} jours")
for p in en_attente:
    Handler().check(p)                                                  # mêmes filtres qu'un nouveau téléchargement

while True: time.sleep(1)                                               # Ctrl + C pour arrêter