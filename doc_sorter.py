# --- Bibliothèques ---
import os, re, io, time, base64, shutil, filecmp, subprocess, anthropic  # processus, mots, mémoire, temps, encodage, fichiers, notifications, API Claude
from datetime import datetime                       # date et heure pour le journal
from pathlib import Path                            # manipulation des chemins
from anthropic import beta_tool                     # transforme une fonction Python en outil pour l'IA
from pypdf import PdfReader, PdfWriter              # lecture des PDF, extraction d'une page
from PIL import Image                               # réduction des images avant envoi
from docx import Document                           # lecture des Word
from openpyxl import load_workbook                  # lecture des Excel
from pptx import Presentation                       # lecture des PowerPoint
from watchdog.observers import Observer             # surveille un dossier
from watchdog.events import FileSystemEventHandler  # réagit aux nouveaux fichiers

# --- Configuration ---
UTILISATEUR = "Phil"                                # prénom utilisé dans le prompt : à adapter
HOME = Path.home()                                  # ex : C:\Users\<nom>
DL = HOME / "Downloads"                             # dossier surveillé
A_TRIER = DL / "A_trier"                            # destination en cas de doute
RACINES = [HOME / "Documents", HOME / "Pictures",   # seuls endroits autorisés (et leurs sous-dossiers)
           HOME / "Music", HOME / "Videos"]
JOURNAL = Path(__file__).parent / "journal.log"     # fichier de log, à côté du script
RATTRAPAGE_JOURS = 7                                # au démarrage, range les fichiers de moins de 7 jours
PID = Path(__file__).parent / "agent.pid"           # numéro du processus, pour n'arrêter que cet agent
PROFONDEUR = 4                                      # niveaux de sous-dossiers montrés à l'IA (Documents/Travail/Client/Projet/2026)
IMAGES = (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tif", ".tiff")  # montrées à l'IA, qui sait les lire

# --- Journal : écrit à l'écran ET dans le fichier ---
def journal(message):
    ligne = f"{datetime.now():%Y-%m-%d %H:%M:%S} | {message}"
    print(ligne)
    with open(JOURNAL, "a", encoding="utf-8") as f:                    # "a" = ajoute à la fin sans effacer
        f.write(ligne + "\n")

# --- Notification Windows : prévient quand un fichier n'est pas rangé ---
TOAST = (
    "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null;"
    "$x = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02);"
    "$t = $x.GetElementsByTagName('text');"
    "$t.Item(0).AppendChild($x.CreateTextNode($env:TOAST_TITRE)) | Out-Null;"
    "$t.Item(1).AppendChild($x.CreateTextNode($env:TOAST_TEXTE)) | Out-Null;"
    "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("
    "'{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe')"
    ".Show([Windows.UI.Notifications.ToastNotification]::new($x))")

def notifier(titre, texte):
    try:                                                                # textes passés par variables : pas d'injection via un nom de fichier
        subprocess.Popen(["powershell.exe", "-NoProfile", "-Command", TOAST],
                         env={**os.environ, "TOAST_TITRE": titre, "TOAST_TEXTE": texte},
                         creationflags=subprocess.CREATE_NO_WINDOW)     # sans fenêtre, sans attendre
    except Exception as e:
        journal(f"  notification impossible : {e}")

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
        for chemin, sous, _ in os.walk(racine):
            sous[:] = sorted(s for s in sous if not s.startswith("."))  # ignore les dossiers cachés
            if len(Path(chemin).relative_to(racine).parts) >= PROFONDEUR:
                sous.clear()                                            # ne descend pas plus bas
            dossiers.append(Path(chemin).relative_to(HOME).as_posix())  # ex : Documents/Banque/Relevés
    return "\n".join(dossiers[:500])                                    # limite pour rester rapide

# --- Outil 2 : retrouver les fichiers proches déjà rangés ---
@beta_tool
def chercher_similaires(nom: str) -> str:
    """Cherche dans les dossiers rangés des fichiers dont le nom ressemble à celui-ci (même sujet, autre version, même série).
    Un fichier proche déjà rangé indique souvent le bon dossier."""
    mots = {m for m in re.split(r"[\W_]+", Path(nom).stem.lower())     # mots du nom, accents compris
            if len(m) >= 4 and not m.isdigit()}                         # ignore "ia", "(1)", "2026"…
    if not mots:
        return "Nom trop court pour chercher des fichiers proches"
    trouves = []
    for racine in RACINES:
        for chemin, sous, fichiers in os.walk(racine):
            sous[:] = [s for s in sous if not s.startswith(".")]
            for f in fichiers:
                communs = mots & set(re.split(r"[\W_]+", Path(f).stem.lower()))
                if len(communs) >= min(2, len(mots)):                   # au moins 2 mots en commun
                    trouves.append((len(communs), (Path(chemin) / f).relative_to(HOME).as_posix()))
    if not trouves:
        return "Aucun fichier proche déjà rangé"
    trouves.sort(reverse=True)                                          # les plus ressemblants d'abord
    return "\n".join(f"{chemin} ({n} mots communs)" for n, chemin in trouves[:10])

# --- Outil 3 : lire le début d'un fichier ---
@beta_tool
def lire_debut(nom: str) -> str | list:
    """Lit le début du contenu d'un fichier PDF, Word, Excel, PowerPoint, texte ou image du dossier Téléchargements.
    Pour une image ou un PDF scanné, renvoie l'image de la 1re page."""
    f = dans_telechargements(nom)
    if not f:
        return "Fichier introuvable dans Téléchargements"
    ext = f.suffix.lower()
    try:
        if ext == ".pdf":
            pdf = PdfReader(f)
            texte = pdf.pages[0].extract_text() or ""                   # 1re page
            if len(texte.strip()) < 20:                                 # pas de texte : PDF scanné
                page = PdfWriter(); page.add_page(pdf.pages[0])         # 1re page seule, pour limiter le coût
                tampon = io.BytesIO(); page.write(tampon)
                journal(f"  {nom} : PDF scanné, 1re page envoyée à l'IA")
                return [{"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                         "data": base64.standard_b64encode(tampon.getvalue()).decode()}},
                        {"type": "text", "text": "PDF scanné : voici sa 1re page."}]
        elif ext in IMAGES:
            img = Image.open(f); img.seek(0)                            # 1re image si GIF animé ou TIFF multipage
            img = img.convert("RGB"); img.thumbnail((1568, 1568))       # réduite : moins de jetons, toujours lisible
            tampon = io.BytesIO(); img.save(tampon, "JPEG", quality=85)
            journal(f"  {nom} : image envoyée à l'IA")
            return [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                     "data": base64.standard_b64encode(tampon.getvalue()).decode()}},
                    {"type": "text", "text": "Voici l'image."}]
        elif ext == ".docx":
            texte = "\n".join(p.text for p in Document(f).paragraphs)   # paragraphes
        elif ext in (".xlsx", ".xlsm"):
            feuille = load_workbook(f, read_only=True, data_only=True).active
            texte = "\n".join(" | ".join(str(c) for c in row if c is not None)
                              for row in feuille.iter_rows(max_row=20, values_only=True))  # 20 lignes
        elif ext == ".pptx":
            texte = "\n".join(forme.text_frame.text                     # texte des diapositives, dans l'ordre
                              for diapo in Presentation(f).slides
                              for forme in diapo.shapes if forme.has_text_frame)
        else:
            with open(f, "rb") as h:
                debut = h.read(8000)                                    # 8 Ko seulement, même pour une vidéo de 2 Go
            if b"\x00" in debut:                                        # octets nuls = binaire (image, programme, archive…)
                journal(f"  {nom} : fichier binaire, non lu")
                return "Fichier binaire (image, programme, archive…) : contenu illisible, décide d'après le nom."
            texte = debut.decode("utf-8", errors="ignore")              # texte brut
        journal(f"  lecture du contenu de {nom}")
        return texte[:1500]                                             # 1500 caractères max
    except Exception as e:
        return f"Illisible : {e}"

# --- Outil 4 : déplacer un fichier ---
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
    if cible.is_file() and filecmp.cmp(src, cible, shallow=False):      # même nom ET même contenu, octet par octet
        src.unlink()                                                    # copie inutile : l'original est déjà rangé
        journal(f"DOUBLON | {nom} déjà présent dans {dossier} (identique) -> copie téléchargée supprimée")
        return "déjà présent à l'identique, copie supprimée"
    if cible.exists():                                                  # même nom, contenu différent : évite d'écraser
        cible = dossier / f"{src.stem}_{datetime.now():%Y-%m-%d_%H%M%S}{src.suffix}"
    shutil.move(src, cible)                                             # déplacement réel
    journal(f"RANGÉ   | {nom} -> {cible} | raison : {raison}")
    if dossier == A_TRIER:
        notifier(f"À trier : {nom}", raison)                            # l'utilisateur doit décider lui-même
    return "déplacé"

# --- L'agent : l'IA décide (prompt structuré ROCIF) ---
def ranger(nom):
    journal(f"NOUVEAU | {nom}")
    fin = anthropic.Anthropic(api_key=os.environ["DOC_SORTER_API_KEY"]).beta.messages.tool_runner(  # clé propre à cet agent
        model="claude-haiku-5-5", max_tokens=2000,                      # modèle rapide et peu cher
        tools=[lister_dossiers, chercher_similaires, lire_debut, deplacer],  # ses 4 outils
        messages=[{"role": "user", "content":
            f"RÔLE : Tu es l'assistant de rangement de fichiers de {UTILISATEUR}.\n"
            f"OBJECTIF : Ranger le fichier téléchargé « {nom} » au bon endroit.\n"
            f"CONTEXTE : {UTILISATEUR} classe ses fichiers par organisme ou projet (ex : Documents/Banque) "
            "avant de les classer par type. Les sous-dossiers précis sont toujours préférés aux dossiers génériques.\n"
            "INSTRUCTIONS :\n"
            "1. Liste les dossiers existants et cherche les fichiers au nom proche déjà rangés.\n"
            "2. Si un fichier proche traite du même sujet (autre version, même série), range-le dans le même dossier.\n"
            "3. Sinon, si le nom ne suffit pas, lis le début du fichier.\n"
            "4. Range-le dans le dossier existant le plus précis.\n"
            "5. S'il n'existe pas de dossier pour l'organisme ou le projet, utilise un dossier par type de document "
            "s'il y en a un (ex : Documents/Factures pour une facture). Mets-le dans A_trier seulement si aucun "
            "dossier ne convient, même par type. Ne devine jamais.\n"
            "FORMAT : Une seule action de déplacement, avec une raison d'une phrase."}]
    ).until_done()                                                      # tourne jusqu'à la fin
    if (DL / nom).exists():                                             # l'IA s'est arrêtée sans déplacer le fichier
        texte = " ".join(b.text for b in fin.content if b.type == "text").replace("\n", " ")
        journal(f"NON RANGÉ | {nom} | arrêt : {fin.stop_reason} | réponse : {texte[:300]}")
        deplacer(nom, "A_trier", "l'IA n'a pas choisi de dossier")      # filet de sécurité

# --- Attend que le fichier ait fini de s'écrire (taille stable) ---
def attendre_fin(p, max_secondes=600):
    taille = -1
    for _ in range(max_secondes):
        time.sleep(1)
        if not p.is_file():
            return False                                                # disparu entre-temps
        if (DL / (p.name + ".part")).exists():                          # Firefox : fichier vide + .part en cours
            return False                                                # on sera rappelé au renommage du .part
        nouvelle = p.stat().st_size
        if nouvelle == taille:
            return True                                                 # 2 mesures identiques : téléchargement fini
        taille = nouvelle
    journal(f"ERREUR  | {p.name} | taille encore instable après {max_secondes} s")
    notifier(f"Erreur : {p.name}", "Téléchargement jamais terminé, fichier non rangé")
    return False

# --- Le déclencheur ---
class Handler(FileSystemEventHandler):
    def on_created(self, e): self.check(e.src_path)                     # nouveau fichier
    def on_moved(self, e): self.check(e.dest_path)                      # fichier renommé (fin de téléchargement)
    def check(self, p):
        p = Path(p)
        try:
            if p.parent == DL and p.suffix.lower() not in (".crdownload", ".tmp", ".part"):  # ignore les téléchargements en cours
                if attendre_fin(p):
                    ranger(p.name)                                      # lance l'agent
        except Exception as e:
            journal(f"ERREUR  | {p.name} | {e}")                        # erreurs aussi dans le journal
            notifier(f"Erreur : {p.name}", str(e)[:200])                # ex : clé API expirée

# --- Démarrage ---
obs = Observer(); obs.schedule(Handler(), str(DL)); obs.start()         # surveille Downloads
PID.write_text(str(os.getpid()))                                        # permet l'arrêt ciblé
journal(f"DÉMARRAGE | surveillance de {DL} | PID {os.getpid()}")

# --- Rattrapage : fichiers arrivés pendant que l'agent était arrêté ---
def deja_vus():
    """Dernière date à laquelle l'agent a traité chaque nom de fichier, d'après le journal."""
    vus = {}
    if JOURNAL.exists():
        for ligne in JOURNAL.read_text(encoding="utf-8", errors="ignore").splitlines():
            date, _, reste = ligne.partition(" | ")
            if reste.startswith("NOUVEAU | "):
                try:
                    vus[reste[10:]] = datetime.strptime(date, "%Y-%m-%d %H:%M:%S").timestamp()
                except ValueError:
                    pass                                                # ligne abîmée : ignorée
    return vus

limite = time.time() - RATTRAPAGE_JOURS * 86400
vus = deja_vus()
recents = [p for p in sorted(DL.iterdir()) if p.is_file() and p.stat().st_mtime > limite]
en_attente = [p for p in recents if p.stat().st_mtime > vus.get(p.name, 0)]  # déjà traité et pas modifié depuis : laissé exprès
journal(f"RATTRAPAGE | {len(en_attente)} fichier(s) de moins de {RATTRAPAGE_JOURS} jours"
        f" ({len(recents) - len(en_attente)} déjà traité(s), laissé(s) en place)")
for p in en_attente:
    Handler().check(p)                                                  # mêmes filtres qu'un nouveau téléchargement

while True: time.sleep(1)                                               # Ctrl + C pour arrêter