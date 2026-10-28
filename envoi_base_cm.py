#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
envoi_base_cm.py
================

Envoi automatisé de la base CM "hors ligne" par courriel via GroupWise.

Ce script prend le relais de la macro Access pour l'étape 5 de la chaîne de
traitement quotidienne (lancée vers 23h30 par Access) :

    1. Lancée à 23h30 tous les jours (Access)
    2. Génère une copie d'une base Access (Access)
    3. Lance des traitements Access (copie des tables liées) (Access)
    4. Compresse la base avec 7zip -> CM_AAAAMMJJ_HHMMSS.7z (Access)
    5. Envoi de l'archive .7z à une liste de diffusion
       >>> DÉSORMAIS RÉALISÉ ICI EN PYTHON VIA GroupwiseCommander <<<

EXÉCUTION 32 BITS (x86) OBLIGATOIRE :
    L'objet OLE `GroupwiseCommander` nécessite Python 32 bits (x86).

Bibliothèques requises :
    pywin32, python-dotenv, pyodbc (+ pilote Microsoft Access)

Configuration (fichier .env, variables) :
    RepFicAenvoyer     = dossier contenant l'archive .7z et le fichier Word (finissant par un backslash)
    DateFicAenvoyer    = "today" (date du jour) OU date exacte "AAAAMMJJ"
    NomFichierWord     = nom du fichier Word (mode opératoire) présent dans RepFicAenvoyer
    BaseUsers          = chemin complet de la base Access (table TbUsers)
    FicLog             = chemin complet du fichier de log commun (mode append)
    FicLogTrtComplet   = (optionnel) chemin d'un second fichier de log (log global
                         de l'ensemble du traitement). S'il est défini, les lignes
                         DEBUT / infos / FIN sont également écrites dans ce fichier.

Sortie :
    Code de retour 0 si succès, 1 en cas d'erreur.
"""

import os
import sys
import re
import glob
import time
import locale
import datetime
import subprocess

# ---------------------------------------------------------------------------
# Dépendances (messages clairs si manquantes)
# ---------------------------------------------------------------------------
try:
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "ERREUR : la bibliothèque 'python-dotenv' est requise.\n"
        "Installez-la avec : pip install python-dotenv\n"
    )
    sys.exit(1)

try:
    import pyodbc
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "ERREUR : la bibliothèque 'pyodbc' est requise.\n"
        "Installez-la avec : pip install pyodbc\n"
    )
    sys.exit(1)

try:
    import win32com.client
    import win32clipboard
    import win32con
    import win32gui
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "ERREUR : la bibliothèque 'pywin32' est requise (modules win32com, "
        "win32clipboard, win32con et win32gui).\n"
        "Installez-la avec : pip install pywin32\n"
    )
    sys.exit(1)

# ---------------------------------------------------------------------------
# Constantes du cahier des charges
# ---------------------------------------------------------------------------
# Séparateur de log : 51 astérisques, strictement identique au séparateur employé
# par le traitement Access afin de rester cohérent dans le fichier de log commun
# (FicLogTrtComplet / TT_Import.log).
SEP_LOG = "*" * 51

LIEN_COPIE_EXCEL = (
    "G:\\DM-SURVEIL\\MEDICAMENTS\\BASES\\BASE_CM\\Import_TT\\ImportExcel_TT\\PROD\\aFaire\\"
)

SUJET_MAIL = (
    "Base Cas marquants hors ligne pour saisie par DP et DS/Sécurisation "
    "en période de télétravail"
)


# ---------------------------------------------------------------------------
# Journalisation (mode append) - format identique au VBA / Access
# ---------------------------------------------------------------------------
def _writelog(fic_log, texte):
    """Écrit '<texte>' en mode append dans le fichier de log.

    L'encodage suit celui de l'OS (ANSI/Windows-1252 sous Windows), identique
    à celui employé par le traitement Access pour rester lisible par le même
    fichier de log commun.

    On ajoute un simple '\\n' : en mode texte sous Windows, Python le traduit en
    '\\r\\n' (CRLF) — la fin de ligne employée par Access. Utiliser os.linesep
    (déjà '\\r\\n') provoquerait un double retour '\\r\\r\\n' et donc une ligne
    vide parasite une ligne sur deux dans le fichier de log partagé.
    """
    enc = locale.getpreferredencoding(False) or "utf-8"
    with open(fic_log, "a", encoding=enc) as fh:
        fh.write(texte + "\n")


def _write_multi(fichiers, texte):
    """Écrit '<texte>' dans chacun des fichiers de la liste 'fichiers'."""
    for fic in fichiers:
        _writelog(fic, texte)


def log_debut(fichiers):
    """Ligne DEBUT conforme au cahier des charges, écrite dans chaque fichier."""
    now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    _write_multi(
        fichiers,
        f"{SEP_LOG} DEBUT création base hors-ligne - Envoi par mail de la base "
        f"hors-ligne - {now} {SEP_LOG}",
    )


def log_fin(fichiers, ddebut):
    """Ligne FIN (avec durée) conforme au cahier des charges, dans chaque fichier."""
    now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    duree = datetime.datetime.now() - ddebut
    total = int(duree.total_seconds())
    hh, rem = divmod(total, 3600)
    mm, ss = divmod(rem, 60)
    duree_txt = f"{hh:02d}:{mm:02d}:{ss:02d}"
    _write_multi(
        fichiers,
        f"{SEP_LOG}  FIN  création base hors-ligne - Envoi par mail de la base "
        f"hors-ligne - {now} ({duree_txt}) {SEP_LOG}",
    )


# ---------------------------------------------------------------------------
# Lecture de la configuration (.env)
# ---------------------------------------------------------------------------
def charger_config():
    """Charge les paramètres depuis le fichier .env et vérifie leur présence."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(base_dir, ".env")
    if not os.path.exists(env_path):
        raise RuntimeError(
            f"Fichier de configuration introuvable : {env_path}. "
            "Créez un fichier .env à partir du modèle fourni."
        )
    load_dotenv(env_path)

    valeurs = {
        "RepFicAenvoyer": os.getenv("RepFicAenvoyer"),
        "DateFicAenvoyer": os.getenv("DateFicAenvoyer"),
        "NomFichierWord": os.getenv("NomFichierWord"),
        "BaseUsers": os.getenv("BaseUsers"),
        "FicLog": os.getenv("FicLog"),
        # Optionnel : log global de l'ensemble du traitement (en plus de FicLog)
        "FicLogTrtComplet": os.getenv("FicLogTrtComplet", ""),
        # Optionnel : 1 / true / dryrun => TEST sans envoi réel du mail
        "ModeTest": os.getenv("ModeTest", ""),
        # Optionnel : mot de passe GroupWise saisi automatiquement (directement
        # dans le champ de connexion) lorsque le client GroupWise n'est pas déjà
        # ouvert / connecté au lancement.
        "MDPgroupWise": os.getenv("MDPgroupWise", ""),
        "base_dir": base_dir,
    }

    manquantes = [
        cle
        for cle, val in valeurs.items()
        if cle not in ("base_dir", "ModeTest", "FicLogTrtComplet", "MDPgroupWise") and not val
    ]
    if manquantes:
        raise RuntimeError(
            "Variables manquantes dans le fichier .env : " + ", ".join(manquantes)
        )
    return valeurs


# ---------------------------------------------------------------------------
# Identification du fichier .7z à envoyer
# ---------------------------------------------------------------------------
def identifier_archive(rep_fic, date_fic):
    """Retourne le chemin complet du fichier CM_AAAAMMJJ_*.7z à envoyer.

    - Si date_fic == "today"  -> date du jour (AAAAMMJJ).
    - Sinon (ex. "20260925")   -> cette date exacte.

    Retourne le fichier le plus récent correspondant au motif.
    """
    if date_fic.strip().lower() == "today":
        date_cible = datetime.date.today().strftime("%Y%m%d")
    else:
        date_cible = date_fic.strip()

    motif = os.path.join(rep_fic, f"CM_{date_cible}_*.7z")
    candidats = glob.glob(motif)

    if not candidats:
        raise FileNotFoundError(
            f"Aucune archive .7z trouvée dans '{rep_fic}' pour la date "
            f"{date_cible} (motif : {motif})."
        )

    # Le plus récent si plusieurs fichiers correspondent au motif
    return os.path.abspath(max(candidats, key=os.path.getmtime))


# ---------------------------------------------------------------------------
# Extraction de la date depuis le nom d'une archive CM_AAAAMMJJ_HHMMSS.7z
# ---------------------------------------------------------------------------
def extraire_date_archive(chemin_archive):
    """Extrait et renvoie (AAAAMMJJ, HHMMSS, date_affichage) depuis le nom.

    Le nom attendu est de la forme CM_AAAAMMJJ_HHMMSS.7z (ex. CM_20261002_233513.7z).
    Renvoie :
      - date_aaaammjj : "AAAAMMJJ" (ex. "20261002")
      - heure_hhmmss   : "HHMMSS"   (ex. "233513")
      - date_affichage : "JJ/MM/AAAA" (ex. "02/10/2026") pour le corps du mail.

    Lève RuntimeError si le nom ne correspond pas au motif attendu.
    """
    nom = os.path.basename(chemin_archive)
    # CM_AAAAMMJJ_HHMMSS.7z  ->  capture "AAAAMMJJ" et "HHMMSS"
    m = re.match(r"^CM_(\d{8})_(\d{6})\.7z$", nom)
    if not m:
        raise RuntimeError(
            f"Impossible d'extraire la date du nom d'archive '{nom}'. "
            "Format attendu : CM_AAAAMMJJ_HHMMSS.7z"
        )
    date_aaaammjj = m.group(1)
    heure_hhmmss = m.group(2)
    try:
        date_affichage = datetime.datetime.strptime(
            date_aaaammjj, "%Y%m%d"
        ).strftime("%d/%m/%Y")
    except ValueError:
        date_affichage = f"{date_aaaammjj[6:8]}/{date_aaaammjj[4:6]}/{date_aaaammjj[0:4]}"
    return date_aaaammjj, heure_hhmmss, date_affichage


# ---------------------------------------------------------------------------
# Récupération des destinataires depuis la base Access
# ---------------------------------------------------------------------------
DRIVERS_ACCESS = [
    "Microsoft Access Driver (*.mdb, *.accdb)",
    "Microsoft Access Driver (*.mdb)",
    "Microsoft Access Driver",
]


def _trouver_pilote_access():
    """Renvoie le nom du pilote ODBC Microsoft Access disponible."""
    pilotes = pyodbc.drivers()
    for nom in DRIVERS_ACCESS:
        if nom in pilotes:
            return nom
    raise RuntimeError(
        "Pilote ODBC Microsoft Access introuvable. Installez le pilote "
        "'Microsoft Access Driver (*.mdb, *.accdb)' (Microsoft Access "
        "Database Engine)."
    )


def lire_destinataires(base_users, champ_bool):
    """Lit et renvoie la liste des adresses mail des destinataires.

    Requête calquée sur le VBA de référence :
        SELECT TbUsers.mail FROM TbUsers
        WHERE actif_CM=True AND <champ_bool>=True

    :param champ_bool: "EnvoiMail_TT_To_CM" (TO) ou "EnvoiMail_TT_Cc_CM" (CC)
    :return: liste (éventuellement vide) d'adresses mail distinctes, sans vide.
    """
    pilote = _trouver_pilote_access()
    cnx = pyodbc.connect(f"DRIVER={{{pilote}}};DBQ={base_users};")
    try:
        cur = cnx.cursor()
        sql = (
            "SELECT TbUsers.mail FROM TbUsers "
            "WHERE (((TbUsers.actif_CM)=True) AND "
            f"((TbUsers.{champ_bool})=True));"
        )
        cur.execute(sql)

        mails = []
        for row in cur.fetchall():
            mail = (row[0] or "").strip()
            if mail and mail not in mails:
                mails.append(mail)
        return mails
    finally:
        cnx.close()


# ---------------------------------------------------------------------------
# Enveloppe de pilotage GroupWise (tokens GroupwiseCommander)
# ---------------------------------------------------------------------------
def _echapper_pour_cmd(texte, mono_ligne=True):
    """Prépare une chaîne destinée à être passée comme argument d'un token
    GroupWise entre guillemets doubles.

    - double les backslashes :  \\  ->  \\\\   (chemins Windows)
    - échappe les guillemets :  "   ->  \\"
    - si mono_ligne : remplace les retours à la ligne (CR/LF) et espaces
      multiples par un espace unique, afin de garder la commande sur UNE
      ligne (le parseur GroupWise/DDE ignore une commande à la syntaxe
      invalide, notamment si elle est coupée par un retour à la ligne).
    """
    if mono_ligne:
        import re as _re
        texte = _re.sub(r"\s+", " ", texte).strip()
    texte = texte.replace("\\", "\\\\").replace('"', '\\"')
    return texte


def _decouper_corps(texte, taille_max=200):
    """Découpe `texte` (déjà aplati sur une ligne) en segments de taille
    <= `taille_max`, sans couper un mot au milieu.

    GroupWise rejette une commande `TextSetMessage(...)` trop longue
    (seuil constaté d'environ 350-400 caractères selon le contenu). On
    contourne la limite en envoyant plusieurs `TextSetMessage` successifs
    avec `Append`, chaque segment restant bien sous le seuil.
    """
    mots = texte.split(" ")
    segments = []
    courant = ""
    for mot in mots:
        if not courant:
            courant = mot
        elif len(courant) + 1 + len(mot) <= taille_max:
            courant += " " + mot
        else:
            segments.append(courant)
            courant = mot
    if courant:
        segments.append(courant)
    return segments


# ---------------------------------------------------------------------------
# Démarrage / connexion automatique du client GroupWise
# ---------------------------------------------------------------------------
# Le client de messagerie GroupWise (OpenText GroupWise CE) est lancé via son
# exécutable `grpwise.exe`. On recherche cet exécutable dans les dossiers
# d'installation classiques (32 bits puis 64 bits). La détection de l'état
# "GroupWise ouvert / fenêtre de connexion" se fait par le TITRE des fenêtres
# Windows (qui contient "GroupWise", ex. "Mot de passe GroupWise" ou
# "OpenText GroupWise - Boîte aux lettres") et non par le nom du processus :
# ce dernier s'avère peu fiable (module inaccessible, client "GroupWise CE").
GROUPWISE_EMPLACEMENTS = [
    r"C:\Program Files (x86)\Novell\GroupWise\grpwise.exe",
    r"C:\Program Files\Novell\GroupWise\grpwise.exe",
    r"C:\Program Files (x86)\GroupWise\grpwise.exe",
    r"C:\Program Files\GroupWise\grpwise.exe",
]

# Délais (en secondes) utilisés pour l'automatisation d'interface :
# - attente de l'apparition de la fenêtre de connexion GroupWise après lancement ;
# - pause après la frappe du mot de passe pour laisser GroupWise s'authentifier.
ATTENTE_FENETRE_CONNEXION_S = 30
ATTENTE_AUTHENTIFICATION_S = 12

# Sous-chaînes (insensibles à la casse) présentes dans le TITRE de la fenêtre
# de connexion GroupWise (ex. "Mot de passe GroupWise") et dans la fenêtre de
# la boîte aux lettres (ex. "OpenText GroupWise - Boîte aux lettres").
CLE_TITRE_GROUPWISE = "groupwise"
CLE_TITRE_MOT_DE_PASSE = ("mot de passe", "password", "motdepasse")
CLE_TITRE_BOITE_AUX_LETTRES = ("boîte aux lettres", "boite aux lettres", "mailbox")


def _trouver_executable_groupwise():
    """Retourne le chemin de `grpwise.exe` s'il existe, sinon ``None``."""
    for chemin in GROUPWISE_EMPLACEMENTS:
        if os.path.isfile(chemin):
            return chemin
    return None


def _windows_visibles_avec(cle):
    """Retourne la liste des (hwnd, titre) des fenêtres visibles de premier
    niveau dont le titre (minuscules) contient `cle` (insensible à la casse)."""
    resultats = []

    def _callback(hwnd, _unused):
        if not win32gui.IsWindowVisible(hwnd):
            return
        try:
            titre = win32gui.GetWindowText(hwnd)
        except Exception:
            return
        if titre and cle in titre.lower():
            resultats.append((hwnd, titre))

    try:
        win32gui.EnumWindows(_callback, None)
    except Exception:  # pragma: no cover
        return []
    return resultats


def _groupwise_est_ouvert():
    """Renvoie ``True`` si une fenêtre visible du client GroupWise existe.

    Une fenêtre GroupWise est détectée dès que son titre contient "GroupWise"
    (fenêtre de connexion, boîte aux lettres, etc.). Si le client est déjà
    ouvert, on continue la logique actuelle sans retaper de mot de passe.
    """
    return bool(_windows_visibles_avec(CLE_TITRE_GROUPWISE))


def _hwnd_fenetre_connexion():
    """Retourne le hwnd de la fenêtre de connexion GroupWise (demande de mot de
    passe), sinon ``None``.

    La fenêtre ciblée est celle dont le titre contient "GroupWise" ET l'un des
    mots-clés "mot de passe"/"password". Si elle n'est pas trouvée mais qu'une
    fenêtre "Boîte aux lettres" est déjà présente (GroupWise déjà connecté),
    on renvoie ``None`` sans lever : il n'y a alors rien à saisir.
    """
    for cle in CLE_TITRE_MOT_DE_PASSE:
        for hwnd, titre in _windows_visibles_avec(cle):
            if CLE_TITRE_GROUPWISE in titre.lower():
                return hwnd
    return None


def _groupwise_deja_connecte():
    """Renvoie ``True`` si la fenêtre de la boîte aux lettres GroupWise est
    déjà affichée (GroupWise déjà connecté et prêt)."""
    for cle in CLE_TITRE_BOITE_AUX_LETTRES:
        for hwnd, titre in _windows_visibles_avec(cle):
            if CLE_TITRE_GROUPWISE in titre.lower():
                return True
    return False


def _controles_connexion(hwnd_fenetre):
    """Retourne ``(hwnd_edit, hwnd_bouton_ok)`` des contrôles enfants de la
    fenêtre de connexion ``hwnd_fenetre``, ou ``(None, None)`` si introuvables.

    La fenêtre de connexion "Mot de passe GroupWise" contient un champ de
    saisie (classe ``Edit``) qui reçoit le mot de passe, et un bouton ``OK``
    (classe ``Button`` dont le libellé commence par "OK"). Cibler directement
    ces contrôles (via ``SendMessage``/``WM_SETTEXT`` + ``BM_CLICK``) est bien
    plus fiable que d'envoyer des touches via SendKeys, qui dépend du focus.
    """
    hwnd_edit = None
    hwnd_ok = None

    def _callback(hwnd, _unused):
        nonlocal hwnd_edit, hwnd_ok
        try:
            classe = win32gui.GetClassName(hwnd)
            texte = win32gui.GetWindowText(hwnd)
        except Exception:
            return
        if classe == "Edit" and hwnd_edit is None:
            hwnd_edit = hwnd
        elif classe == "Button" and texte.strip().lower().startswith("ok"):
            hwnd_ok = hwnd

    try:
        win32gui.EnumChildWindows(hwnd_fenetre, _callback, None)
    except Exception:  # pragma: no cover
        return None, None
    return hwnd_edit, hwnd_ok


def _lancer_groupwise():
    """Lance le client GroupWise s'il n'est pas déjà ouvert.

    Recherche `grpwise.exe` dans les emplacements connus. Si introuvable,
    on lance quand même GroupwiseCommander (qui démarre le client à sa façon).
    """
    executable = _trouver_executable_groupwise()
    if executable:
        # Démarrage sans attendre le retour du processus
        subprocess.Popen([executable])
    else:
        # Fallback : l'objet GroupwiseCommander peut démarrer le client lui-même.
        try:
            win32com.client.Dispatch("GroupwiseCommander")
        except Exception:
            # On laisse l'appelant tenter de continuer ; l'erreur remontera
            # ensuite au moment de l'envoi si GroupWise n'a pas pu démarrer.
            return


def _saisir_mot_de_passe_groupwise(mdp):
    """Attend la fenêtre de connexion GroupWise et y saisit le mot de passe.

    - Attend (jusqu'à ATTENTE_FENETRE_CONNEXION_S) l'apparition de la fenêtre
      de connexion dont le titre contient "mot de passe"/"password" ;
    - envoie le mot de passe directement dans le champ "Edit" de cette fenêtre
      (WM_SETTEXT) puis clique sur le bouton OK (BM_CLICK), indépendamment du
      focus ; en secours, retombe sur un envoi de touches SendKeys + Entrée ;
    - laisse un court délai pour que GroupWise s'authentifie.

    Cas particuliers gérés :
      * GroupWise se connecte sans demander de mot de passe (fenêtre "Boîte
        aux lettres" déjà affichée) → rien n'est saisi, on retourne simplement ;
      * Aucune fenêtre de connexion n'apparaît dans le délai → erreur claire.

    :param mdp: mot de passe GroupWise (variable MDPgroupWise du .env).
    """
    coque = win32com.client.Dispatch("WScript.Shell")

    # 1. Attente de l'apparition de la fenêtre de connexion
    hwnd = None
    debut = time.time()
    while time.time() - debut < ATTENTE_FENETRE_CONNEXION_S:
        if _groupwise_deja_connecte():
            # GroupWise s'est connecté sans mot de passe : rien à saisir.
            return
        hwnd = _hwnd_fenetre_connexion()
        if hwnd is not None:
            break
        time.sleep(0.5)

    if hwnd is None:
        raise RuntimeError(
            "GroupWise a été lancé mais aucune fenêtre de connexion (mot de "
            "passe) n'est apparue dans le délai imparti. Mot de passe non saisi."
        )

    # 2. Mise au premier plan de la fenêtre GroupWise (améliore la fiabilité si
    #    on doit retomber sur l'envoi de touches).
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        # Le focus peut parfois être refusé ; on tente quand même la saisie.
        pass
    time.sleep(1.0)

    # 3. Saisie du mot de passe puis validation.
    #    Approche privilégiée : envoyer le mot de passe directement au champ
    #    "Edit" de la fenêtre (WM_SETTEXT) puis cliquer sur le bouton "OK"
    #    (BM_CLICK). C'est indépendant du focus et donc bien plus fiable que
    #    SendKeys. Si les contrôles ne sont pas trouvés, on retombe sur
    #    SendKeys (focus de la fenêtre déjà mis au premier plan en étape 2).
    hwnd_edit, hwnd_ok = _controles_connexion(hwnd)
    if hwnd_edit is not None and hwnd_ok is not None:
        # Le contenu du champ Edit n'apparaît pas avec GetWindowText car c'est
        # un champ mot de passe (masqué) ; WM_SETTEXT le remplit correctement.
        win32gui.SendMessage(hwnd_edit, win32con.WM_SETTEXT, 0, str(mdp))
        time.sleep(0.5)
        win32gui.SendMessage(hwnd_ok, win32con.BM_CLICK, 0, 0)
    else:
        # Fallback par envoi de touches (moins fiable mais couvre les cas où la
        # fenêtre a une structure différente).
        mdp_safe = str(mdp).replace("{", "{{").replace("}", "}}")
        coque.SendKeys(mdp_safe)
        time.sleep(0.5)
        coque.SendKeys("{ENTER}")

    # 4. Délai pour laisser GroupWise s'authentifier
    time.sleep(ATTENTE_AUTHENTIFICATION_S)


class GroupWiseCommander:
    """Encapsule l'objet OLE GroupwiseCommander et l'exécution des tokens."""

    def __init__(self):
        try:
            self.commander = win32com.client.Dispatch("GroupwiseCommander")
        except Exception as exc:  # pragma: no cover
            raise RuntimeError(
                "Échec de l'initialisation de l'objet COM GroupwiseCommander : "
                f"{exc}\nVérifiez que le client GroupWise 32 bits est installé "
                "et enregistré, et que ce script tourne en Python 32 bits (x86)."
            ) from exc

    def connecter(self, mdp):
        """Garantit que le client GroupWise est ouvert ET connecté.

        - Si la boîte aux lettres GroupWise est déjà affichée (session
          connectée), ne fait rien (on reprend la logique actuelle).
        - Sinon, s'assure que le client est lancé, attend la fenêtre de
          connexion, y saisit le mot de passe ``mdp`` (variable MDPgroupWise du
          .env) et valide (via le bouton OK), puis laisse GroupWise
          s'authentifier.

        :param mdp: mot de passe GroupWise (chaîne) ; s'il est vide et qu'une
                    connexion est nécessaire, on lève une erreur claire.
        """
        if _groupwise_deja_connecte():
            # GroupWise est déjà connecté (boîte aux lettres ouverte) : rien à
            # faire, on reprend exactement la logique actuelle.
            return

        if not mdp:
            raise RuntimeError(
                "GroupWise n'est pas connecté et la variable 'MDPgroupWise' "
                "n'est pas renseignée dans le fichier .env. Impossible de "
                "saisir le mot de passe automatiquement."
            )

        print("GroupWise non connecté : vérification/lancement du client et "
              "saisie du mot de passe...")
        # Si le client n'est même pas ouvert (aucune fenêtre GroupWise visible),
        # on le lance. S'il est déjà ouvert mais en attente de mot de passe
        # (fenêtre "Mot de passe GroupWise"), on lance quand même la saisie.
        if not _groupwise_est_ouvert():
            _lancer_groupwise()
        _saisir_mot_de_passe_groupwise(mdp)

    def executer(self, commande):
        """Exécute un token GroupWise et lève une erreur en cas d'échec.

        Execute(CMD, RetString) renvoie 1 en cas de succès, 0 en cas d'échec.
        """
        ret_string = ""
        try:
            res = self.commander.Execute(commande, ret_string)
            if isinstance(res, tuple):
                retval = res[0]
                if len(res) > 1:
                    ret_string = res[1]
            else:
                retval = res
        except Exception as exc:  # pragma: no cover
            raise RuntimeError(
                f"Exception Python lors du token GroupWise '{commande}' : {exc}"
            ) from exc
        if not retval:
            raise RuntimeError(
                f"Échec du token GroupWise '{commande}'. Détail : {ret_string}"
            )
        return ret_string

    def envoyer_mail(self, to_liste, cc_liste, sujet, corps, pieces_jointes):
        """Compose et envoie le courriel via les tokens GroupWise."""
        # 1. Nouveau brouillon
        self.executer("NewMail()")

        # 2. Destinataires TO (liste séparée par des virgules)
        if to_liste:
            to_texte = ",".join(to_liste)
            self.executer(f'TextSetTo("{to_texte}")')

        # 3. Destinataires CC
        if cc_liste:
            cc_texte = ",".join(cc_liste)
            self.executer(f'TextSetCC("{cc_texte}")')

        # 4. Objet du message
        self.executer(f'TextSetSubject("{sujet}")')

        # 5. Corps du message injecté via le presse-papier Windows (Option 1).
        #    Cette méthode préserve fidèlement les sauts de ligne (CRLF)
        #    et s'affranchit des limites de longueur et de syntaxe du parseur
        #    DDE de GroupWise (pas de troncature ni de commande coupée par \n).
        texte_presse_papier = corps.replace("\r\n", "\n").replace("\n", "\r\n")
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(
                win32clipboard.CF_UNICODETEXT, texte_presse_papier
            )
        finally:
            win32clipboard.CloseClipboard()

        # Donne le focus au corps du message (10 = Message) et colle le texte
        self.executer("FocusSet(10)")
        self.executer("EditPaste()")

        # 6. Pièces jointes (AttachmentAdd prend le chemin complet du fichier).
        #    On double les backslashes du chemin et on échappe les guillemets,
        #    sans écraser les espaces éventuels du nom de fichier.
        for chemin in pieces_jointes:
            if not os.path.isfile(chemin):
                raise FileNotFoundError(f"Pièce jointe introuvable : {chemin}")
            chemin_safe = _echapper_pour_cmd(chemin, mono_ligne=False)
            self.executer(f'AttachmentAdd("{chemin_safe}")')

        # 7. Envoi du message (ItemSend() sans argument envoie le message ouvert)
        self.executer('ItemSend()')


# ---------------------------------------------------------------------------
# Point d'entrée
# ---------------------------------------------------------------------------
def main():
    """Pilote l'ensemble du processus et renvoie un code de retour."""
    ddebut = datetime.datetime.now()
    config = None
    fichier_logs = None

    try:
        config = charger_config()

        rep_fic = config["RepFicAenvoyer"]
        date_fic = config["DateFicAenvoyer"]
        base_users = config["BaseUsers"]
        fic_log = config["FicLog"]

        # Liste des fichiers de log : FicLog (toujours) + FicLogTrtComplet (si défini)
        fichier_logs = [fic_log]
        ficlog_trt = config.get("FicLogTrtComplet", "").strip()
        if ficlog_trt:
            fichier_logs.append(ficlog_trt)

        # Ligne DEBUT dans le(s) fichier(s) de log
        log_debut(fichier_logs)

        # --- 1. Identification de l'archive .7z à envoyer ---
        archive = identifier_archive(rep_fic, date_fic)

        # --- 2. Lecture des destinataires TO / CC ---
        to_liste = lire_destinataires(base_users, "EnvoiMail_TT_To_CM")
        if not to_liste:
            _write_multi(fichier_logs, "pas de destinataires TO ")
        cc_liste = lire_destinataires(base_users, "EnvoiMail_TT_Cc_CM")
        if not cc_liste:
            _write_multi(fichier_logs, "pas de destinataires CC ")

        # --- 3. Pièces jointes à envoyer ---
        #   - l'archive .7z identifiée ci-dessus
        #   - le fichier Word (mode opératoire) dont le nom vient du .env,
        #     situé dans le dossier RepFicAenvoyer
        fichier_word = os.path.join(rep_fic, config["NomFichierWord"])
        if not os.path.isfile(fichier_word):
            raise FileNotFoundError(
                f"Fichier Word introuvable dans '{rep_fic}' : "
                f"{config['NomFichierWord']}"
            )
        pieces_jointes = [archive, fichier_word]

        # Date affichée dans le corps du mail, extraite du nom de l'archive
        _, _, date_archive = extraire_date_archive(archive)

        # --- Infos pertinentes intercalées entre DEBUT et FIN (log global) ---
        _write_multi(fichier_logs, f"Envoi du fichier (archive) : {os.path.basename(archive)}")
        _write_multi(fichier_logs, f"Date de la base hors-ligne : {date_archive}")
        _write_multi(
            fichier_logs,
            f"Destinataires TO ({len(to_liste)}) : " + (", ".join(to_liste) if to_liste else "(aucun)"),
        )
        _write_multi(
            fichier_logs,
            f"Destinataires CC ({len(cc_liste)}) : " + (", ".join(cc_liste) if cc_liste else "(aucun)"),
        )
        _write_multi(fichier_logs, f"Fichier Word joint : {os.path.basename(fichier_word)}")

        # --- Corps du mail (texte simple, la date est rendue dynamiquement) ---
        corps_texte = (
            f"Voici la base cas marquant \"hors ligne\" du {date_archive} au soir "
            "ainsi que son mode opératoire.\n\n"
            "Si vous désirez modifier un cas marquant, vous devez impérativement, "
            "utiliser la dernière base envoyée, il s'agit de la plus à jour.\n\n"
            "Dans ce cas, vos modifications seront exportées automatiquement, via un fichier Excel, dans le répertoire suivant : "
            "G:\\DM-SURVEIL\\MEDICAMENTS\\BASES\\BASE_CM\\Import_TT\\ImportExcel_TT\\PROD\\aFaire\\\n\n"
            "Frédéric Rannou"
        )

        mode_test = str(config.get("ModeTest", "")).strip().lower() in (
            "1", "true", "oui", "o", "yes", "y", "dryrun", "test"
        )

        if mode_test:
            # ====== MODE TEST (dry-run) : aucune commande GroupWise n'est exécutée ======
            print("=" * 60)
            print("MODE TEST (dry-run) : validation sans envoi de mail.")
            print("-" * 60)
            print(f"Archive .7z identifiée : {archive}")
            print(f"Fichier Word            : {fichier_word}")
            print(f"Date extraite (archive) : {date_archive}")
            print(f"Destinataires TO ({len(to_liste)}) : {', '.join(to_liste) if to_liste else '(aucun)'}")
            print(f"Destinataires CC ({len(cc_liste)}) : {', '.join(cc_liste) if cc_liste else '(aucun)'}")
            print(f"Objet du mail          : {SUJET_MAIL}")
            print(f"Pièces jointes         : {archive}")
            print(f"                         {fichier_word}")
            print("-" * 60)
            # Reconstruit les commandes GroupWise (sans les exécuter) pour
            # vérifier que le corps et les pièces jointes sont bien formatés.
            print("Commandes GroupWise qui seraient envoyées :")
            print(f"  NewMail()")
            print(f"  TextSetTo(\"{_echapper_pour_cmd(','.join(to_liste), mono_ligne=False) if to_liste else ''}\")")
            print(f"  TextSetSubject(\"{_echapper_pour_cmd(SUJET_MAIL, mono_ligne=False)}\")")
            print("  [Copie du corps dans le presse-papier Windows (CF_UNICODETEXT) avec sauts de ligne réels]")
            print("  FocusSet(10)")
            print("  EditPaste()")
            for chemin in pieces_jointes:
                print(f"  AttachmentAdd(\"{_echapper_pour_cmd(chemin, mono_ligne=False)}\")")
            print(f"  ItemSend()")
            print("-" * 60)
            print("Le mail n'a PAS été envoyé (mode test).")
            print("=" * 60)
            log_fin(fichier_logs, ddebut)
            return 0

        # --- 4. Envoi via GroupWise ---
        commander = GroupWiseCommander()
        # Si la session GroupWise n'est pas déjà connectée (boîte aux lettres
        # ouverte), on s'assure de la connexion (lancement + saisie du mot de
        # passe MDPgroupWise le cas échéant) avant d'envoyer le mail.
        mdp_groupwise = config.get("MDPgroupWise", "")
        if not _groupwise_deja_connecte():
            if mdp_groupwise:
                _write_multi(
                    fichier_logs,
                    "GroupWise non connecté : saisie automatique du mot de passe "
                    "MDPgroupWise avant envoi.",
                )
            commander.connecter(mdp_groupwise)
        commander.envoyer_mail(
            to_liste=to_liste,
            cc_liste=cc_liste,
            sujet=SUJET_MAIL,
            corps=corps_texte,
            pieces_jointes=pieces_jointes,
        )

        # Nettoyage du fichier HTML temporaire (plus de fichier HTML généré)
        # try:
        #     os.remove(chemin_html)
        # except OSError:
        #     pass

        # --- 5. Fin de traitement (ligne FIN + durée) ---
        log_fin(fichier_logs, ddebut)

        print("Envoi effectué avec succès.")
        return 0

    except Exception as exc:  # pragma: no cover
        # Journalise l'erreur (et une ligne FIN) dans le(s) fichier(s) de log
        # si possible. On réutilise la liste déjà construite (FicLog +
        # FicLogTrtComplet), sinon on retombe sur FicLog seul.
        try:
            logs_except = fichier_logs
            if not logs_except and config and config.get("FicLog"):
                logs_except = [config["FicLog"]]
            if logs_except:
                _write_multi(
                    logs_except,
                    f"ERREUR envoi par mail de la base hors-ligne : {exc}",
                )
                log_fin(logs_except, ddebut)
        except Exception:
            pass

        sys.stderr.write(f"\n[ERREUR] {exc}\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
