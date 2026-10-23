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
    RepFicAenvoyer     = dossier contenant l'archive .7z (finissant par un backslash)
    DateFicAenvoyer    = "today" (date du jour) OU date exacte "AAAAMMJJ"
    BaseUsers          = chemin complet de la base Access (table TbUsers)
    FicLog             = chemin complet du fichier de log commun (mode append)

Sortie :
    Code de retour 0 si succès, 1 en cas d'erreur.
"""

import os
import sys
import glob
import locale
import datetime

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
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "ERREUR : la bibliothèque 'pywin32' est requise (module win32com).\n"
        "Installez-la avec : pip install pywin32\n"
    )
    sys.exit(1)

# ---------------------------------------------------------------------------
# Constantes du cahier des charges
# ---------------------------------------------------------------------------
SEP_LOG = "*" * 55

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
    """
    enc = locale.getpreferredencoding(False) or "utf-8"
    with open(fic_log, "a", encoding=enc) as fh:
        fh.write(texte + os.linesep)


def log_debut(fic_log):
    """Ligne DEBUT conforme au cahier des charges."""
    now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    _writelog(
        fic_log,
        f"{SEP_LOG} DEBUT création base hors-ligne - Envoi par mail de la base "
        f"hors-ligne - {now} {SEP_LOG}",
    )


def log_fin(fic_log, ddebut):
    """Ligne FIN (avec durée) conforme au cahier des charges."""
    now = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    duree = datetime.datetime.now() - ddebut
    total = int(duree.total_seconds())
    hh, rem = divmod(total, 3600)
    mm, ss = divmod(rem, 60)
    duree_txt = f"{hh:02d}:{mm:02d}:{ss:02d}"
    _writelog(
        fic_log,
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
        "BaseUsers": os.getenv("BaseUsers"),
        "FicLog": os.getenv("FicLog"),
        # Optionnel : 1 / true / dryrun => TEST sans envoi réel du mail
        "ModeTest": os.getenv("ModeTest", ""),
        "base_dir": base_dir,
    }

    manquantes = [
        cle
        for cle, val in valeurs.items()
        if cle not in ("base_dir", "ModeTest") and not val
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
# Génération du fichier HTML joint (corps HTML), reproduit du VBA
# ---------------------------------------------------------------------------
def generer_html(base_dir, date_ficjour):
    """Génère le fichier 'text.htm' reproduisant le HTML du VBA d'origine.

    Retourne le chemin complet du fichier généré.
    """
    # date_ficjour paramétrée (non utilisée pour le texte, la date du jour
    # est affichée comme dans le VBA via sDateDuJour)
    jour = datetime.date.today().strftime("%d/%m/%Y")

    chemin_html = os.path.join(base_dir, "text.htm")

    entete = (
        '<meta name="Generator" content="Novell Groupwise Client '
        '(Version 14.2.1  Build: 124595)">'
        '<meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>'
    )

    html = (
        "<html><head>\n"
        + entete
        + "\n"
        "<body>\n"
        f'<div>Voici la base cas marquant "hors ligne" du {jour} au soir '
        "ainsi que son mode opératoire.</div><div><br></div>\n"
        "<div>Si vous désirez modifier un cas marquant, vous devez "
        "impérativement, utiliser la dernière base envoyée, il s'agit de la "
        "plus à jour.</div><div><br></div>\n"
        "<div>Il est également nécessaire d’envoyer les fichiers générés le "
        "jour même, et ainsi réduire le risque de collisions de modification "
        "(modifications du même cas marquant par deux utilisateurs "
        "différents).</div><div><br></div>\n"
        f'<div>Merci de copier les fichiers Excel dans le répertoire suivant : '
        f'<a href="{LIEN_COPIE_EXCEL}">{LIEN_COPIE_EXCEL}</a></div>'
        "<div><br></div>\n"
        "<div>Frédéric Rannou<br></div></div>\n"
        "</body></html>\n"
    )

    with open(chemin_html, "w", encoding="utf-8") as fh:
        fh.write(html)

    return chemin_html


# ---------------------------------------------------------------------------
# Enveloppe de pilotage GroupWise (tokens GroupwiseCommander)
# ---------------------------------------------------------------------------
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

        # 5. Corps du message (le VBA envoie "Frédéric Rannou" comme corps,
        #    le contenu détaillé étant généré dans text.htm joint en pièce jointe)
        corps_echappe = corps.replace('"', '\\"')
        self.executer(f'TextSetMessage("{corps_echappe}")')

        # 6. Pièces jointes (AttachmentAdd prend directement le chemin du fichier)
        for chemin in pieces_jointes:
            if not os.path.isfile(chemin):
                raise FileNotFoundError(f"Pièce jointe introuvable : {chemin}")
            chemin_safe = chemin.replace('"', '\\"')
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

    try:
        config = charger_config()

        rep_fic = config["RepFicAenvoyer"]
        date_fic = config["DateFicAenvoyer"]
        base_users = config["BaseUsers"]
        fic_log = config["FicLog"]

        # Ligne DEBUT dans le log commun
        log_debut(fic_log)

        # --- 1. Identification de l'archive .7z à envoyer ---
        archive = identifier_archive(rep_fic, date_fic)

        # --- 2. Lecture des destinataires TO / CC ---
        to_liste = lire_destinataires(base_users, "EnvoiMail_TT_To_CM")
        if not to_liste:
            _writelog(fic_log, "pas de destinataires TO ")
        cc_liste = lire_destinataires(base_users, "EnvoiMail_TT_Cc_CM")
        if not cc_liste:
            _writelog(fic_log, "pas de destinataires CC ")

        # --- 3. Génération du fichier HTML joint ---
        chemin_html = generer_html(config["base_dir"], date_fic)

        mode_test = str(config.get("ModeTest", "")).strip().lower() in (
            "1", "true", "oui", "o", "yes", "y", "dryrun", "test"
        )

        if mode_test:
            # ====== MODE TEST (dry-run) : aucune commande GroupWise n'est exécutée ======
            print("=" * 60)
            print("MODE TEST (dry-run) : validation sans envoi de mail.")
            print("-" * 60)
            print(f"Archive .7z identifiée : {archive}")
            print(f"Fichier HTML généré    : {chemin_html}")
            print(f"Destinataires TO ({len(to_liste)}) : {', '.join(to_liste) if to_liste else '(aucun)'}")
            print(f"Destinataires CC ({len(cc_liste)}) : {', '.join(cc_liste) if cc_liste else '(aucun)'}")
            print(f"Objet du mail          : {SUJET_MAIL}")
            print(f"Pièces jointes         : {chemin_html}, {archive}")
            print("-" * 60)
            print("Le mail n'a PAS été envoyé (mode test).")
            print("=" * 60)
            # Nettoyage du fichier HTML temporaire
            try:
                os.remove(chemin_html)
            except OSError:
                pass
            log_fin(fic_log, ddebut)
            return 0

        # --- 4. Envoi via GroupWise ---
        commander = GroupWiseCommander()
        commander.envoyer_mail(
            to_liste=to_liste,
            cc_liste=cc_liste,
            sujet=SUJET_MAIL,
            corps="Frédéric Rannou",
            pieces_jointes=[chemin_html, archive],
        )

        # Nettoyage du fichier HTML temporaire
        try:
            os.remove(chemin_html)
        except OSError:
            pass

        # --- 5. Fin de traitement (ligne FIN + durée) ---
        log_fin(fic_log, ddebut)

        print("Envoi effectué avec succès.")
        return 0

    except Exception as exc:  # pragma: no cover
        # Journalise l'erreur (et une ligne FIN) dans le log commun si possible
        try:
            if config and config.get("FicLog"):
                _writelog(
                    config["FicLog"],
                    f"ERREUR envoi par mail de la base hors-ligne : {exc}",
                )
                log_fin(config["FicLog"], ddebut)
        except Exception:
            pass

        sys.stderr.write(f"\n[ERREUR] {exc}\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
