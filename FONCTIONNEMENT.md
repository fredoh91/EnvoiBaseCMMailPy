# FONCTIONNEMENT — envoi_base_cm.py

Documentation technique du script **`envoi_base_cm.py`**, destinée à un
programmeur (novice en Python). Elle décrit, étape par étape, ce que fait le
script, dans quel ordre il appelle ses fonctions, et ce que fait chacune d'elles.

---

## 1. Rôle du script

Ce script Python **prend le relais de la macro Access** pour réaliser l'**étape 5**
d'une chaîne de traitement quotidienne lancée vers **23h30** :

| Étape | Traitement | Outil |
|------:|------------|-------|
| 1 | Lancer le traitement quotidien (23h30) | Access |
| 2 | Générer une copie de la base Access | Access |
| 3 | Copier les tables liées dans la base locale | Access |
| 4 | Compresser la base avec 7-Zip → `CM_AAAAMMJJ_HHMMSS.7z` | Access |
| 5 | **Envoyer l'archive `.7z` par courriel GroupWise** | **Python (ici)** |

Le script est lancé par le **Planificateur de tâches Windows**, 15 à 20 minutes
après le démarrage du traitement Access (pour laisser le temps à l'archive
d'être générée).

Il retourne un **code de retour** :
- `0` → succès ;
- `1` → erreur (journalisée puis renvoyée au planificateur).

---

## 2. Vue d'ensemble du flux

```
┌─────────────────────────────────────────────────────────────┐
│  main()  ── point d'entrée (ligne 369)                      │
│                                                             │
│  1. charger_config()          → lit le fichier .env         │
│  2. log_debut()               → ligne DEBUT dans le log     │
│  3. identifier_archive()      → trouve le CM_*.7z du jour   │
│  4. lire_destinataires() × 2  → listes TO et CC (Access)    │
│  5. generer_html()            → crée le fichier text.htm    │
│  6. [ModeTest ?]              → dry-run ou envoi réel       │
│  7. GroupWiseCommander()      → envoi via GroupWise         │
│  8. log_fin()                 → ligne FIN + durée           │
└─────────────────────────────────────────────────────────────┘
```

---

## 3. Chemin d'exécution détaillé (ordre des fonctions)

### 3.1 Point d'entrée : le bloc `if __name__ == "__main__"`

```python
if __name__ == "__main__":
    sys.exit(main())
```

Quand on exécute `python envoi_base_cm.py`, Python charge le fichier puis
exécute cette dernière ligne. `main()` est appelée et son résultat sert de code
de sortie du processus (`sys.exit`).

> **Notion Python** : `if __name__ == "__main__"` permet de n'exécuter le code
> que si le fichier est lancé directement (et non importé comme module).


### 3.2 `main()` — le pilote (ligne 369)

C'est la **fonction centrale** qui orchestre tout. Dans l'ordre :

1. **`ddebut = datetime.datetime.now()`** — mémorise l'heure de départ (servira
   à calculer la durée dans le log de fin).

2. **`config = charger_config()`** — charge les paramètres depuis `.env`
   (voir §4.1).

3. **`log_debut(fic_log)`** — écrit la ligne `DEBUT...` dans le fichier de log
   (voir §4.2).

4. **`archive = identifier_archive(rep_fic, date_fic)`** — identifie l'archive
   `.7z` du jour (voir §4.3).

5. **`to_liste = lire_destinataires(base_users, "EnvoiMail_TT_To_CM")`** — lit la
   liste des destinataires **TO** dans la base Access (voir §4.4).
   - même appel avec `"EnvoiMail_TT_Cc_CM"` pour les destinataires **CC**.
   - si une liste est vide, une ligne `pas de destinataires TO/CC` est écrite au log.

6. **`chemin_html = generer_html(config["base_dir"], date_fic)`** — génère le
   fichier `text.htm` joint (voir §4.5).

7. **Détection du mode test** : le script vérifie la variable optionnelle
   `ModeTest` du `.env`.
   - **Si mode test actif** : il affiche à l'écran l'archive trouvée, le HTML
     généré, les destinataires, l'objet et les pièces jointes, *sans envoyer de
     mail*, nettoie le HTML, écrit la ligne FIN et retourne `0`. (Voir §5.)
   - **Sinon** (envoi réel) : il poursuit.

8. **`commander = GroupWiseCommander()`** — instancie l'objet COM
   `GroupwiseCommander` (voir §4.6).

9. **`commander.envoyer_mail(...)`** — compose et envoie le courriel (voir §4.7).

10. **Nettoyage** : le fichier `text.htm` temporaire est supprimé
    (`os.remove`), dans un bloc maîtrisé (si la suppression échoue, on ignore
    l'erreur car ce n'est pas bloquant).

11. **`log_fin(fic_log, ddebut)`** — écrit la ligne `FIN...` (avec durée) dans le
    log (voir §4.8).

12. **`print("Envoi effectué avec succès.")`** et **`return 0`**.

### 3.3 Gestion des erreurs dans `main()`

Tout le corps de `main()` est enveloppé dans un `try / except Exception` :

- En cas d'**erreur** : le script tente d'écrire une ligne
  `ERREUR envoi par mail...` dans le fichier de log (si `FicLog` est connu), puis
  une ligne `FIN`, affiche `[ERREUR] <message>` sur la sortie d'erreur standard
  (`sys.stderr`) et retourne **`1`**.


---

## 4. Description des fonctions

### 4.1 `charger_config()` (ligne 129) — lecture du `.env`

- Calcule le répertoire du script (`base_dir`) via
  `os.path.dirname(os.path.abspath(__file__))` :
  - `__file__` = chemin du fichier courant ;
  - `os.path.abspath` = version absolue ;
  - `os.path.dirname` = le dossier le contenant.
- Vérifie que `.env` existe à côté du script. Sinon, lève une `RuntimeError`.
- **`load_dotenv(env_path)`** (bibliothèque `python-dotenv`) : charge les
  variables du fichier `.env` dans les variables d'environnement du processus.
- Relit les valeurs via `os.getenv(...)` et les range dans un dictionnaire :
  `RepFicAenvoyer`, `DateFicAenvoyer`, `BaseUsers`, `FicLog`, `ModeTest`
  (optionnel) et `base_dir`.
- Contrôle que les 4 variables obligatoires sont présentes (liste
  `manquantes`) ; sinon lève une `RuntimeError` listant celles manquantes.
- **Retourne** le dictionnaire de configuration.

### 4.2 `log_debut(fic_log)` (ligne 101) — ligne DEBUT du log

- Formate la date/heure courante au format `%d/%m/%Y %H:%M:%S`
  (ex. `01/10/2026 23:40:09`).
- Écrit dans le fichier de log une ligne encadrée par `SEP_LOG`
  (`SEP_LOG = "*" * 55`, soit 55 astérisques) indiquant
  `DEBUT création base hors-ligne - Envoi par mail...`.
- L'écriture passe par le petit utilitaire `_writelog` (voir §4.9).

### 4.3 `identifier_archive(rep_fic, date_fic)` (ligne 165) — trouve le `.7z`

- Détermine la **date cible** :
  - si `date_fic` vaut `today` → `datetime.date.today().strftime("%Y%m%d")`
    (ex. `20261001`) ;
  - sinon → on utilise `date_fic` telle quelle (ex. `20260925`).
- Construit le motif de recherche :
  `os.path.join(rep_fic, f"CM_{date_cible}_*.7z")`
  (ex. `D:\...\fichiers_a_envoyer\CM_20261001_*.7z`).
- **`glob.glob(motif)`** : la bibliothèque `glob` renvoie la liste de tous les
  fichiers correspondant au motif (le `*` = n'importe quelle suite de
  caractères).
- Si la liste est vide → lève `FileNotFoundError` (« Aucune archive .7z trouvée »).
- S'il y en a **plusieurs**, choisit le **plus récent** avec
  `max(candidats, key=os.path.getmtime)` (l'`os.path.getmtime` donne la date de
  modification d'un fichier).
- **Retourne** le chemin **absolu** du fichier choisi (`os.path.abspath`).

### 4.4 `lire_destinataires(base_users, champ_bool)` (ligne 214) — lecture des mails

Lit les adresses mail des destinataires dans la table `TbUsers` de la base
Access `BaseUsers`. Le paramètre `champ_bool` choisit la colonne :
- `"EnvoiMail_TT_To_CM"` → liste **TO** ;
- `"EnvoiMail_TT_Cc_CM"` → liste **CC**.

Fonctionnement :

1. **`_trouver_pilote_access()`** (ligne 201) : cherche parmi les pilotes ODBC
   installés (`pyodbc.drivers()`) celui nommé
   `Microsoft Access Driver (*.mdb, *.accdb)` (plus d'autres variantes en
   secours). S'il n'y en a pas → `RuntimeError`.
2. Ouvre une connexion **`pyodbc.connect(...)`** avec `DRIVER={pilote};DBQ=base`.
   `DBQ` = chemin de la base.
3. Exécute la requête **SQL** (calquée sur le VBA d'origine) :
   ```sql
   SELECT TbUsers.mail FROM TbUsers
   WHERE ((TbUsers.actif_CM)=True) AND ((TbUsers.<champ_bool>)=True)
   ```
   Autrement dit : on ne prend que les utilisateurs **actifs CM** ET qui ont le
   champ TO (ou CC) coché.
4. Parcourt chaque ligne, nettoie la valeur (`(row[0] or "").strip()`) et
   ajoute l'adresse à la liste si elle n'est **pas vide et pas déjà présente**
   (donc sans doublons).
5. Ferme la connexion dans un bloc `finally` (toujours fermée, même en cas
   d'erreur).
6. **Retourne** la liste (éventuellement vide) d'adresses mail.

> **Notion Python** : le bloc `try / finally` garantit que `cnx.close()` est
> exécuté quoi qu'il arrive.


### 4.5 `generer_html(base_dir, date_ficjour)` (ligne 248) — corps HTML du mail

Reconstruit, dans le langage Python, le corps **HTML** que la macro VBA générait
(pièce jointe `text.htm` envoyée avec l'archive) :

- Formate la date du jour au format `%d/%m/%Y`.
- Construit le chemin de sortie : `os.path.join(base_dir, "text.htm")`.
- Assemble une chaîne de caractères `html` :
  - un en-tête avec `content="Novell Groupwise Client..."` (reprise du VBA) et
    `charset=utf-8` ;
  - un `<body>` contenant les messages d'accompagnement (base « hors ligne »,
    consignes d'utilisation, lien de copie des fichiers Excel défini par la
    constante `LIEN_COPIE_EXCEL`), terminé par `Frédéric Rannou`.
- Écrit le contenu dans `text.htm` avec l'**encodage UTF-8**.
- **Retourne** le chemin du fichier généré.

> **Notion Python** : `open(..., "w", encoding="utf-8")` crée/écrase le fichier
> en mode écriture texte. Le `with` ferme automatiquement le fichier à la fin du
> bloc, même en cas d'erreur.

### 4.6 `GroupWiseCommander` (classe, ligne 295) — pilotage de GroupWise

Encapsule l'objet **OLE COM** `GroupwiseCommander` (le remplaçant moderne de
`NovellGroupWareSession`).

#### `__init__` (ligne 298)

- Instancie l'objet COM : `win32com.client.Dispatch("GroupwiseCommander")`.
- En cas d'échec, lève une `RuntimeError` expliquant qu'il faut GroupWise 32
  bits enregistré **et** un exécuteur **Python 32 bits (x86)**.

> **Pourquoi 32 bits ?** L'objet OLE `GroupwiseCommander` n'est disponible que
> dans un environnement 32 bits : d'où l'obligation d'utiliser Python x86.

#### `executer(commande)` (ligne 308) — exécute un token GroupWise

- Appelle `self.commander.Execute(commande, ret_string)`.
- `Execute` renvoie `1` en cas de succès, `0` en cas d'échec. Le résultat peut
  être un simple entier **ou** un tuple `(valeur, chaîne de retour)` selon le
  type de retour COM — le code gère les deux cas.
- Si le renvoi est `0`/faux → `RuntimeError` avec la commande et le détail.
- **Retourne** la chaîne de retour en cas de succès.

#### `envoyer_mail(to_liste, cc_liste, sujet, corps, pieces_jointes)` (ligne 332)

Séquence de **tokens GroupWise** exécutés dans l'ordre :

1. `NewMail()` — crée un nouveau brouillon de mail.
2. `TextSetTo("<mails séparés par des virgules>")` — remplit le champ **TO**.
3. `TextSetCC("<mails séparés par des virgules>")` — remplit le champ **CC**.
4. `TextSetSubject("<sujet>")` — définit l'**objet** du mail.
5. `TextSetMessage("<corps>")` — définit le **corps** du mail (ici simplement
   `"Frédéric Rannou"` ; le contenu détaillé est dans `text.htm` joint).
6. Pour chaque pièce jointe (`text.htm` puis l'archive `.7z`) :
   - vérifie que le fichier existe (`os.path.isfile`) sinon `FileNotFoundError` ;
   - exécute `AttachmentAdd("<chemin>")`.
7. `ItemSend()` — **envoie** le message.

> Les valeurs insérées dans les commandes sont **échappées** (les guillemets
> `"` sont remplacés par `\"`) pour éviter de casser la syntaxe du token.

### 4.7 `log_fin(fic_log, ddebut)` (ligne 111) — ligne FIN + durée

- Formate la date/heure courante.
- Calcule la **durée** : `datetime.datetime.now() - ddebut` (un objet
  `timedelta`).
- Convertit la durée en `HH:MM:SS` avec `divmod` (division entière pour les
  heures, puis minutes, puis secondes).
- Écrit la ligne `FIN... (HH:MM:SS)` dans le log.

### 4.8 `_writelog(fic_log, texte)` (ligne 89) — écriture dans le log

- Avec `locale.getpreferredencoding(False)`, détermine l'encodage de l'OS
  (ANSI/Windows-1252 sous Windows), identique à celui utilisé par Access afin de
  rester **lisible par le même fichier de log commun**.
- Ouvre le fichier de log en **mode append** (`"a"`) et écrit le texte suivi d'un
  retour à la ligne (`os.linesep`).

> **Notion Python** : écrire en mode `"a"` **ajoute** à la fin du fichier sans
> écraser son contenu existant — indispensable pour un fichier de log partagé.



---

## 5. Mode test (dry-run)

Si la variable `ModeTest` du `.env` vaut une valeur « vrai » (`1`, `true`,
`oui`, `yes`, `dryrun`, `test`…), `main()` :

1. Affiche à l'écran : l'archive `.7z` identifiée, le `text.htm` généré, les
   destinataires TO/CC, l'objet et les pièces jointes.
2. **N'exécute aucune commande GroupWise** (pas d'envoi).
3. Supprime le `text.htm` temporaire.
4. Écrit la ligne `FIN` au log et retourne `0`.

Cela permet de **valider la configuration** (`.env`, archive, destinataires,
HTML) sans risquer un envoi réel.

---

## 6. Schéma récapitulatif des appels

```
envoi_base_cm.py
│
├─ main()
│    ├─ charger_config()
│    │     └─ load_dotenv() , os.getenv(...)
│    ├─ log_debut() ──> _writelog()
│    ├─ identifier_archive()       (glob, datetime)
│    ├─ lire_destinataires("...To_CM")  ──> _trouver_pilote_access() , pyodbc
│    ├─ lire_destinataires("...Cc_CM")  ──> _trouver_pilote_access() , pyodbc
│    ├─ generer_html()                      (écrit text.htm)
│    ├─ [ModeTest ?] → print + return 0
│    ├─ GroupWiseCommander()
│    │     └─ envoyer_mail()
│    │           ├─ executer("NewMail()")
│    │           ├─ executer("TextSetTo...")
│    │           ├─ executer("TextSetCC...")
│    │           ├─ executer("TextSetSubject...")
│    │           ├─ executer("TextSetMessage...")
│    │           ├─ executer("AttachmentAdd...") × 2
│    │           └─ executer("ItemSend()")
│    ├─ os.remove(text.htm)
│    └─ log_fin() ──> _writelog()
│
└─ return 0 (ou 1 si exception)
```

---

## 7. Points techniques à retenir (pour un débutant Python)

| Notion | Explication courte | Exemple dans le code |
|--------|--------------------|----------------------|
| `if __name__ == "__main__"` | Exécution seulement en lancement direct | ligne 463 |
| `try / except` | Gestion des erreurs | `main()`, `GroupWiseCommander` |
| `try / finally` | Toujours fermer (connexion) | `lire_destinataires` |
| `with open(...)` | Ouverture auto-fermée d'un fichier | `_writelog`, `generer_html` |
| mode `"a"` (append) | Ajout en fin de fichier | `_writelog` |
| `glob.glob` | Recherche par motif `*` | `identifier_archive` |
| `max(..., key=...)` | Élément max selon une clé | `identifier_archive` |
| `divmod` | Division entière + reste | `log_fin` |
| `os.path.*` | Manipulation des chemins | un peu partout |
| `pyodbc` | Connexion aux bases Access (ODBC) | `lire_destinataires` |
| `win32com.client.Dispatch` | Instanciation d'un objet COM / OLE | `GroupWiseCommander` |
| `python-dotenv` | Charger le `.env` | `charger_config` |
| f-strings `f"..."` | Insertion de variables dans une chaîne | partout |

---

## 8. Dépendances externes

Déclarées dans `requirements.txt` :

| Module | Rôle |
|--------|------|
| `python-dotenv` | Lecture du fichier de configuration `.env` |
| `pyodbc` | Accès ODBC à la base Access (lecture de `TbUsers`) |
| `pywin32` | Fournit `win32com.client` pour piloter l'objet COM `GroupwiseCommander` |

**Prérequis machine** :
- Python **32 bits (x86)** (obligatoire pour `GroupwiseCommander`) ;
- client **GroupWise 32 bits** installé et enregistré ;
- pilote **ODBC Microsoft Access** installé.
