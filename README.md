# EnvoiBaseCMMailPy

Envoi automatisé de la base CM **« hors ligne »** (archive `.7z`) par courriel **GroupWise**,
en prenant le relais de la macro Access (étape 5 de la chaîne de traitement quotidienne Access).

> **Adapté de** : `INSTRUCTIONS.md` (cahier des charges) et du code VBA/Access d'origine
> (`NovellGroupWareSession` → remplacé par `GroupwiseCommander`).

---

## 🎯 Rôle du script

Le traitement Access quotidien (lancé vers 23h30) génère une archive
`CM_AAAAMMJJ_HHMMSS.7z`. Ce script Python est appelé **15 à 20 minutes après**
le démarrage du traitement Access (via le Planificateur de tâches Windows) et :

1. Identifie l'archive `.7z` du jour dans le dossier `RepFicAenvoyer`.
2. Lit la liste des destinataires **TO** et **CC** dans la table `TbUsers` de la
   base Access (`BaseUsers`).
3. Génère le fichier HTML joint (`text.htm`) reproduisant le corps HTML du VBA.
4. Compose et **envoie** le courriel via l'objet OLE `GroupwiseCommander`.
5. Journalise le tout dans un fichier de log commun (mode **append**).

Retour : **0** si succès, **1** en cas d'erreur.

---

## ⚠️ PRÉREQUIS OBLIGATOIRE : Python 32 bits (x86)

L'objet OLE **`GroupwiseCommander`** nécessite un environnement d'exécution
**Python 32 bits (x86)**.

- **Python** : 3.12+ (x86 / 32 bits uniquement)
- Activez le venv (ex. sous Git Bash) :
  ```bash
  source .venv/Scripts/activate
  ```

### Bibliothèques requises

Installez-les avec `pip` :

```bash
pip install pywin32 python-dotenv pyodbc
```

| Bibliothèque    | Rôle                                    |
|-----------------|-----------------------------------------|
| `pywin32`       | Objet COM `GroupwiseCommander`          |
| `python-dotenv` | Lecture du fichier `.env`               |
| `pyodbc`        | Lecture de la base Access (table `TbUsers`) |

### Pilote ODBC Microsoft Access

Un **pilote ODBC Microsoft Access** doit être installé sur la machine
(ex. *Microsoft Access Database Engine*), fournissant au minimum le pilote :
`Microsoft Access Driver (*.mdb, *.accdb)`.

---

## 📝 Configuration (fichier `.env`)

Copiez le modèle puis renseignez vos valeurs :

```bash
copy .env.example .env
```

Variables :

| Variable          | Description                                                            |
|-------------------|------------------------------------------------------------------------|
| `RepFicAenvoyer`  | Dossier contenant l'archive `.7z` (chemin **terminé par un backslash**) |
| `DateFicAenvoyer` | `today` (date du jour) **ou** une date exacte `AAAAMMJJ`               |
| `BaseUsers`       | Chemin complet de la base Access (table `TbUsers`)                     |
| `FicLog`          | Chemin complet du fichier de log commun (mode append)                  |
| `ModeTest`        | *(optionnel)* `1`/`true`/`dryrun` → TEST **sans envoi** réel du mail    |

> **Important** : ne mettez **pas** de guillemets autour des valeurs.

### Fichier `.7z` identifié

- `DateFicAenvoyer=today` → recherche `CM_AAAA-MM-JJ_*.7z` à la date du jour.
- `DateFicAenvoyer=20260925` → recherche `CM_20260925_*.7z`.

S'il existe plusieurs archives correspondant à la date, la **plus récente** est choisie.

### Destinataires (table `TbUsers`)

Deux requêtes sont exécutées (logique issue du VBA) :

- **TO** : `EnvoiMail_TT_To_CM = True` **et** `actif_CM = True`
- **CC** : `EnvoiMail_TT_Cc_CM = True` **et** `actif_CM = True`

Le champ utilisé est `TbUsers.mail`.

---

## ▶️ Utilisation

```bash
python envoi_base_cm.py
```

### 🧪 Mode test (dry-run) — sans envoi réel

Pour valider la configuration, l'identification du `.7z`, la lecture des
destinataires et la génération du HTML **sans** envoyer de mail, ajoutez
`ModeTest=1` dans `.env` puis relancez :

```bash
python envoi_base_cm.py
```

Le script affiche alors les informations de test (archive trouvée, destinataires
TO/CC, objet, pièces jointes) et retourne `0` **sans exécuter** les commandes
GroupWise. Retirez `ModeTest` (ou mettez-le à vide) pour un envoi réel.

```
*************************************************** DEBUT création base hors-ligne - Envoi par mail de la base hors-ligne - 01/10/2026 23:40:09 ***************************************************
***************************************************  FIN  création base hors-ligne - Envoi par mail de la base hors-ligne - 01/10/2026 23:40:13 (00:00:04) ***************************************************
```

En l'absence de destinataires :
```
pas de destinataires TO
pas de destinataires CC
```

En cas d'erreur, une ligne `ERREUR envoi par mail de la base hors-ligne : ...`
est ajoutée au log et le script retourne **1**.

---

## 📧 Contenu du courriel envoyé

- **Objet** :
  `Base Cas marquants hors ligne pour saisie par DP et DS/Sécurisation en période de télétravail`
- **Corps** : `Frédéric Rannou` (comme dans le VBA d'origine).
- **Pièces jointes** :
  1. `text.htm` (corps HTML généré, reproduit du VBA),
  2. l'archive `CM_AAAAMMJJ_HHMMSS.7z`.

> Conformément au choix retenu, **aucun fichier mode opératoire** n'est joint.

---

## 🗂️ Structure

```
envoi_base_cm.py     Script principal
.env                 Configuration (ne pas versionner)
.env.example         Modèle de configuration
README.md            Cette documentation
```

---

## 🔧 Dépannage

| Symptôme                          | Cause probable / vérification                                        |
|-----------------------------------|----------------------------------------------------------------------|
| `GroupwiseCommander` indisponible | Vérifier Python **32 bits** + client GroupWise 32 bits enregistré    |
| Pilote Access introuvable         | Installer *Microsoft Access Database Engine* (pilote ODBC)           |
| « Aucune archive .7z trouvée »    | Vérifier `RepFicAenvoyer` (backslash final) et `DateFicAenvoyer`      |
| Variables manquantes              | Vérifier la présence des 4 variables dans `.env`                      |
