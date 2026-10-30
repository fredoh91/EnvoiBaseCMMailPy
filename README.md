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
3. Génère le corps HTML détaillé dans un **fichier temporaire `text.htm`**
   (reproduit du HTML du VBA/Access), qui sera **joint** au courriel.
4. Compose et **envoie** le courriel via l'objet OLE `GroupwiseCommander` : le
   corps du message est un **texte simple** (`Frédéric Rannou`), le HTML détaillé
   (lien cliquable compris) étant la pièce jointe `text.htm`.
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
| `FicLogTrtComplet`| *(optionnel)* Chemin d'un second fichier de log (log global du traitement) → les lignes DEBUT/infos/FIN y sont aussi écrites |
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

> **Destinataires multiples dans GroupWise** : chacun est ajouté au brouillon
> par **un token par adresse** (`TextSetTo("<mail1>")` puis `TextSetTo("<mail2>"; True)`
> pour les suivants, idem avec `TextSetCC`), via le paramètre `Append=True`.
> La concaténation en une seule commande `TextSetTo("mail1,mail2")` produirait une
> adresse unique invalide (champ destinataire incohérent, envoi impossible).

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
- **Corps** : texte simple `Frédéric Rannou` (le signataire). Le HTML détaillé
  n'est **pas** inséré dans le corps du mail.
- **Pièces jointes** : l'archive `CM_AAAAMMJJ_HHMMSS.7z` **plus** le fichier
  temporaire `text.htm` (supprimé après l'envoi).
- Le fichier `text.htm` contient le **HTML complet** (base « hors ligne »,
  consignes, **lien cliquable** de copie des fichiers Excel) reproduit du VBA
  d'origine.

> Ce choix (corps simple + HTML en pièce jointe `text.htm`) reproduit exactement
> le comportement du VBA/Access d'origine et contourne la **limite du token
> `TextSetMessage`** de GroupWise qui **rejette les balises `<a>`** dans le corps
> (démontré par un diagnostic progressif : `TextSetMessage` accepte `<div>`,
> `<br>`, accents et guillemets échappés, mais **échoue dès qu'on ajoute un lien**).

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
