# Contexte & Directives : EnvoiBaseCMMailPy

## 1. Vue d'ensemble du projet
Ce script Python prend le relais d'une macro Access pour l'envoi automatisé d'une archive `.7z` (base de données compactée) à une liste de diffusion.


### Contexte :
j'utilisais jusqu’à présent une base access qui :

1- Est lancée a 23h30 tous les jours 
2- Génère une copie d'une base Access
3- Lance des traitements Access (copie des tables liées dans le base locale)
4- Compresse la base avec 7zip
5- Envoi cette archive 7Zip a une liste de diffusion (stockée dans une table Access)

Cette dernière étape se faisait avec Access qui pilotait groupwise a travers l'objet NovellGroupWareSession, que j'instanciais de la manière suivante :

```vba
Set objGroupwise = CreateObject("NovellGroupWareSession")
```
Malheureusement suite a un passage a windows 11, ma DSI a également upgradé groupwise, je n'ai donc plus accès a ce NovellGroupWareSession.

J'ai par contre la possibilité d'utiliser l'objet GroupwiseCommander (version VBA)

```vba
Set GWCom = CreateObject("GroupwiseCommander")
```
Les étapes de 1 à 4 seront toujours fait avec Access.

Mais il faudrait redévelopper le 5 en Python.

### Flux de traitement :
1. **Invocateur** : Appelé à la fin des traitements quotidiens Access (vers 23h30).
Le script Python sera lancé a travers un planificateur de taches, 15 ou 20 minutes apres le lancement du traitement Access.
2. **Action** :
   - Récupère le chemin de l'archive 7Zip générée par Access.
    Le début du chemin doit etre paramétrable dans un fichier .env
    RepFicAenvoyer = "D:\Users\frannou\Dev\Access\CM\CM_Import_TT\fichiers_a_envoyer\"
    Le nom du fichier 7Zip a envoyer est construit comme cela : CM_AAAAMMJJ_HHMMSS.7z
    Le fichier est horodaté avec sa date de génération dans le traitement Access initial (ex. CM_20261001_233513.7z)
    Le traitement python doit donc pouvoir identifier le fichier à envoyer selon un paramètre disponible dans le fichier .env :
    Si DateFicAenvoyer = today => le script identifiera le fichier a la date du jour.
    Si DateFicAenvoyer contient une date donnée (ex. DateFicAenvoyer = 20260925), le fichier qui sera envoyé contiendra cette date (ex. CM_20260925_233514.7z)
   - Lit la liste des destinataires stockée dans une base Access.
   L'emplacement de cette base sera stockée dans le fichier .env.
   ex. BaseUsers = "G:\DM-SURVEIL\MEDICAMENTS\BASES\BASES_DIVERSES\Users\Users.accdb"
    Ensuite on obtient deux liste de mails (champs TbUsers.mail), une pour remplir le champ "TO" et l'autre pour le champ "CC" des mails à envoyer :

```vba
    sSQLTo = "SELECT TbUsers.actif_CM, TbUsers.EnvoiMail_TT_To_CM, TbUsers.idUser, TbUsers.UserName, TbUsers.nom, TbUsers.prenom, TbUsers.mail, TbUsers.metier, TbUsers.direction, TbUsers.fonction " & _
            "FROM TbUsers " & _
            "WHERE (((TbUsers.actif_CM)=True) AND ((TbUsers.EnvoiMail_TT_To_CM)=True));"
    sSQLCc = "SELECT TbUsers.actif_CM, TbUsers.EnvoiMail_TT_To_CM, TbUsers.idUser, TbUsers.UserName, TbUsers.nom, TbUsers.prenom, TbUsers.mail, TbUsers.metier, TbUsers.direction, TbUsers.fonction " & _
            "FROM TbUsers " & _
            "WHERE (((TbUsers.actif_CM)=True) AND ((TbUsers.EnvoiMail_TT_Cc_CM)=True));"
```


   - Instancie l'objet OLE **`GroupwiseCommander`** (`win32com.client.Dispatch("GroupwiseCommander")`).
   - Exécute les commandes GroupWise (ex: `ItemSendSecret...` ou syntaxe token GroupWise) pour composer le mail, joindre le fichier `.7z` et envoyer le message.
   - Le code VBA/Access avec l'ancienne méthode "Set objGroupwise = CreateObject("NovellGroupWareSession")" pour que tu puisses t'en inspirer afin de la coder en python qui utilisera l'objet "Set GWCom = CreateObject("GroupwiseCommander")" se situe a la fin de ce fichier, dans la section Annexes


3. **Sortie** : Code de retour `0` si succès, `1` en cas d'erreur.
    - Un fichier de log, commun avec le traitement access lancé en premier, aura son emplacement paramétrable dans le fichier .env.
    ex. FicLog = "G:\DM-SURVEIL\MEDICAMENTS\BASES\BASE_CM\Import_TT\TT_Import.log"

    - Typiquement les logs, en mode append dans ce fichier, seront pour ce traitement Python :

```log
*************************************************** DEBUT création base hors-ligne - Envoi par mail de la base hors-ligne - 01/10/2026 23:40:09 ***************************************************
***************************************************  FIN  création base hors-ligne - Envoi par mail de la base hors-ligne - 01/10/2026 23:40:13 (00:00:04) ***************************************************
```

---

## 2. Contraintes d'architecture

> **EXÉCUTION 32 BITS OBLIGATOIRE** :
> L'objet OLE `GroupwiseCommander` nécessite un environnement d'exécution **Python 32 bits (x86)**.

- **Language** : Python 3.12+ (x86 / 32 bits).
- **Bibliothèque OLE** : `pywin32`.
- **Composant GroupWise** : `GroupwiseCommander`.

### Activation (Git Bash)
```bash
source .venv/Scripts/activate
```

---

## 3. Contraintes

- Toujours me répondre en français

---

## 4. Annexes

### Aide Groupwise sur l'objet GroupwiseCommander
D:\Users\Frannou\Dev\Access\EnvoiMail_Test_Windows_11\pour_Nico_bertrand\gwsdk_token.pdf
ou en format Word : D:\Users\Frannou\Dev\Python\EnvoiBaseCMMailPy\Doc\gwsdk_token.docx
ou en HTML : D:\Users\Frannou\Dev\Python\EnvoiBaseCMMailPy\Doc\gwsdk_token.html

### Code Access avec l'objet NovellGroupWareSession

Voici le code VBA/Access avec l'ancienne méthode "Set objGroupwise = CreateObject("NovellGroupWareSession")" pour que tu puisses t'en inspirer a la coder en python qui utilisera l'objet "Set GWCom = CreateObject("GroupwiseCommander")" :

```vba
Public Function EnvoiMailBaseHL()
Dim sPathZipProgram As String

Dim sFolderZipFile As String
Dim sNameZipFile As String
Dim sNameBaseHL As String
Dim sFolderToZip As String
Dim sSQLTo As String
Dim sSQLCc As String
Dim sNow As String
Dim ShellStr As String
Dim db As Database
Dim RstTT_CreationBaseHorsLigne As Recordset
Dim RstListeMail As Recordset
Dim sFicLog As String
Dim sRepBase As String
Dim sPathTextHtm_Stream As String
Dim dDebut As Date
Dim dFin As Date
Dim sCorpsHTML As String
Dim sPathModeOpe As String
Dim sAdressTempo As String
Dim objStream As Object

Dim sBodyText As String
Dim sDateDuJour As String
Dim objGroupwise As Object, objAccount As Object, objProxyAccount As Object, objMailbox As Object
Dim objMessage As Object, objMessages As Object, objRecipients As Object, objAttachments As Object
Dim objRecipient As Object, MessageSent As Variant
'Déclaration des objets
Set objGroupwise = CreateObject("NovellGroupWareSession")
'Set objAccount = objGroupwise.Login
Set objAccount = objGroupwise.Login("Frannou", , "couack")

'If Nz(strAdresseMandataire) <> "" Then
'    Set objProxyAccount = objGroupwise.Proxy(strAdresseMandataire)
'    Set objMailbox = objProxyAccount.Mailbox
'Else
    Set objMailbox = objAccount.Mailbox
'End If

sSQLTo = "SELECT TbUsers.actif_CM, TbUsers.EnvoiMail_TT_To_CM, TbUsers.idUser, TbUsers.UserName, TbUsers.nom, TbUsers.prenom, TbUsers.mail, TbUsers.metier, TbUsers.direction, TbUsers.fonction " & _
         "FROM TbUsers " & _
         "WHERE (((TbUsers.actif_CM)=True) AND ((TbUsers.EnvoiMail_TT_To_CM)=True));"
sSQLCc = "SELECT TbUsers.actif_CM, TbUsers.EnvoiMail_TT_To_CM, TbUsers.idUser, TbUsers.UserName, TbUsers.nom, TbUsers.prenom, TbUsers.mail, TbUsers.metier, TbUsers.direction, TbUsers.fonction " & _
         "FROM TbUsers " & _
         "WHERE (((TbUsers.actif_CM)=True) AND ((TbUsers.EnvoiMail_TT_Cc_CM)=True));"
sRepBase = CurrentProject.Path & "\"
Set db = CurrentDb
Set RstTT_CreationBaseHorsLigne = db.OpenRecordset("TT_CreationBaseHorsLigne", dbOpenDynaset)
dDebut = Now()
DoCmd.Hourglass (True)
If DonneParam("ENV_importTT", True) = "PROD" Then
    sFicLog = "TT_Import.log"
ElseIf DonneParam("ENV_importTT", True) = "DEV" Then
    sFicLog = "TT_Import_DEV.log"
Else
    sFicLog = "TT_Import.log"
End If

Close #1
Open sRepBase & sFicLog For Append As #1
Print #1, "*************************************************** DEBUT création base hors-ligne - Envoi par mail de la base hors-ligne - " & Now() & " ***************************************************"
Close #1
sDateDuJour = Format(Day(Now), "00") & "/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
RstTT_CreationBaseHorsLigne.MoveFirst

sFolderZipFile = RstTT_CreationBaseHorsLigne![RepBaseExpDuJour]
sNameZipFile = "CM_" & RstTT_CreationBaseHorsLigne![UnicBaseExpDuJour] & ".7z"

sPathModeOpe = RstTT_CreationBaseHorsLigne![PathModeOpe]
' sFolderZipFile & "\" & sNameZipFile

Set objMessages = objMailbox.Messages


'Génération du mail
Set objMessage = objMessages.Add("GW.MESSAGE.MAIL", "Draft")

Set objRecipients = objMessage.Recipients

Set RstListeMail = db.OpenRecordset(sSQLTo, dbOpenDynaset)
If RstListeMail.NoMatch = False Then
    Do Until RstListeMail.EOF
        'To
'        Set objRecipient = objRecipients.Add("Frederic.RANNOU@ansm.sante.fr", , 0)
        'Set objRecipient = objRecipients.Add("Frederic.RANNOU@ansm.sante.fr")
        sAdressTempo = RstListeMail![mail]
'        Set objRecipient = objRecipients.Add(sAdressTempo, , 0)
        Set objRecipient = objRecipients.Add(sAdressTempo, , 0)
        RstListeMail.MoveNext
    Loop
Else
    Close #1
    Open sRepBase & sFicLog For Append As #1
    Print #1, "pas de destinataires TO "
    Close #1
End If
RstListeMail.Close

Set RstListeMail = db.OpenRecordset(sSQLCc, dbOpenDynaset)
If RstListeMail.NoMatch = False Then
    Do Until RstListeMail.EOF
        'Cc
        'Set objRecipient = objRecipients.Add("fredoh@gmail.com", , 1)
        'Set objRecipient = objRecipients.Add("fredoh@gmail.com")

        sAdressTempo = RstListeMail![mail]
'        Set objRecipient = objRecipients.Add(sAdressTempo, , 1)
        Set objRecipient = objRecipients.Add(sAdressTempo, , 1)
        RstListeMail.MoveNext
    Loop
Else
    Close #1
    Open sRepBase & sFicLog For Append As #1
    Print #1, "pas de destinataires CC "
    Close #1
End If
RstListeMail.Close


'''''''''''''''''''''''''''''''''''''''''
sPathTextHtm_Stream = CurrentProject.Path & "\" & "text.htm"
'Create the stream
Set objStream = CreateObject("ADODB.Stream")
'Initialize the stream
objStream.Open
'Reset the position and indicate the charactor encoding
objStream.Position = 0
objStream.Charset = "UTF-8"
'Write to the steam
objStream.WriteText "<html><head>" & vbCrLf
objStream.WriteText "" & vbCrLf
objStream.WriteText "<meta name=" & Chr(34) & "Generator" & Chr(34) & " content=" & Chr(34) & "Novell Groupwise Client (Version 14.2.1  Build: 124595)" & Chr(34) & ">" & vbCrLf
objStream.WriteText "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=utf-8" & Chr(34) & "></head>" & vbCrLf
objStream.WriteText "<body>" & vbCrLf

'objStream.WriteText "<div>base cas marquants télétravail<br></div>" & vbCrLf
objStream.WriteText "<div>Voici la base cas marquant " & Chr(34) & "hors ligne" & Chr(34) & " du " & sDateDuJour & " au soir ainsi que son mode opératoire.</div><div><br></div>" & vbCrLf
objStream.WriteText "<div>Si vous désirez modifier un cas marquant, vous devez impérativement, utiliser la dernière base envoyée, il s'agit de la plus à jour.</div><div><br></div>" & vbCrLf
objStream.WriteText "<div>Il est également nécessaire d’envoyer les fichiers générés le jour même, et ainsi réduire le risque de collisions de modification (modifications du même cas marquant par deux utilisateurs différents).</div><div><br></div>" & vbCrLf
'objStream.WriteText "<div>Merci de copier les fichiers Excel dans le répertoire suivant : <a href=" & Chr(34) & "G:\DM-SURV\VIGIL1\COMMUNDP\PROCESSUS CM\BaseDev\ImportFromSM_CV\ImportExcel_TT\PROD\aFaire\" & Chr(34) & ">G:\DM-SURV\VIGIL1\COMMUNDP\PROCESSUS CM\BaseDev\ImportFromSM_CV\ImportExcel_TT\PROD\aFaire\</a></div><div><br></div>" & vbCrLf
objStream.WriteText "<div>Merci de copier les fichiers Excel dans le répertoire suivant : <a href=" & Chr(34) & "G:\DM-SURVEIL\MEDICAMENTS\BASES\BASE_CM\Import_TT\ImportExcel_TT\PROD\aFaire\" & Chr(34) & ">G:\DM-SURVEIL\MEDICAMENTS\BASES\BASE_CM\Import_TT\ImportExcel_TT\PROD\aFaire\</a></div><div><br></div>" & vbCrLf
objStream.WriteText "<div>Frédéric Rannou<br></div></div>" & vbCrLf
objStream.WriteText "</body></html>" & vbCrLf


'Save the stream to a file
objStream.SaveToFile sPathTextHtm_Stream, 2

Set objAttachments = objMessage.Attachments.Add(sPathTextHtm_Stream)
'''''''''''''''''''''''''''''''''''''''''

Set objAttachments = objMessage.Attachments.Add(sFolderZipFile & "\" & sNameZipFile)
Set objAttachments = objMessage.Attachments.Add(sPathModeOpe)


objMessage.Subject = "Base Cas marquants hors ligne pour saisie par DP et DS/Sécurisation en période de télétravail"


' Ajout du corps du mail
sBodyText = "Voici la base cas marquant " & Chr(34) & "hors ligne" & Chr(34) & " du " & sDateDuJour & " au soir ainsi que son mode opératoire." & vbTab
sBodyText = sBodyText & "Si vous désirez modifier un cas marquant, vous devez impérativement, utiliser la dernière base envoyée, il s'agit de la plus à jour." & vbTab
sBodyText = sBodyText & "Il est également nécessaire d’envoyer les fichiers générés le jour même, et ainsi réduire le risque de collisions de modification (modifications du même cas marquant par deux utilisateurs différents)." & vbTab
'sBodyText = sBodyText & "Merci de copier les fichiers Excel dans le répertoire suivant : G:\DM-SURV\VIGIL1\COMMUNDP\PROCESSUS CM\BaseDev\ImportFromSM_CV\ImportExcel_TT\PROD\aFaire\" & vbTab
sBodyText = sBodyText & "Merci de copier les fichiers Excel dans le répertoire suivant : G:\DM-SURVEIL\MEDICAMENTS\BASES\BASE_CM\Import_TT\ImportExcel_TT\PROD\aFaire\" & vbTab
sBodyText = "Frédéric Rannou" & vbTab




'Envoi du mail

objMessage.Send
dFin = Now()

Close #1
Open sRepBase & sFicLog For Append As #1
Print #1, "***************************************************  FIN  création base hors-ligne - Envoi par mail de la base hors-ligne - " & Now() & " (" & CDate(dFin - dDebut) & ")" & " ***************************************************"
Close #1

Set objAccount = Nothing
Set objProxyAccount = Nothing
Set objGroupwise = Nothing

DoCmd.Hourglass (False)

End Function
```


## 3. Commandes Utiles

### Activation (Git Bash)

```bash
source .venv/Scripts/activate
```