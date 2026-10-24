# -*- coding: utf-8 -*-
"""Diagnostic Option 1 (SANS ENVOI, pas d'ItemSend) :
ouvre un brouillon GroupWise et y colle le corps avec de VRAIS sauts de ligne
via le presse-papier Windows (FocusSet(10) + EditPaste())."""
import win32clipboard
from envoi_base_cm import GroupWiseCommander

corps_avec_sauts = (
    "Voici la base cas marquant \"hors ligne\" du 02/10/2026 au soir "
    "ainsi que son mode opératoire.\r\n\r\n"
    "Si vous désirez modifier un cas marquant, vous devez impérativement, "
    "utiliser la dernière base envoyée, il s'agit de la plus à jour.\r\n\r\n"
    "Il est également nécessaire d’envoyer les fichiers générés le jour même, "
    "et ainsi réduire le risque de collisions de modification "
    "(modifications du même cas marquant par deux utilisateurs différents).\r\n\r\n"
    "Merci de copier les fichiers Excel dans le répertoire suivant : "
    "G:\\DM-SURVEIL\\MEDICAMENTS\\BASES\\BASE_CM\\Import_TT\\ImportExcel_TT\\PROD\\aFaire\\\r\n\r\n"
    "Frédéric Rannou"
)

cm = GroupWiseCommander()
cm.executer("NewMail()")
cm.executer('TextSetSubject("TEST OPTION 1 - VRAIS SAUTS DE LIGNE")')

# Copie dans le presse-papier Windows
win32clipboard.OpenClipboard()
try:
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, corps_avec_sauts)
finally:
    win32clipboard.CloseClipboard()

# Focus sur le champ Message (10 = Message) puis EditPaste()
cm.executer("FocusSet(10)")
cm.executer("EditPaste()")

print("Brouillon ouvert dans GroupWise avec les vrais sauts de ligne (sans envoi).")
