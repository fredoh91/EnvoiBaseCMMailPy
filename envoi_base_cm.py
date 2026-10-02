import os
import sys
import win32com.client


class GroupWiseMailSender:
    """Automatise l'envoi de courriels via le client Novell / OpenText GroupWise

    en utilisant l'interface COM GroupWiseCommander et les jetons (Tokens).
    """

    def __init__(self):
        try:
            # Initialisation de l'objet COM GroupWise Commander
            self.commander = win32com.client.Dispatch("GroupwiseCommander")
        except Exception as e:
            print(
                f"[ERREUR] Échec de l'initialisation de l'objet COM GroupWiseCommander : {e}"
            )
            print(
                "Vérifiez que le client GroupWise est correctement installé et enregistré sur cette machine."
            )
            sys.exit(1)

    def send_token(self, command: str) -> str:
        """Exécute une commande de jeton (Token) GroupWise via l'API Commander.

        :param command: La chaîne du jeton à exécuter (ex. 'NewMail()')
        :return: La chaîne de résultat ou d'erreur renvoyée par GroupWise Commander
        """
        # Initialisation du paramètre de chaîne de sortie pour éviter les incohérences de type COM
        ret_string = ""

        # Appel de Execute(CMD, RetString)
        # Renvoie 1 en cas de succès, 0 en cas d'échec
        retval, ret_string = self.commander.Execute(command, ret_string)

        if retval == 0:
            raise RuntimeError(
                f"Échec de l'exécution du jeton GroupWise pour '{command}'. "
                f"Détail de l'erreur : {ret_string}"
            )

        return ret_string

    def send_email(
        self,
        to_recipients: str,
        subject: str,
        body: str,
        cc_recipients: str = "",
        bc_recipients: str = "",
        attachments: list = None,
    ):
        """Crée et envoie un courriel via les jetons GroupWise.

        :param to_recipients: Adresses ou noms des destinataires principaux (séparés par des virgules ou points-virgules)
        :param subject: Objet du message
        :param body: Corps du message
        :param cc_recipients: Adresses des destinataires en copie (optionnel)
        :param bc_recipients: Adresses des destinataires en copie cachée (optionnel)
        :param attachments: Liste des chemins d'accès complets des fichiers à joindre (optionnel)
        """
        print("Début du processus de création du courriel GroupWise...")

        # 1. Ouvrir une nouvelle vue de rédaction de courriel
        print("-> Ouverture d'un nouveau brouillon (NewMail)...")
        self.send_token("NewMail()")

        # 2. Définir les destinataires principaux (Pour)
        if to_recipients:
            print(f"-> Définition des destinataires (TO) : {to_recipients}")
            self.send_token(f'TextSetTo("{to_recipients}")')

        # 3. Définir les destinataires en copie (CC - Optionnel)
        if cc_recipients:
            print(f"-> Définition des destinataires en copie (CC) : {cc_recipients}")
            self.send_token(f'TextSetCC("{cc_recipients}")')

        # 4. Définir les destinataires en copie cachée (CCI - Optionnel)
        if bc_recipients:
            print(
                f"-> Définition des destinataires en copie cachée (BC) : {bc_recipients}"
            )
            self.send_token(f'TextSetBC("{bc_recipients}")')

        # 5. Définir l'objet
        print(f"-> Définition de l'objet : {subject}")
        # Échappement des guillemets doubles internes si présent
        safe_subject = subject.replace('"', '\\"')
        self.send_token(f'TextSetSubject("{safe_subject}")')

        # 6. Définir le corps du message
        print("-> Définition du corps du message...")
        safe_body = body.replace('"', '\\"')
        self.send_token(f'TextSetMessage("{safe_body}")')

        # 7. Ajouter les pièces jointes (Optionnel)
        # Utilise l'identifiant temporaire 'X00' qui représente la fenêtre de rédaction active non envoyée
        if attachments:
            for file_path in attachments:
                abs_path = os.path.abspath(file_path)
                if not os.path.exists(abs_path):
                    print(
                        f"[AVERTISSEMENT] Fichier joint introuvable : {abs_path}. Ignoré."
                    )
                    continue

                print(f"-> Ajout de la pièce jointe : {abs_path}")
                safe_path = abs_path.replace('"', '\\"')
                self.send_token(f'AttachmentAdd("X00", "{safe_path}")')

        # 8. Envoyer le message
        # 'X00' cible la vue de message active en cours de rédaction
        print("-> Envoi du courriel (ItemSend)...")
        self.send_token('ItemSend("X00")')

        print("Courriel envoyé / mis en file d'attente avec succès via GroupWise.")


if __name__ == "__main__":
    # Exemple d'utilisation / exécution de test du script
    try:
        sender = GroupWiseMailSender()

        # Définition des paramètres du message
        recipient_to = "destinataire@exemple.fr"
        subject_line = (
            "Test d'envoi automatique d'email via Python et GroupWise Commander"
        )
        email_body = (
            "Bonjour,\n\nCeci est un test d'envoi automatique généré via Python "
            "et le kit SDK GroupWise Commander.\n\nCordialement."
        )

        # Liste optionnelle de fichiers à joindre
        # sample_attachments = [r"C:\temp\rapport.pdf"]
        sample_attachments = []

        sender.send_email(
            to_recipients=recipient_to,
            subject=subject_line,
            body=email_body,
            attachments=sample_attachments,
        )

    except Exception as err:
        print(
            f"\n[ERREUR CRITIQUE] Une erreur est survenue lors de l'exécution du script : {err}"
        )
        sys.exit(1)