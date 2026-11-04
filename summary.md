---
ai_project: null
status: en cours
---

## Attentes
- Site vitrine de Cyrano Conseil (cyrano-conseil.com), hébergé chez Hostinger.
- Formulaire de contact envoyé par Hostinger (`src/envoi.php`) vers admin@cyrano-conseil.com, en remplacement de SmtpJS (PR #3 mergée le 4 novembre 2026, pas encore en ligne).

## Blocages
- Déploiement FTPS en échec : connexion et identification acceptées, puis le serveur coupe le canal de transfert dès la première opération (`Client is closed because read ECONNRESET (data socket)`). Même problème sur lissenpage ; cause non identifiée, à traiter dans un projet dédié.

## Prochaines étapes
- Diagnostiquer le déploiement (journal détaillé avec `log-level: verbose`), corriger, puis relancer.
- Après le premier déploiement réussi : envoyer un message de test et vérifier sa réception sur admin@cyrano-conseil.com (indésirables compris).
- Neutraliser le jeton SmtpJS encore visible dans l'historique du dépôt public : changer le mot de passe SMTP qu'il contient (expéditeur contact.cyranoc@gmail.com).
- Formulaire de contact inutilisable sur smartphone (marges fixes de 15rem, largeur nulle à 412 px) : correction à décider.
