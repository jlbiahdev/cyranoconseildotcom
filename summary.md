---
ai_project: null
status: en cours
---

## Attentes
- Site vitrine de Cyrano Conseil (cyrano-conseil.com), hébergé chez Hostinger.
- Formulaire de contact envoyé par Hostinger (`src/envoi.php`) vers admin@cyrano-conseil.com, en remplacement de SmtpJS (en PR).

## Blocages
- Aucun

## Prochaines étapes
- Révoquer le jeton SmtpJS (SecureToken) publié dans l'ancienne version de la page : il reste visible dans l'historique du dépôt public.
- Créer les secrets GitHub du déploiement (`FTP_SERVER` = nom d'hôte du serveur Hostinger, `FTP_USERNAME`, `FTP_PASSWORD`) et, si besoin, la variable `FTP_SERVER_DIR`, puis merger la PR.
- Après le premier déploiement : envoyer un message de test et vérifier sa réception sur admin@cyrano-conseil.com (indésirables compris).
- Formulaire de contact inutilisable sur smartphone (marges fixes de 15rem, largeur nulle à 412 px) : correction à décider.
