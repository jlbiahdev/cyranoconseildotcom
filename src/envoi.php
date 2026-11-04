<?php
// Réception des formulaires du site : chaque message est envoyé par e-mail à l'adresse du domaine
// hébergée chez Hostinger. Aucun service tiers, rien n'est enregistré sur le serveur.
// Même mécanisme que nzuane.org (sites/local/envoi.php).

declare(strict_types=1);

// ---------------------------------------------------------------- configuration du site
const DESTINATAIRE = 'admin@cyrano-conseil.com';
const EXPEDITEUR = 'Cyrano Conseil <admin@cyrano-conseil.com>';
const PAGE_MERCI = 'merci.html';

$formulaires = [
    'contact' => [
        'introduction' => 'Nouveau message reçu depuis le formulaire de contact de cyrano-conseil.com.',
        'sujet' => fn (array $v): string => '[cyrano-conseil.com] ' . ($v['sujet'] !== '' ? $v['sujet'] : 'Message depuis le site'),
        'champs' => [
            'nom' => ['libelle' => 'Nom', 'requis' => true],
            'sujet' => ['libelle' => 'Sujet'],
            'email' => ['libelle' => 'E-mail', 'type' => 'email'],
            'telephone' => ['libelle' => 'Téléphone'],
            'message' => ['libelle' => 'Message', 'type' => 'long', 'requis' => true],
        ],
        'joignable' => ['email', 'telephone'],
    ],
];

const PAGE_ERREUR = <<<'HTML'
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="robots" content="noindex">
  <link rel="shortcut icon" href="./favicon.ico">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous">
  <link rel="stylesheet" type="text/css" href="./public/css/index.css">
  <title>Message non envoyé - Cyrano Conseil</title>
</head>
<body>
  <main class="container my-5">
    <h1 class="h3 mb-4">Votre message n'a pas pu être envoyé.</h1>
    <p>Vérifiez que votre nom, votre message et un moyen de vous répondre (e-mail ou téléphone) sont bien renseignés, puis réessayez.</p>
    <p>Si le problème persiste, écrivez-nous directement à <a href="mailto:admin@cyrano-conseil.com">admin@cyrano-conseil.com</a>.</p>
    <p class="mt-4"><a class="btn btn-full" href="index.html#contactus">retour au formulaire</a></p>
  </main>
</body>
</html>
HTML;
// ---------------------------------------------------------------- fin de la configuration

function valeur(string $nom, int $longueurMax): string
{
    $valeur = $_POST[$nom] ?? '';
    if (!is_string($valeur)) {
        return '';
    }
    return mb_substr(trim($valeur), 0, $longueurMax);
}

// Les retours à la ligne sont retirés des champs d'une ligne, pour empêcher l'injection d'en-têtes.
function surUneLigne(string $valeur): string
{
    return preg_replace('/[\r\n]+/', ' ', $valeur) ?? '';
}

function pageErreur(int $code): never
{
    http_response_code($code);
    header('Content-Type: text/html; charset=UTF-8');
    echo PAGE_ERREUR;
    exit;
}

if (($_SERVER['REQUEST_METHOD'] ?? '') !== 'POST') {
    header('Allow: POST');
    pageErreur(405);
}

// Robot : le champ caché est rempli. On fait comme si tout allait bien, sans rien envoyer.
if (valeur('site_web', 200) !== '') {
    header('Location: ' . PAGE_MERCI, true, 303);
    exit;
}

$nomFormulaire = valeur('formulaire', 50);
if (!isset($formulaires[$nomFormulaire])) {
    pageErreur(400);
}
$formulaire = $formulaires[$nomFormulaire];

$valeurs = [];
foreach ($formulaire['champs'] as $nom => $champ) {
    $type = $champ['type'] ?? 'texte';
    $longueurMax = $type === 'long' ? 5000 : 200;
    $v = valeur($nom, $longueurMax);
    if ($type !== 'long') {
        $v = surUneLigne($v);
    }
    if ($v === '' && !empty($champ['requis'])) {
        pageErreur(400);
    }
    if ($v !== '' && $type === 'email' && filter_var($v, FILTER_VALIDATE_EMAIL) === false) {
        pageErreur(400);
    }
    if ($v !== '' && $type === 'choix' && !in_array($v, $champ['choix'], true)) {
        pageErreur(400);
    }
    $valeurs[$nom] = $v;
}

// Au moins un moyen de répondre (e-mail ou téléphone), quand le formulaire l'exige.
if (!empty($formulaire['joignable'])) {
    $joignable = false;
    foreach ($formulaire['joignable'] as $nom) {
        if (($valeurs[$nom] ?? '') !== '') {
            $joignable = true;
        }
    }
    if (!$joignable) {
        pageErreur(400);
    }
}

$sujet = surUneLigne(($formulaire['sujet'])($valeurs));

$lignes = [$formulaire['introduction'], ''];
foreach ($formulaire['champs'] as $nom => $champ) {
    $v = $valeurs[$nom] !== '' ? $valeurs[$nom] : '(non renseigné)';
    if (($champ['type'] ?? 'texte') === 'long') {
        $lignes[] = '';
        $lignes[] = $champ['libelle'] . ' :';
        $lignes[] = $v;
    } else {
        $lignes[] = $champ['libelle'] . ' : ' . $v;
    }
}

$entetes = [
    'From' => EXPEDITEUR,
    'MIME-Version' => '1.0',
    'Content-Type' => 'text/plain; charset=UTF-8',
    'Content-Transfer-Encoding' => '8bit',
];
$emailVisiteur = $valeurs['email'] ?? '';
if ($emailVisiteur !== '') {
    $entetes['Reply-To'] = $emailVisiteur;
}

$envoye = mail(DESTINATAIRE, mb_encode_mimeheader($sujet, 'UTF-8'), implode("\n", $lignes), $entetes);

if (!$envoye) {
    pageErreur(500);
}

header('Location: ' . PAGE_MERCI, true, 303);
exit;
