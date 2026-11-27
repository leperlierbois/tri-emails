# Tri d'emails marketing

Petite application web (100 % navigateur) pour trier une liste d'emails avant une campagne :

- Import de fichiers **CSV** ou **Excel (.xlsx)**, ou d'archives **ZIP, RAR, 7Z** qui les contiennent (plusieurs fichiers sont fusionnés) ;
- Tri des contacts **actifs** / **inactifs**, avec la raison de chaque exclusion ;
- Filtre par **mot-clé** sur les informations de chaque contact ;
- **Pays** de chaque contact (colonne du fichier ou extension du domaine), filtre et classement par pays ;
- Récapitulatif : total, actifs, actifs liés au mot-clé ;
- Export en CSV ou Excel, copie des emails.

Aucune donnée n'est envoyée sur un serveur. Seuls les noms de domaine (ex. `gmail.com`) sont envoyés au DNS public de Cloudflare pour vérifier qu'ils reçoivent des emails.

## Comment un contact est jugé « actif »

1. Adresse bien formée, sans faute de frappe connue (`gmial.com`…).
2. Pas de doublon, pas d'adresse jetable (liste dans `disposable.js`).
3. Statut du fichier non négatif (désabonné, bounce, plainte, inactif…).
4. Dernière activité (ouverture, clic, achat…) plus récente que la période choisie.
5. Domaine qui reçoit des emails (enregistrement MX).

Limite : un site statique ne peut pas vérifier qu'une boîte mail précise existe (il faudrait un serveur). Pour cela, utilisez un service de vérification spécialisé.

## Mise en ligne sur GitHub Pages

1. Créez un dépôt sur GitHub (ex. `tri-emails`).
2. Déposez tous les fichiers de ce dossier à la racine du dépôt (bouton **Add file → Upload files**), puis **Commit changes**.
3. Dans le dépôt : **Settings → Pages**. Sous **Build and deployment**, choisissez **Deploy from a branch**, branche `main`, dossier `/ (root)`, puis **Save**.
4. Après une à deux minutes, le site est disponible à l'adresse `https://<votre-identifiant>.github.io/tri-emails/`.

Remarque : avec un compte GitHub gratuit, GitHub Pages nécessite un dépôt **public**. Le code est alors visible, mais vos fichiers d'emails ne le sont pas : ils ne quittent jamais votre navigateur.

## Fichiers

- `index.html`, `style.css`, `app.js` : l'application
- `disposable.js` : domaines jetables (modifiable)
- `countries.js` : codes pays → noms
- `exemple.csv` : fichier de démonstration
- `vendor/xlsx.full.min.js` : bibliothèque SheetJS (licence Apache 2.0) pour lire et écrire Excel
- `vendor/libarchive/` : libarchive.js (licence MIT) pour ouvrir les archives ZIP, RAR et 7Z
