# Briefing Claude Code — Add-in Outlook Bella

## Contexte du projet

L'objectif est de créer un **Add-in Outlook** permettant aux agents du service client Belmond d'obtenir un brouillon de réponse généré par **Bella** (l'IA conversationnelle de Belmond) directement depuis leur interface Outlook, en un clic.

Ce add-in est la **Hypothèse 1** du projet d'intégration Bella × Outlook. Il fonctionne via un **MCP Connector** qui fait le lien entre le add-in et l'API Bella.

---

## Ce que doit produire Claude Code

### Structure des fichiers à générer

```
bella-addin/
├── manifest.xml         ← carte d'identité du add-in pour Outlook
├── index.html           ← interface utilisateur du panneau latéral
├── app.js               ← logique métier (Office.js + appel MCP/API)
├── styles.css           ← styles du panneau
└── README.md            ← instructions de déploiement et de test
```

---

## Détail de chaque fichier

### 1. `manifest.xml`

Le manifest doit être un **add-in only manifest** (format XML classique, pas le nouveau unified manifest JSON) pour maximiser la compatibilité.

Il doit déclarer :

- Un `Id` unique (GUID, peut être généré aléatoirement)
- `DisplayName` : "Bella — Assistant Belmond"
- `Description` : "Génère une réponse à cet email grâce à Bella, l'IA de Belmond"
- `IconUrl` et `HighResolutionIconUrl` : pointer vers des placeholders (ex. une image hébergée sur GitHub Pages ou une URL publique)
- `DefaultLocale` : `fr-FR`
- `Hosts` : uniquement `Mailbox` (Outlook)
- `Requirements` : `Mailbox` version `1.1` minimum
- `DesktopFormFactor` avec un `FunctionFile` et un `ExtensionPoint` de type `MessageReadCommandSurface`
- Un bouton dans le ruban Outlook avec le label **"Répondre avec Bella"** qui ouvre un `TaskPane`
- Le `SourceLocation` du TaskPane doit pointer vers `index.html` via une `DefaultValue` configurable (utiliser un placeholder `HTTPS_BASE_URL` à remplacer)

### 2. `index.html`

Interface simple et propre du panneau latéral (largeur ~320px imposée par Outlook).

Elle doit contenir :

- Un titre "Bella" avec le sous-titre "Assistant Belmond"
- Un bouton principal **"Générer une réponse"**
- Une zone de résultat qui affiche :
  - Un spinner / message "Bella réfléchit..." pendant le chargement
  - Le brouillon généré dans une `<textarea>` éditable
  - Un bouton **"Insérer dans la réponse"** (visible uniquement quand un brouillon est disponible)
  - Un bouton **"Régénérer"**
- Une zone d'erreur en cas d'échec (message clair en français)
- Import de `Office.js` via le CDN officiel Microsoft : `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`
- Import de `app.js` et `styles.css`

### 3. `app.js`

Logique principale. Doit être structuré en fonctions claires :

```
Office.onReady()
  └── attache les événements aux boutons

generateResponse()
  ├── lit l'email avec Office.js :
  │     Office.context.mailbox.item.body.getAsync("text", callback)
  │     Office.context.mailbox.item.subject  (propriété synchrone)
  │     Office.context.mailbox.item.from.emailAddress
  ├── affiche le spinner
  ├── appelle callBellaAPI(emailContent)
  └── affiche le résultat ou l'erreur

callBellaAPI(emailBody, emailSubject, emailFrom)
  ├── Pour la phase de test : simuler une réponse après 1.5s
  │     (la vraie URL MCP/API sera branchée plus tard)
  ├── Structure de la requête prévue (à décommenter quand l'API est dispo) :
  │     fetch(BELLA_MCP_URL, {
  │       method: "POST",
  │       headers: { "Content-Type": "application/json", "Authorization": "Bearer TOKEN" },
  │       body: JSON.stringify({ emailBody, emailSubject, emailFrom })
  │     })
  └── retourne le brouillon sous forme de string

insertDraft(draftText)
  └── Office.context.mailbox.item.body.setAsync(
        draftText,
        { coercionType: Office.CoercionType.Text },
        callback
      )
```

**Variables de configuration** en haut du fichier (faciles à remplacer) :

```javascript
const BELLA_MCP_URL = "https://VOTRE_URL_MCP/generate"; // à remplacer
const BELLA_API_TOKEN = "VOTRE_TOKEN"; // à remplacer
const SIMULATE_API = true; // passer à false quand l'API est branchée
```

### 4. `styles.css`

CSS simple et professionnel pour le panneau latéral :

- Font : Segoe UI (standard Microsoft 365)
- Couleurs : s'inspirer de la charte Belmond (tons chauds, élégants) — à défaut, utiliser un thème sobre en blanc/gris avec une couleur d'accent bordeaux/or
- Le bouton principal doit être bien visible
- La textarea du brouillon doit être lisible et redimensionnable
- États visuels : chargement (spinner CSS), erreur (fond rouge clair), succès (fond vert clair)
- Tout doit tenir dans 320px de large sans scroll horizontal

### 5. `README.md`

Instructions complètes pour :

1. **Hébergement sur GitHub Pages** :

   - Créer un repo GitHub public
   - Pousser les fichiers à la racine
   - Activer GitHub Pages sur la branche `main`
   - Remplacer `HTTPS_BASE_URL` dans le manifest par l'URL GitHub Pages obtenue (format : `https://USERNAME.github.io/REPO_NAME`)

2. **Sideloading pour tester** :

   - Aller sur `https://aka.ms/olksideload`
   - Mes add-ins → Add-ins personnalisés → Ajouter depuis un fichier
   - Charger le `manifest.xml`
   - Ouvrir un email dans Outlook → le bouton "Répondre avec Bella" doit apparaître dans le ruban

3. **Tester le add-in** :

   - Ouvrir un email reçu
   - Cliquer sur "Répondre avec Bella" dans le ruban
   - Vérifier que le panneau s'ouvre
   - Cliquer sur "Générer une réponse"
   - Vérifier que la simulation fonctionne (brouillon fictif affiché)
   - Cliquer sur "Insérer dans la réponse" et vérifier l'insertion dans la zone de rédaction

4. **Brancher la vraie API Bella** (étape suivante) :
   - Obtenir l'URL du MCP Connector et le token d'authentification
   - Remplacer `BELLA_MCP_URL` et `BELLA_API_TOKEN` dans `app.js`
   - Passer `SIMULATE_API = false`

---

## Contraintes techniques importantes

| Contrainte          | Détail                                                                                           |
| ------------------- | ------------------------------------------------------------------------------------------------ |
| HTTPS obligatoire   | Outlook refuse tout contenu non-HTTPS, même en dev                                               |
| Office.js CDN       | Toujours charger depuis `appsforoffice.microsoft.com`, jamais en local                           |
| Largeur panneau     | 320px maximum, tout doit s'adapter                                                               |
| Format manifest     | XML classique (add-in only manifest), pas le format JSON unifié                                  |
| Compatibilité       | Doit fonctionner sur Outlook Web, Outlook Desktop Windows, et Outlook Desktop Mac                |
| Langue              | Interface entièrement en français                                                                |
| Pas d'auth complexe | Pour la phase de test, pas de SSO ni Azure AD — authentification simple par token dans le header |

---

## Ce que Claude Code ne doit PAS faire

- Ne pas utiliser de framework lourd (React, Vue, Angular) — du HTML/CSS/JS vanilla uniquement pour garder le add-in simple et sans pipeline de build
- Ne pas implémenter l'authentification SSO Microsoft (trop complexe pour la phase de test, et peut bloquer le sideloading)
- Ne pas mettre de données sensibles (tokens, URLs internes) en dur dans le code — utiliser des placeholders clairs
- Ne pas utiliser `item.body.setAsync` avec `coercionType: HTML` sans test préalable (peut causer des problèmes selon la version d'Outlook)

---

## Résultat attendu

À l'issue de la génération, on doit pouvoir :

1. Pousser les fichiers sur GitHub Pages (5 minutes)
2. Remplacer `HTTPS_BASE_URL` dans le manifest par l'URL GitHub Pages
3. Sideloader le manifest sur `https://aka.ms/olksideload`
4. Ouvrir un email dans Outlook et voir le bouton "Répondre avec Bella"
5. Cliquer et obtenir un brouillon simulé dans le panneau latéral
6. Insérer ce brouillon dans la zone de réponse

**Sans avoir à installer Node.js, npm, ou quoi que ce soit d'autre.**
