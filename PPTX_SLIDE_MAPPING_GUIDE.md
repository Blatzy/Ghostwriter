# Guide du Slide Mapping PPTX - Ghostwriter

## Vue d'ensemble

Le système de slide mapping permet de personnaliser complètement la structure et l'apparence des présentations PowerPoint générées par Ghostwriter.

## Types de slides

### Slides Statiques vs Dynamiques

#### Slides **STATIQUES**
- **Utilisation** : Slides avec contenu fixe ou variables Jinja2 simples
- **Exemples** : Page de titre, agenda personnalisé, slide de conclusion
- **Fonctionnement** :
  - Le layout est copié tel quel depuis le masque PowerPoint
  - Les variables Jinja2 dans le texte sont remplacées
  - La mise en forme du masque est **entièrement préservée**
- **Variables Jinja2 disponibles** :
  - `{{client.name}}`, `{{client.short_name}}`
  - `{{project.codename}}`, `{{project.type}}`, `{{project.start_date}}`, `{{project.end_date}}`
  - `{{report.title}}`, `{{report.complete_date}}`
  - `{{company.name}}`, `{{company.email}}`
  - `{{now}}` (date/heure actuelle)

#### Slides **DYNAMIQUES**
- **Utilisation** : Slides générées par boucle avec données du rapport
- **Exemples** : Slides individuelles de findings, observations
- **Fonctionnement** :
  - Le code Python remplit automatiquement les placeholders
  - Une slide est créée pour chaque finding/observation
  - Utilise la logique de mapping des placeholders (voir ci-dessous)

---

## Variables Jinja2 pour Slides Dynamiques

### Individual Finding Slide (DYNAMIQUE recommandé)

Au lieu d'utiliser des indices de placeholders, utilisez simplement des **variables Jinja2** dans vos zones de texte PowerPoint :

| Variable Jinja2 | Contenu | Exemple |
|----------------|---------|---------|
| `{{ title }}` | Titre du finding | SQL Injection |
| `{{ severity }}` | Niveau de sévérité | Critical, High, Medium, Low |
| `{{ description }}` | Description de la vulnérabilité | L'application est vulnérable à... |
| `{{ impact }}` | Impact de la vulnérabilité | Compromission de la base de données |
| `{{ affected_entities }}` | Systèmes affectés | Server1, Server2, App3 |
| `{{ mitigation }}` | Recommandations | Utiliser prepared statements |
| `{{ recommendation }}` | Alias de mitigation | (même contenu) |
| `{{ replication }}` | Étapes de reproduction | 1. Naviguer vers... |
| `{{ replication_steps }}` | Alias de replication | (même contenu) |
| `{{ host_detection }}` | Techniques de détection hôte | Vérifier les logs Apache... |
| `{{ network_detection }}` | Techniques de détection réseau | Surveiller le trafic SQL... |
| `{{ references }}` | Références | OWASP A1, CWE-89 |
| `{{ cvss_score }}` | Score CVSS | 9.8 |
| `{{ cvss_vector }}` | Vecteur CVSS | CVSS:3.1/AV:N/AC:L... |

**Comment configurer votre layout PowerPoint :**

1. Ouvrez le **Masque des diapositives** dans PowerPoint
2. Sélectionnez le layout que vous voulez utiliser pour les findings
3. Ajoutez des **zones de texte** avec des variables Jinja2 :
   - Insertion → Zone de texte
   - Tapez par exemple : `{{ description }}`
   - Ajoutez du texte autour si besoin : `Description: {{ description }}`
4. Appliquez la mise en forme souhaitée (polices, couleurs, bordures)
5. Les variables seront automatiquement remplacées lors de la génération

**Exemple de layout Finding :**

```
┌─────────────────────────────────────────┐
│  {{ title }} [{{ severity }}]           │
├─────────────────────────────────────────┤
│  Description                            │
│  {{ description }}                      │
├──────────────┬──────────────────────────┤
│ Impact       │ Mitigation               │
│ {{ impact }} │ {{ mitigation }}         │
└──────────────┴──────────────────────────┘
```

### Individual Observation Slide (DYNAMIQUE recommandé)

| Variable Jinja2 | Contenu | Exemple |
|----------------|---------|---------|
| `{{ title }}` | Titre de l'observation | Bonne pratique détectée |
| `{{ description }}` | Description | L'équipe utilise MFA... |

**Exemple simple :**
```
┌─────────────────────────────────────────┐
│  {{ title }}                            │
├─────────────────────────────────────────┤
│  {{ description }}                      │
└─────────────────────────────────────────┘
```

### Autres slides dynamiques

Les autres slides (Agenda, Timeline, etc.) sont générées par code Python et n'utilisent pas de variables Jinja2 personnalisables pour le moment.

---

## Configuration dans l'interface admin

### 1. Accéder à la configuration
- Allez dans **Reporting → Report Templates**
- Sélectionnez votre template PPTX
- Cliquez sur **"Configure Slide Mapping"**

### 2. Pour chaque type de slide

#### Champs disponibles :
- **Layout** : Choisissez quel layout du masque utiliser (détecté automatiquement)
- **Mode** :
  - `Dynamic` : Le code Python remplit les placeholders
  - `Static` : Copie le layout avec rendu Jinja2
- **Enabled** : Activé/désactivé ce type de slide
- **Position** : Ordre d'apparition (utilisez les boutons ↑/↓)

### 3. Slides personnalisées (Custom Slides)

Utilisez la section "Custom Slides" pour ajouter des slides statiques personnalisées :
- **Type/Name** : Identifiant unique (ex: `company_intro`)
- **Layout** : Layout du masque à utiliser
- **Mode** : Généralement `Static`
- **Position** : Où insérer la slide

---

## Exemples d'utilisation

### Exemple 1 : Slide de titre statique avec variables Jinja2

**Dans votre masque PowerPoint** :
- Créez un layout "Custom Title"
- Ajoutez du texte : `Audit de {{client.name}}`
- Sous-titre : `Projet : {{project.codename}}`
- Date : `Du {{project.start_date}} au {{project.end_date}}`

**Configuration** :
- Type : Custom Slide
- Layout : "Custom Title"
- Mode : **Static**
- Position : 1

**Résultat** : Les variables seront remplacées automatiquement, la mise en forme du masque sera conservée.

### Exemple 2 : Finding slide structurée avec variables Jinja2

**Dans votre masque PowerPoint** :
- Layout "Finding Detail"
- Zone de texte titre : `{{ title }} [{{ severity }}]` (Police 32pt, rouge)
- Zone de texte description : `{{ description }}` (Grande zone)
- Zone de texte gauche : `Impact: {{ impact }}`
- Zone de texte droite : `Mitigation: {{ mitigation }}`

**Configuration** :
- Type : Individual Finding Slide
- Layout : "Finding Detail"
- Mode : **Dynamic**
- Position : 10

**Résultat** : Pour chaque finding, une slide est générée avec :
- Titre : "SQL Injection [CRITICAL]"
- Description : Texte complet de la vulnérabilité
- Impact : "Compromission de la base de données"
- Mitigation : "Utiliser prepared statements"

Toutes les variables `{{ ... }}` sont automatiquement remplacées !

### Exemple 3 : Désactiver certaines slides

Si vous ne voulez pas de slide "Attack Path Overview" :
- Trouvez "Attack Path Overview" dans la liste
- Décochez **Enabled**
- La ligne devient grisée
- Cette slide ne sera pas générée

---

## Dépannage

### Problème : Variable Jinja2 non remplacée (ex: `{{project.end_date}}` reste tel quel)

**Solutions** :
1. Vérifiez que la variable est bien orthographiée
2. Consultez les logs Django pour voir les variables disponibles :
   ```
   DEBUG: Jinja2 context - project keys: ['codename', 'start_date', 'end_date', ...]
   ```
3. Vérifiez que la donnée existe dans la base de données (pas NULL)

### Problème : Placeholder non rempli dans une slide dynamique

**Solutions** :
1. Vérifiez l'**index du placeholder** dans PowerPoint :
   - Masque des diapositives → Sélectionner la forme → Vérifier les propriétés
2. Le placeholder 1 est réservé pour la description
3. Si vous voulez Impact/Recommendations, utilisez placeholders 2 et 3

### Problème : Mise en forme perdue

**Pour les slides statiques** : La mise en forme devrait être préservée automatiquement
**Pour les slides dynamiques** : Définissez la mise en forme dans le **masque PowerPoint**, pas dans les slides générées

### Problème : Erreur "no placeholder with idx == X"

Le layout sélectionné n'a pas de placeholder à cet index. Solutions :
- Ajoutez des placeholders dans le masque PowerPoint
- OU changez de layout qui a les placeholders nécessaires
- OU le système utilisera un fallback automatique

---

## Bonnes pratiques

1. **Testez avec un petit rapport d'abord** avant de configurer tout
2. **Documentez vos layouts** : Notez quel layout sert à quoi
3. **Utilisez des noms de layout explicites** dans PowerPoint (ex: "Finding - 2 colonnes")
4. **Sauvegardez votre configuration** : Exportez le template configuré
5. **Mode dynamique pour les boucles** : Findings, observations
6. **Mode statique pour le contenu fixe** : Titre, agenda personnalisé, conclusion

---

## Support et logs

En cas de problème, activez le mode DEBUG dans Django et consultez les logs :

```python
# settings.py
LOGGING = {
    'loggers': {
        'ghostwriter.modules.reportwriter': {
            'level': 'DEBUG',
        },
    },
}
```

Les logs montreront :
- Variables Jinja2 disponibles
- Placeholders trouvés/non trouvés
- Erreurs de rendu

---

**Questions ?** Consultez la documentation Ghostwriter ou ouvrez une issue sur GitHub.
