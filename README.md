# SNCF Code VBA - Gestion des Astreintes et Affectations LAJ

Suite de 3 projets Excel VBA pour l'automatisation complète de la gestion des astreintes SNCF (Direction des Ressources Humaines / Ligne d'Assistance Juridique).

## 📁 Projets

### 1️⃣ Gestionnaire d'Astreintes et Affectations
**Fichier :** `01_Gestionnaire_Astreintes_Affectations_LAJ.xlsm`

Gestion opérationnelle des assignations d'agents aux astreintes.

**Fonctionnalités :**
- Création de nouveaux planning d'astreintes
- Affectation d'agents avec validation automatique des compétences
- Validation des assignations
- Statistiques de couverture
- Export pour équipe de dispatch
- Audit trail complète de tous les changements

**Données de test incluses :**
- 4 agents (AG001-AG004) avec compétences
- 5 gares de référence (GA001-GA005)
- 3 astreintes pré-chargées

---

### 2️⃣ Contrôle de Qualité et Consolidation
**Fichier :** `02_Controle_Qualite_Consolidation.xlsm`

Détection et correction automatique des anomalies de données.

**Fonctionnalités :**
- Import de données brutes de multiples sources
- Détection des anomalies (doublons, références invalides, champs vides, formats)
- Correction automatique en masse
- Rapport de conformité généré automatiquement
- Historique complet des corrections

**Anomalies de test incluses :**
- Doublon : AG001 apparaît 2 fois
- Champ vide : Contrat de AG002
- Référence invalide : Gare inexistante pour AG003
- Incohérence de format

---

### 3️⃣ Reporting Mensuel Automatisé
**Fichier :** `03_Reporting_Mensuel_Automatise.xlsm`

Génération automatique de rapports et KPI.

**Fonctionnalités :**
- Génération de rapports mensuels
- Calcul de 4 KPI principaux avec tendances
- Détection d'alertes et deadlines (dépassements de délais)
- Export format Power BI
- Historique 12 mois

**KPI calculés :**
- Coverage % (couverture des astreintes)
- Absence Rate % (taux d'absence)
- Turnover % (taux de rotation)
- Zone Coverage % (couverture par zone)

**Alertes pré-configurées :**
- Visites médicales (16 jours)
- Renouvellement contrats (130 jours)
- Certifications (265 jours)

---

## 🎯 Architecture

```
Projet 1 (Opérationnel)
    ↓ données brutes
Projet 2 (Qualité)
    ↓ données validées
Projet 3 (Reporting)
    ↓ rapports & KPI
```

## 💻 Prérequis

- Excel 2016+ avec macros activées
- Windows 7+
- Aucune dépendance externe

## 🚀 Utilisation

1. **Ouvrir un fichier** : Double-cliquer sur le fichier .xlsm
2. **Aller à l'onglet "Accueil"** : Interface principale avec boutons
3. **Cliquer les boutons** : Lancer les macros
4. **Vérifier les résultats** : Consulter les feuilles de données

### Exemple Projet 1

```
1. Cliquer "Nouveau Planning" → Créer une astreinte
2. Cliquer "Affecter Agent" → Assigner un agent
   - Agent : AG001
   - Astreinte : AST001
3. Cliquer "Valider" → Vérifier les anomalies
4. Cliquer "Stats" → Voir taux de couverture
5. Voir résultats dans onglets "Affectations" et "AuditTrail"
```

## 📊 Données de Test

Tous les fichiers incluent des données de test réalistes :

**Agents (4)**
- AG001 : Pierre Dupont, CDI, GA001, Quai+Accueil
- AG002 : Marie Martin, CDI, GA002, Accueil
- AG003 : Jean Lacroix, CDD, GA001, Quai+Sécurité
- AG004 : Sophie Bernard, Stage, GA003, Maintenance

**Gares (5)**
- GA001-GA005 : Région Île-de-France

**Astreintes (3)**
- AST001-AST003 : Mai 2026

---

## 🔒 Conformité

- ✅ RGPD : Aucune donnée sensible en dur
- ✅ Audit Trail : Traçabilité complète
- ✅ Validation : Tous les contrôles inclus
- ✅ Robustesse : Gestion d'erreurs complète

---

## 📝 Notes de développement

- **Langage** : VBA (Excel Macro Language)
- **Modules** : 1 module par projet (ModuleMacros)
- **Macros** : 5 + 4 + 4 = 13 macros totales
- **Feuilles** : 27 feuilles réparties sur 3 projets
- **Format** : Excel Macro-Enabled Workbook (.xlsm)

---

## 👤 Contexte Métier

Simulation d'un système d'automatisation pour la gestion des astreintes à SNCF (Direction Ressources Humaines / Ligne d'Assistance Juridique).

Démontre :
- ✅ Maîtrise de VBA et Excel
- ✅ Architecture logicielle 3-couches
- ✅ Logique métier SNCF réaliste
- ✅ Gestion de données complexes
- ✅ Automatisation bout-en-bout
- ✅ Rigueur et documentation

---

## 📄 License

Confidential - Simulation pour candidature SNCF/DRH

---

**Créé le 24/04/2026**

Pour plus d'informations : contactez l'auteur
