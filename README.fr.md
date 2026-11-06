[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Modèle Excel de suivi de la paie et de planification de la capacité de main-d'œuvre | Calculateur d'heures supplémentaires, de FTE et de coûts de main-d'œuvre
<img width="1920" height="790" alt="image" src="https://github.com/user-attachments/assets/f36d608b-7a02-4a4b-b72a-ec7702d528f5" />


![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Browser%20%2B%20Microsoft%20Excel-success)
![Tool](https://img.shields.io/badge/Tool-Payroll%20%26%20Workforce%20Planning-blue)

**Ce modèle Excel de suivi de la paie prêt à l'emploi et ce tableur de planification de la capacité de main-d'œuvre aident les dirigeants de petites entreprises, les généralistes RH et les responsables des opérations à calculer les coûts mensuels de main-d'œuvre, à suivre les heures supplémentaires des employés et à piloter la capacité d'effectif en équivalent temps plein (FTE) en moins d'une minute. Conçu avec des formules matricielles dynamiques automatisées pour Microsoft 365 et Excel 2021+, ce suivi des heures des employés et ce calculateur de coûts de main-d'œuvre clés en main ne demandent aucune configuration de macro, aucun glisser de formule et aucune compétence en analyse de données : il suffit de copier-coller vos données brutes de feuilles de temps pour générer un tableau de bord de paie et de bande passante d'effectif prêt pour la direction.**

Testez le calculateur web interactif gratuitement dans votre navigateur, ou téléchargez le tableur Excel déverrouillé pour vos cycles de paie mensuels récurrents (garantie satisfait ou remboursé de 30 jours).

> 🌐 **Calculateur gratuit dans le navigateur** → [Lancez le tableau de bord interactif gratuit de paie et de planification de la capacité de main-d'œuvre dans votre navigateur (application web HTML, sans connexion requise)](https://hyvoid.github.io/workforce-planning-payroll-excel-dashboard/)
>
> 📥 **Modèle Excel prêt à l'emploi** → [Téléchargez le modèle Excel automatisé de suivi de la paie d'entreprise et de planification de la capacité d'effectif (.xlsx pour Excel 2021+ / Microsoft 365)](https://www.theseusworkshop.com/l/ufscrp?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=payroll-workforce-planning)

---

## À quelles questions sur les coûts de main-d'œuvre, les heures supplémentaires et la capacité d'effectif ce tableur répond-il ?

Traiter les salaires bruts mensuels et les retenues fiscales ne vous dit que ce qui a déjà quitté votre compte bancaire. Pour éviter les dépassements budgétaires et les goulots d'étranglement des plannings le mois suivant, les managers ont besoin d'un **tableur unifié d'analyse des coûts de main-d'œuvre et de planification des effectifs** qui relie les dépenses de paie aux heures de travail restantes :

- **Capacité FTE et vitesse d'épuisement de la charge de travail :** quels services consomment leurs **heures contractuelles annuelles en équivalent temps plein (FTE)** plus vite que prévu au budget ?
- **Fuites de coûts d'heures supplémentaires (multiplicateurs 1.5x et 2.0x) :** la hausse des **dépenses mensuelles en heures supplémentaires** provient-elle d'un pic saisonnier temporaire ou d'une pénurie structurelle d'effectif ?
- **Conformité des heures travaillées et prévention de l'épuisement :** quels employés précis atteindront leurs **plafonds annuels de conformité des heures de travail** avant le pic d'activité du quatrième trimestre ?
- **Décomposition de l'écart de masse salariale brute :** la progression des **coûts mensuels de main-d'œuvre** résulte-t-elle d'augmentations de salaires de base, d'heures supplémentaires non validées ou de primes d'équipe complémentaires ?
- **Bande passante des projets et affectation des ressources :** votre équipe opérationnelle actuelle peut-elle absorber de nouvelles commandes clients ce mois-ci avec la **capacité d'effectif inutilisée** existante, ou devez-vous recruter ?
- **Prévision des goulots d'étranglement par service :** quelle unité d'affaires épuisera la première ses heures de main-d'œuvre disponibles pendant les mois de production de pointe ?

Les logiciels de paie standards exportent des journaux historiques bruts. Ce **tableau de bord Excel de gestion de la main-d'œuvre (WFM)** transforme ces lignes brutes de feuilles de temps en un **système automatisé d'aide à la décision opérationnelle (DSS)**. Comme chaque calcul repose sur des fonctions matricielles dynamiques natives d'Excel (`LET`, `XLOOKUP`, `SUMIFS`, `FILTER`) sans aucune macro VBA, le fichier s'ouvre proprement sur tout poste de travail moderne, sans blocage de sécurité ni script cassé.

---

## Ce qu'il vous aide à suivre : correspondance entre difficultés RH et opérationnelles et solutions

Utilisez ce tableau de référence structuré pour voir comment des **problèmes précis de suivi de la paie, d'audit des heures supplémentaires et de capacité d'effectif** correspondent directement à des feuilles automatisées du classeur :

| Point de douleur opérationnel fréquent (ce qui vous bloque) | Ce que ce modèle Excel suit (indicateurs et formules automatisés) | Solution managériale immédiate (résultat décisionnel en 1 minute) |
| :--- | :--- | :--- |
| **Dépassements budgétaires imprévus sur les heures supplémentaires mensuelles** | Suit les heures normales par rapport aux **heures supplémentaires à 1.5x et 2.0x**, les ratios heures supplémentaires/salaires normaux et les pics de primes par service. | Le **tableau de bord de suivi des heures supplémentaires en temps réel** détermine si le coût des heures supplémentaires relève d'une anomalie d'une semaine ou d'une pénurie chronique d'effectif nécessitant un recrutement. |
| **Épuisement des employés et dépassement des plafonds annuels d'heures** | Suit les **heures réellement travaillées depuis le début de l'exercice (YTD)** par rapport aux heures contractuelles annuelles en FTE et aux seuils d'alerte d'utilisation personnalisés (par exemple, plafond à 85%). | Le **suivi des heures contractuelles annuelles** signale les employés à haut risque des mois avant qu'ils ne franchissent les limites légales du travail ou ne s'épuisent pendant les périodes de forte activité. |
| **Consolidation laborieuse des feuilles de temps multi-services** | Consolide **le salaire de base, les salaires des heures supplémentaires et les primes complémentaires** sur jusqu'à 7 onglets de service dans un grand livre maître unique. | Le **moteur automatisé de synthèse de la paie** élimine les erreurs de copier-coller manuel et les liens `VLOOKUP` cassés entre les tableurs des superviseurs d'équipe. |
| **Planification à l'aveugle de nouveaux projets ou de nouvelles équipes** | Calcule **les heures annuelles disponibles restantes** par employé et les **taux d'utilisation nets de la main-d'œuvre (%)** par service. | Le **prévisionniste de capacité d'effectif** montre la bande passante exacte disponible afin d'attribuer de nouvelles charges aux équipes sous-utilisées sans aucun coût de recrutement supplémentaire. |
| **Incapacité à expliquer les pics mensuels de coûts de main-d'œuvre** | Sépare **les coûts salariaux de base fixes** des **coûts variables liés aux comportements** (multiplicateurs d'heures supplémentaires, primes d'équipe, bonus). | La **vue des écarts de coûts de main-d'œuvre par service** offre aux responsables financiers et opérationnels une décomposition instantanée pour les revues budgétaires mensuelles. |

---

## Tutoriel de démarrage rapide : comment lancer votre analyse de paie et de capacité en 1 minute (4 étapes simples)

Vous n'avez pas besoin de compétences avancées sur Excel, d'une formation aux tableaux croisés dynamiques ni d'une expérience en modélisation financière pour utiliser cet outil. Les calculs de main-d'œuvre sous-jacents sont complexes, **mais toutes les règles de correspondance, les multiplicateurs d'heures supplémentaires et les formules matricielles dynamiques sont préconstruits et s'étendent automatiquement**. Vous n'avez même pas besoin de double-cliquer sur les coins de cellules ni de faire glisser des formules vers le bas : collez vos données et le tableau de bord se met à jour immédiatement.

### Étape 1 : définissez une seule fois vos règles globales de paie et d'heures supplémentaires
Ouvrez l'onglet `Settings` pour saisir les paramètres de référence de votre entreprise :
- Mois de début de l'exercice fiscal et horizon de planification
- Multiplicateurs standards de rémunération des heures supplémentaires (par exemple, `1.5x` pour les heures supplémentaires normales, `2.0x` pour le taux jours fériés/week-end)
- Seuil d'alerte d'utilisation de la capacité des employés (par exemple, déclencher l'alerte à `85%` des heures contractuelles annuelles)
- Symbole de la devise locale et noms des services

*Résultat :* chaque graphique du tableau de bord, chaque alerte de conformité et chaque table de paie se mettent à jour automatiquement dans tout le classeur.

### Étape 2 : copiez-collez votre liste d'employés et vos feuilles de temps mensuelles
Collez votre liste du personnel dans l'onglet `Employee_Master`, puis collez les relevés mensuels de pointage dans les feuilles de service :
- ID employé, nom, service et heures contractuelles annuelles
- Heures de travail normales enregistrées pour la période de paie
- Heures supplémentaires validées (`1.5x` / `2.0x`)
- Primes complémentaires ou bonus d'équipe versés

*Résultat :* copiez-collez directement les exports CSV bruts de **ADP, Gusto, QuickBooks Payroll, Workday, BambooHR ou de pointeuses manuelles**. Les formules dynamiques préconstruites associent automatiquement les ID employés et s'étendent à votre nombre de lignes : aucun glisser de formule nécessaire.

### Étape 3 : consultez votre tableau de bord instantané des coûts de main-d'œuvre et de la capacité d'effectif
Cliquez sur `Payroll_Summary`, `Annual_Hours_Tracker` ou `Dashboard` pour lire votre analyse finalisée :
- Dépense totale de masse salariale brute mensuelle (base vs heures supplémentaires vs primes)
- Heures contractuelles annuelles restantes par employé
- Pourcentage d'utilisation des FTE par service et équilibre de la charge de travail
- Signalements automatiques de conformité rouge/orange pour les collaborateurs surchargés

*Résultat :* identifiez vos goulots d'étranglement opérationnels exacts en **60 secondes** et faites une capture d'écran ou imprimez directement les graphiques pour votre réunion de direction hebdomadaire.

### Étape 4 : réutilisez le modèle chaque mois et gagnez des heures en pleine saison (obtenez le fichier Excel)
À la clôture de la période de paie du mois suivant, collez les lignes de feuilles de temps du nouveau mois. Le modèle conserve les mois historiques et met à jour automatiquement votre trajectoire depuis le début de l'exercice (YTD).

> 💡 **Transformez des heures de lutte avec des tableurs en une routine de copier-coller d'une minute :**
> Pourquoi perdre des soirées de haute saison à réparer des formules cassées ou à construire des tableaux croisés dynamiques de zéro ? Même sans aucune expérience en analyse de données, ce modèle Excel pré-ingénieré fait tout le gros travail mathématique à votre place. Collez vos chiffres de paie bruts, repérez vos décisions managériales clés en 1 minute et revenez à la gestion de votre activité principale.
>
> 👉 **[Téléchargez ici le modèle Excel complet et réutilisable de paie et de planification de la capacité de main-d'œuvre (.xlsx)](https://www.theseusworkshop.com/l/ufscrp?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=workforce-planning-payroll-dashboard)** pour conserver les données de votre entreprise en local, réaliser un nombre illimité d'analyses de paie mensuelles hors ligne et payer une seule fois, sans frais SaaS récurrents.

---

## Comparaison : pourquoi utiliser ce modèle Excel automatisé plutôt que des tableurs manuels ou un logiciel RH ancien ?

| Scénario opérationnel et indicateur | Tableurs Excel manuels et exports de paie basiques (sans cet outil) | Boîte à outils Excel automatisée de capacité d'effectif et de paie (avec cet outil) |
| :--- | :--- | :--- |
| **Audit mensuel de la paie et des heures supplémentaires** | Les comptables auditent les salaires bruts uniquement *après* la clôture des périodes comptables et la dépense des liquidités, sans aucune marge pour prévenir les dépassements budgétaires liés aux heures supplémentaires. | Le suivi continu par période de paie **du salaire de base, des majorations d'heures supplémentaires à 1.5x/2.0x et des primes** met en évidence les pics de coûts avant que le budget trimestriel ne soit épuisé. |
| **Maintenance des formules et saisie des données** | Les utilisateurs passent des heures à faire glisser des formules vers le bas, à corriger les erreurs `#REF!` et `VLOOKUP` cassées et à mettre à jour manuellement les plages de graphiques chaque mois. | Les **formules matricielles dynamiques natives** préconstruites associent automatiquement les lignes collées et actualisent tous les graphiques KPI — **aucun double-clic ni glisser de formule nécessaire**. |
| **Heures contractuelles annuelles et conformité** | Les responsables d'équipe suivent les heures FTE sur des brouillons déconnectés et ne découvrent les infractions aux limites d'heures de travail que lorsque les employés s'épuisent ou démissionnent. | Le calcul automatique **des heures travaillées depuis le début de l'exercice (YTD) par rapport aux plafonds contractuels annuels** déclenche des alertes visuelles instantanées lorsque le personnel franchit votre seuil d'utilisation de 85%. |
| **Consolidation des données multi-services** | Les RH et la finance perdent 1 à 2 jours par mois à fusionner des onglets de feuilles de temps incohérents provenant de 5 à 7 responsables de service différents. | Jusqu'à **7 onglets de saisie départementaux standardisés** alimentent automatiquement un tableau de bord maître centralisé unique dès le collage des données brutes de pointage au format CSV. |
| **Calendrier des décisions d'effectif et de recrutement** | Les managers demandent de nouvelles embauches sur la base d'une intuition ou d'un pic d'activité d'une seule semaine, sans vérifier si les équipes voisines disposent d'heures disponibles. | Les **taux d'utilisation FTE par service** comparés côte à côte prouvent si vous avez besoin d'effectifs supplémentaires ou si vous pouvez réaffecter des tâches aux équipes disposant d'une capacité restante. |

---

## À qui s'adresse ce modèle Excel de paie et de planification de la capacité ? (par rôle et par scénario de recherche)

Ce tableur clés en main est conçu pour les managers, les dirigeants d'entreprise et les analystes qui disposent d'exports bruts de feuilles de temps ou de paie (provenant d'ADP, Gusto, Paychex, QuickBooks ou Rippling) et qui ont besoin d'une **visibilité immédiate sur le contrôle des coûts et la charge de travail** sans payer plus de $5,000+ par an pour un logiciel d'entreprise de Workforce Management (WFM) :

- **Dirigeants de petites entreprises et fondateurs de startups à la recherche d'un `Simple Monthly Payroll & Overtime Tracker Excel Sheet` :** cessez de deviner pourquoi la paie mensuelle ne cesse d'augmenter. Collez les chiffres de votre cycle de paie pour comparer précisément le salaire normal et le salaire des heures supplémentaires de votre équipe en 1 minute.
- **Responsables RH et responsables People Ops à la recherche d'un `Annualized Hours & Employee Utilization Tracking Template` :** suivez les heures travaillées depuis le début de l'exercice par rapport aux contrats FTE annuels, prévenez l'épuisement du personnel et gardez une longueur d'avance sur les plafonds légaux d'heures de travail.
- **Administrateurs de paie et comptables ayant besoin d'un `Multi-Department Timesheet Consolidation Spreadsheet` :** agrégez les heures normales, les multiplicateurs d'heures supplémentaires et les primes d'équipe de jusqu'à 7 services dans un tableau de synthèse unique et propre, sans écrire une seule macro.
- **Responsables des opérations, d'entrepôt et d'équipe cherchant un `Staffing Capacity & Workload Planning Calculator` :** vérifiez combien d'heures de main-d'œuvre utilisables il reste à chaque équipe avant la haute saison afin d'équilibrer les plannings et d'éviter les paniques d'heures supplémentaires de dernière minute.
- **Directeurs financiers, comptables et directeurs financiers à temps partagé ayant besoin d'un `Labor Cost Variance & Headcount Budget Model` :** auditez les vitesses d'épuisement des coûts de main-d'œuvre par service, distinguez la croissance structurelle des salaires des fuites liées aux heures supplémentaires et fournissez à vos clients des graphiques d'effectif prêts pour le conseil d'administration.

---

## Quand avez-vous besoin d'un tableur prêt à l'emploi pour les heures des employés et les coûts de main-d'œuvre ? (scénarios courants)

Si vous reconnaissez l'une de ces situations opérationnelles du quotidien, ce modèle vous apporte une solution immédiate dès aujourd'hui :

1. **« Ma facture de paie mensuelle a augmenté de 18%, et je n'arrive pas à identifier quel service ou quelle équipe en heures supplémentaires en est la cause. »** → Collez vos pointages mensuels pour classer instantanément les services par ratio d'heures supplémentaires et dépenses de primes.
2. **« Nous entrons dans la haute saison Q3/Q4 et je ne sais pas si mon personnel actuel pourra absorber les commandes supplémentaires. »** → Consultez la colonne **Capacité annuelle restante** pour voir les heures exactement disponibles par équipe avant de vous engager sur de nouvelles embauches.
3. **« Je déteste passer mon vendredi après-midi à réparer des formules Excel cassées et à mettre à jour des tableaux croisés dynamiques pour la réunion des opérations du lundi. »** → Collez vos données brutes dans les tables de saisie préformatées et laissez le tableau de bord automatisé générer vos visuels de réunion en 60 secondes.
4. **« Notre prestataire de paie traite correctement les bulletins, mais ses tableaux de bord de reporting sont inutilisables pour la planification de la capacité. »** → Exportez votre fichier CSV brut de pointage depuis votre prestataire de paie, déposez-le dans ce classeur et obtenez la vue stratégique de capacité que votre logiciel de paie omet.

---

## La logique métier et la méthodologie : 3 angles morts cachés de la paie expliqués simplement

Vous n'avez pas besoin d'un MBA ni d'un diplôme de finance pour mener une analyse intelligente de la main-d'œuvre. La plupart des entreprises perdent de l'argent sur la main-d'œuvre parce que les rapports comptables standards masquent trois vérités opérationnelles simples. Voici la logique métier en langage clair intégrée à ce modèle Excel — et comment elle protège votre marge :

### 1. Angle mort n°1 : suivre les « euros dépensés » masque les « heures restantes » (traiter la main-d'œuvre comme un stock d'heures)
* **Le fait objectif :** les logiciels de paie standards ne comptabilisent l'argent qu'*après* que vous l'avez dépensé. Ils traitent la main-d'œuvre comme une facture mensuelle sans plafond plutôt que comme un volume d'heures fixe acheté pour l'année.
* **Pourquoi cela vous coûte de l'argent :** supposez qu'un employé soit contractuellement engagé pour 2,000 heures de travail par an. Si une planification d'équipe médiocre consomme 1,300 de ces heures dès juin, un rapport comptable mensuel standard paraîtra encore normal — jusqu'à ce qu'octobre arrive et que l'employé atteigne son plafond légal ou contractuel d'heures pendant votre trimestre commercial le plus chargé. Vous êtes alors contraint d'interrompre les commandes, de payer des heures supplémentaires à 1.5x–2.0x ou de régler des frais de recrutement pour des intérimaires d'urgence.
* **Comment le modèle le résout automatiquement :** le classeur déduit chaque heure enregistrée du **volume d'heures contractuelles annuelles** de chaque employé. Vous voyez chaque mois votre **stock d'heures restantes** exact, ce qui vous donne plusieurs mois d'avance pour rééquilibrer les charges de travail avant qu'un goulot d'étranglement d'effectif ne survienne.

### 2. Angle mort n°2 : un seul chiffre de « paie totale » masque les fuites d'heures supplémentaires (salaire de base vs heures supplémentaires variables)
* **Le fait objectif :** regarder la « rémunération brute mensuelle totale » mélange des heures de base productives avec des heures de pénalité coûteuses (heures supplémentaires à 1.5x, paie de jours fériés à 2.0x et primes ponctuelles).
* **Pourquoi cela vous coûte de l'argent :** payer des heures supplémentaires à 1.5x sur une équipe de 10 personnes pendant trois mois consécutifs coûte souvent plus cher en trésorerie pure que le recrutement d'un salarié supplémentaire à temps plein au taux de base standard de 1.0x — tout en épuisant physiquement vos meilleurs collaborateurs et en augmentant le turnover.
* **Comment le modèle le résout automatiquement :** le moteur de calcul répartit chaque euro entre **salaire de base structurel**, **majorations d'heures supplémentaires** et **primes**. En un coup d'œil, vous savez si des coûts de main-d'œuvre élevés proviennent d'un pic temporaire de deux semaines (conservez l'effectif actuel) ou d'une pénurie permanente de charge de travail (il est moins coûteux de recruter un nouveau membre d'équipe).

### 3. Angle mort n°3 : payer des heures supplémentaires dans une équipe alors qu'une autre dispose de bande passante (planification en silo)
* **Le fait objectif :** lorsque les responsables de service suivent les plannings dans des tableurs séparés, l'équipe A valide régulièrement des heures supplémentaires coûteuses ou demande de nouvelles embauches alors que l'équipe B tourne à 65% de sa capacité.
* **Pourquoi cela vous coûte de l'argent :** chaque heure supplémentaire payée alors qu'une autre équipe qualifiée dispose d'heures contractuelles inutilisées constitue une déduction directe et évitable sur le bénéfice net de votre entreprise.
* **Comment le modèle le résout automatiquement :** en regroupant jusqu'à 7 services dans une **comparaison côte à côte du taux d'utilisation de la main-d'œuvre (%)**, vous repérez immédiatement les équipes en surcharge (>85% de capacité) et celles qui disposent de bande passante (<70% de capacité) pour absorber les travaux à venir sans coût de main-d'œuvre supplémentaire.

---

## Spécifications techniques et compatibilité des fichiers

- **Format :** classeur `.xlsx` natif + aperçu HTML interactif dans le navigateur
- **Compatibilité Excel :** Microsoft Excel 2021+, Microsoft 365 (Windows et Mac) avec fonctions matricielles dynamiques natives
- **Macros / VBA :** 100% sans macro (format `.xlsx` — aucune alerte de sécurité, aucun blocage IT, aucun script VBA cassé)
- **Capacité :** préconfiguré pour la consolidation multi-services (jusqu'à 7 onglets de service localisés alimentant 1 tableau de bord exécutif centralisé)

---

## À propos de Theseus Workshop | Des outils d'aide à la décision métier prêts à l'emploi

**Theseus Workshop** (*Pure utilité. Zéro superflu.*) conçoit des modèles financiers Excel légers et clés en main, des estimateurs de coûts paramétriques et des tableaux de bord d'aide à la décision opérationnelle pour les fondateurs, les responsables des opérations et les équipes financières. Nous remplaçons les logiciels d'entreprise surdimensionnés et les brouillons de tableurs fragiles par des outils Microsoft 365 propres et sans macro qui apportent une clarté managériale immédiate en moins d'une minute.

- 🛒 **Parcourez notre catalogue complet d'outils sur Gumroad :** [Découvrez les modèles financiers Excel, tableaux de bord de trésorerie et modèles d'estimation prêts à l'emploi de Theseus Workshop (theseusworkshop.com)](https://www.theseusworkshop.com/?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20footer&utm_content=payroll-workforce-planning)
- 📥 **Obtenez cette boîte à outils main-d'œuvre et paie :** [Téléchargez le modèle Excel complet de suivi de la paie d'entreprise et de planification de la capacité de main-d'œuvre (.xlsx)](https://www.theseusworkshop.com/l/ufscrp?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20footer&utm_content=payroll-workforce-planning)
- 🌐 **Essayez la version navigateur en direct :** [Ouvrez le tableau de bord interactif gratuit de capacité de main-d'œuvre et de paie (HTML)](https://hyvoid.github.io/workforce-planning-payroll-excel-dashboard/)


---

## Architecture technique & référence des formules

<details>
<summary>Pour les développeurs Excel, les modélisateurs financiers et les analystes de données</summary>

### Architecture relationnelle du classeur

Le tableur fonctionne selon une séparation stricte des responsabilités : configuration, données de référence, entrée transactionnelle et présentation agrégée. Cela garantit une haute intégrité des données et zéro écrasement de formules lors des mises à jour mensuelles.

| Couche architecturale | Feuilles de calcul attribuées | Fonction technique principale |
|---------------------|---------------------|-------------------------|
| **Paramètres globaux** | `Settings` | Variables centralisées (exercice fiscal, multiplicateur d'heures supplémentaires, seuils KPI). |
| **Gestion des données de référence**| `Employee_Master` | Base de données à clé primaire pour les ID employés, les taux de base et les cibles ETP annuelles. |
| **Entrée transactionnelle** | `Dept_1` – `Dept_7` | Environnements de saisie mensuelle des fiches de temps isolés pour les managers de première ligne. |
| **Agrégation des données** | `Payroll_Summary`, `Annual_Hours_Tracker` | Modélisation de données inter-feuilles utilisant des tableaux dynamiques, `SUMIFS` et des algorithmes d'utilisation. |
| **Interface de présentation** | `Dashboard` | Résumés décisionnels en lecture seule, mise en forme conditionnelle et traçage de graphiques. |

### Logique de calcul principale

L'outil utilise strictement des fonctions natives Microsoft 365 / Excel 2021+. **Zéro macro (VBA) ou dépendance Power Query.**

#### Récupération des données employés (XLOOKUP)
```excel
=XLOOKUP(Employee_ID, Employee_Master[Employee_ID], Employee_Master[Hourly_Rate], 0, 0)

```

*Garantit que les grands livres de paie en aval reflètent instantanément les changements de taux effectués dans le jeu de données de référence.*

#### Agrégation de la paie brute inter-départements

```excel
=SUMIFS(Dept_1[Gross_Pay], Dept_1[Employee_ID], [@Employee_ID]) + SUMIFS(Dept_2...

```

*Crée un cumul d'entreprise transparent sans nécessiter d'actualisations complexes de tableaux croisés dynamiques.*

#### Moteur de classification des risques de capacité

```excel
=IFS(
    Utilization_Rate >= Settings!High_Threshold, "High Risk - Burnout",
    Utilization_Rate <= Settings!Low_Threshold, "Under-Utilized - Idle Capacity",
    TRUE, "Optimal Allocation"
)

```

*Pilote la mise en forme conditionnelle du tableau de bord décisionnel pour mettre en évidence les interventions requises.*

</details>

---

## Découvrez d'autres modèles d'opérations métier

Je me spécialise dans la conception de modèles Excel d'aide à la décision légers et à fort impact pour des flux de travail opérationnels complexes. Explorez l'écosystème d'outils au sens large :

* **[project-time-tracking-excel-template](https://github.com/HyVoid/project-time-tracking-excel-template)** — Quels projets dérivent sur les coûts de temps ou la marge ? — Analysez l'affectation des coûts de main-d'œuvre par projet et la rentabilité sur l'ensemble des portefeuilles de projets actifs.
* **[Caf-audit-tool-for-efficient-reviews](https://github.com/HyVoid/Caf-audit-tool-for-efficient-reviews)** — Quelle dimension opérationnelle coûte le plus de chiffre d'affaires et les fuites financières exigent-elles une intervention immédiate ? — Transformez des observations structurées en café en un score de diagnostic pondéré, une estimation des fuites financières et une liste d'actions correctives priorisées.
* **[Restaurant-Menu-Configuration-Modifier-Pricing-Toolkit](https://github.com/HyVoid/Restaurant-Menu-Configuration-Modifier-Pricing-Toolkit)** — Le menu est-il tarifé de façon cohérente sur toutes les combinaisons de modificateurs et toutes les tailles, ou des erreurs de prix cachées attendent-elles de surgir ? — Centralisez les règles de tarification des modificateurs de configuration de menu de restaurant et les données de menu prêtes pour le POS dans une source de vérité unique et gouvernée.
* **[wedding-planner](https://github.com/HyVoid/wedding-planner)** — Quelles tâches, quels coûts et quels délais de planification de mariage nécessitent une coordination ? — Coordonnez la gestion des prestataires, le suivi des tâches et le contrôle budgétaire sur l'ensemble des flux de planification de mariage.
* **[manufacturing-capacity-planning-template](https://github.com/HyVoid/manufacturing-capacity-planning-template)** — La capacité de main-d'œuvre est-elle utilisée efficacement ou l'opération nécessite-t-elle un effectif ou une planification différents ? — Transformez les plans de production et les données de main-d'œuvre en décisions de taux d'utilisation de la capacité, d'affectation des coûts de main-d'œuvre et d'effectif.
* **[cin7-excel-reporting-template](https://github.com/HyVoid/cin7-excel-reporting-template)** — Que disent les données opérationnelles Cin7 sur la performance de l'entreprise ? — Traduisez les données d'inventaire et d'opérations Cin7 en informations exploitables sur la performance de l'entreprise.
* **[employee-performance-dashboard-excel](https://github.com/HyVoid/employee-performance-dashboard-excel)** — Quels employés et quels projets risquent un échec de livraison et où la direction doit-elle intervenir avant l'échéance ? — Reliez la planification annuelle du travail, le suivi de l'exécution, l'évaluation pondérée de la performance et le reporting managérial dans un flux opérationnel continu unique.
* **[rental-property-management-excel-template](https://github.com/HyVoid/rental-property-management-excel-template)** — Quels logements perdent du chiffre d'affaires à cause de vacances non détectées et quelles rotations nécessitent une action immédiate ? — Transformez les réservations en statut d'occupation en temps réel, en intelligence sur les vacances et en alertes opérationnelles sur l'ensemble d'un portefeuille locatif.
* **[warehouse-billing-excel-template](https://github.com/HyVoid/warehouse-billing-excel-template)** — Les frais de facturation d'entrepôt sont-ils exacts et traçables jusqu'aux enregistrements opérationnels sous-jacents ? — Calculez les frais de stockage, de manutention et de transport de l'entrepôt et transformez-les en états de facturation client auditables.
* **[commercial-hvac-plumbing-bid-estimator](https://github.com/HyVoid/commercial-hvac-plumbing-bid-estimator)** — Quel est le prix de soumission défendable et auditable pour ce projet commercial de CVC et de plomberie ? — Estimez et suivez les coûts des projets commerciaux de CVC et de plomberie, du métré jusqu'à la soumission finale.
* **[landed-cost-calculator-excel](https://github.com/HyVoid/landed-cost-calculator-excel)** — Pourquoi le coût à l'arrivée par SKU issu de mon ERP diffère-t-il de celui de l'API du transitaire, et quelle méthode d'affectation provoque l'érosion de la marge ? — Rapprochez et comparez les calculs de coût à l'arrivée à un modèle de praticien afin de relier les écarts d'affectation des coûts par SKU à leur logique comptable sous-jacente.
* **[hoa-budget-accounting-spreadsheet](https://github.com/HyVoid/hoa-budget-accounting-spreadsheet)** — Les fonds de notre association de copropriétaires sont-ils correctement affectés et respectons-nous notre budget ? — Catégorisez les transactions et générez des rapports d'écart budgétaire pour la gestion financière de l'association.
* **[homeowner-management-excel-template](https://github.com/HyVoid/homeowner-management-excel-template)** — Quel est mon coût total réel de propriété, y compris toutes les dépenses cachées ? — Calculez le coût total de propriété en intégrant le crédit immobilier, les charges, l'entretien et les améliorations d'investissement.
* **[crew-scheduling-excel-template](https://github.com/HyVoid/crew-scheduling-excel-template)** — Mon équipe actuelle pourra-t-elle absorber la charge de travail lorsque les dates de projet se décalent et se chevauchent ? — Prévisionnez la demande d'équipes et les écarts d'effectif sur un horizon de 12 mois pour appuyer les décisions de planification de la capacité.
* **[Schedule-Driven-Construction-Cash-Curve-Workbook](https://github.com/HyVoid/Schedule-Driven-Construction-Cash-Curve-Workbook)** — Comment la trésorerie et l'avancement planifié de mon projet de construction vont-ils évoluer dans le temps en fonction du planning ? — Transformez les plannings de construction et les registres de postes de soumission en courbes en S automatisées, prévisions de coûts échelonnées dans le temps et suivis d'avancement planifié.
* **[shop-floor-mrp-kanban-excel-system](https://github.com/HyVoid/shop-floor-mrp-kanban-excel-system)** — Notre planning de soudage par équipe fonctionnera-t-il sans ruptures de matières arrêtant la ligne ? — Calculez les besoins matières pilotés par la nomenclature et les signaux de réapprovisionnement Kanban à partir des plannings de production par équipe.

---

## Licence

Sous licence **Apache License 2.0**.

Tu es libre d'utiliser, de modifier et de distribuer ce cadre logiciel pour un usage organisationnel commercial et privé conformément aux termes de la licence Apache 2.0.

Copyright © 2026.

