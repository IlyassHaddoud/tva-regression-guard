# Prompt : application de contrôle et de delta du fichier TVA (Sage X3)

## Contexte

Sur Sage X3, un développement spécifique calcule la TVA et génère un fichier CSV. Ce fichier contient les règlements, les factures rattachées, les montants et la TVA à déclarer pour chaque règlement.

Quand on modifie le spécifique, on tombe parfois sur des régressions. Je veux un outil qui me permette de contrôler le fichier généré et de comparer deux versions du fichier.

## Objectif

Écris une application Python en ligne de commande avec **deux volets** :

1. **`controle`** : contrôler un seul fichier. C'est la priorité de la première version.
2. **`delta`** : comparer un ancien et un nouveau fichier (rapprochement) pour détecter :
   - les nouveaux règlements,
   - les règlements supprimés,
   - les règlements dont les montants ont changé.

## Format du fichier CSV

- Séparateur `;`, encodage UTF-8 avec BOM (prévoir un repli en cp1252), première ligne = en-têtes.
- Les montants utilisent le point comme séparateur décimal, mais certaines valeurs peuvent contenir des espaces de remplissage ou une virgule : lire toutes les colonnes en texte, nettoyer, puis convertir.
- Ordre d'un millier à quelques dizaines de milliers de lignes.
- Colonnes :

```
N° règlement; N° pièce règlement; Ligne règlement; N° relevé; N° pièce facture;
Référence règlement; Type règlement; Date règlement; Montant règlement; Montant signé;
Date de pointage; Pointage; Compte; Client; Raison sociale; ICE; Banque; Remarque;
N° facture; Identifiant facture; Type facture; Date facture; HT facture; TTC facture;
Montant réglé; Total TVA; TVA à déclaré; Taux de TVA; Type de TVA; Accnum; Déclaré;
Lettrage; TVA NR
```

- Une ligne du fichier correspond à une facture (ou une ligne) rattachée à un règlement. Un règlement a donc plusieurs lignes.
- Un règlement est identifié par `N° pièce règlement`. Une ligne est identifiée par `N° pièce règlement` + `Ligne règlement`.
- `Montant signé` (le total du règlement) est répété sur chaque ligne d'un même règlement.
- Valeurs de `Type de TVA` : `Collectée ventes`, `Déductible achats`, `Déductible immos`, ou vide.
- Valeurs de `Remarque` à connaître : `RAPPROCHE LETTRE`, `IMPAY`, `ETAPE INTERMEDIAIRE NON LETTRE`, `FRAIS BANCAIRE`, `Reclassement multiple`, ainsi que d'autres (`Facturation interne`, `HORS TVA`, `RAPPROCHE NON LETTRE`, `REG ANNULE`, `NON RAPPROCHE`, `DATE DE VALEUR VIDE`).

## Volet 1 : contrôle d'un seul fichier

### Affichage 1 : total de TVA par type

- Afficher le total de `TVA à déclaré` par `Type de TVA`.
- Afficher ensuite la TVA nette : ventes moins achats, en fusionnant `Déductible achats` et `Déductible immos`.
- Tout ça dans **un seul tableau**.

### Affichage 2 : contrôle de la partie vente

- Contrôler chaque type de TVA, en commençant par la partie **vente** (`Collectée ventes`).
- Pour chaque règlement, comparer :
  - le total du règlement : colonne `Montant signé`,
  - avec le détail du règlement : somme de `Montant réglé` regroupée par `N° pièce règlement`.
- Les deux doivent être égaux. Sinon, le règlement est en écart.
- Faire attention à la colonne `Remarque` : on se concentre normalement sur `RAPPROCHE LETTRE`, mais un même règlement peut aussi avoir des lignes dont la remarque est `IMPAY`, `ETAPE INTERMEDIAIRE NON LETTRE`, `FRAIS BANCAIRE` ou `Reclassement multiple`.
- Les frais bancaires sont normalement côté **achat** : pour les récupérer, il faut aussi chercher dans la partie `Déductible achats`.
- **Point d'attention sur `Reclassement multiple`** : ces lignes ont un comportement ambigu. Parfois elles **régularisent** l'écart (le règlement n'est cohérent qu'en les prenant en compte), et parfois elles **le doublent** (les prendre en compte crée un écart qui n'existait pas). Ne les ajoute donc pas systématiquement : compare les deux cas sur les données réelles, cherche ce qui les distingue (montant de la ligne par rapport au total du règlement, lettrage, type de TVA…), propose-moi une règle, et signale clairement les règlements où le doute persiste.
- Les décaissements ont un `Montant signé` négatif, donc comparer avec ce point en tête.
- Quand une règle métier est ambiguë, **pose-moi la question** plutôt que de deviner, et teste tes hypothèses sur les données réelles.

### Ensuite

Faire de même pour la partie **achat**.

## Volet 2 : delta entre deux fichiers

- Comparer règlement par règlement (regroupement par `N° pièce règlement`) : nombre de lignes, somme de `Montant réglé`, somme de `TVA à déclaré`, `Montant signé`.
- Classer chaque règlement : NOUVEAU, SUPPRIME, MODIFIE ou IDENTIQUE.
- Afficher le delta de TVA par type.
- Signaler les écarts de contrôle apparus dans le nouveau fichier par rapport à l'ancien.
- Dans une version suivante : descendre au niveau de la ligne (facture par facture).

## Contraintes d'usage

- Le fichier contrôlé est `file.csv`, placé à la racine du dossier du projet (chemin par défaut).
- Un seul script suffit, avec deux commandes : `controle` et `delta`.
- Python avec pandas, et openpyxl seulement pour l'export optionnel.
- Je travaille sous **Windows / PowerShell** : donne-moi les commandes exactes pour installer et lancer.
- Les bibliothèques s'installent en global, sans environnement virtuel.

## Contraintes d'affichage

- Le résultat s'affiche **dans la console**, sans générer de fichier Excel par défaut. L'export Excel est une option (`-o rapport.xlsx`).
- Les tableaux sont affichés avec **des bordures** (caractères box-drawing), avec :
  - les montants alignés à droite au format français (`1 234 567,89`),
  - le texte aligné à gauche.
- Ne pas répéter la même information dans deux tableaux.
- Afficher d'abord le tableau des règlements **en écart**, puis un second tableau avec **tous les règlements** (avec et sans écart) et une colonne `statut` (`OK` / `ECART`).
- Prévoir une option pour n'afficher que les écarts.
- Forcer l'UTF-8 en sortie pour que les bordures et les accents s'affichent bien sous Windows.

## Livrables attendus

1. Le script complet.
2. Les commandes PowerShell pour l'installer et le lancer.
3. Une courte explication de chaque colonne du tableau de contrôle.
4. La liste des hypothèses métier que tu as dû faire, pour que je les valide.

## Méthode

- Commence par lire un extrait du fichier pour confirmer le format réel (séparateur, décimales, valeurs distinctes de `Remarque` et de `Type de TVA`).
- Construis la version 1 (contrôle de la partie vente) et teste-la sur le fichier réel avant d'aller plus loin.
- Analyse les règlements en écart pour distinguer les vrais écarts de ceux qui viennent d'une règle métier oubliée, et propose-moi les règles à valider.
- Avance par étapes et montre-moi le résultat console après chaque étape.
