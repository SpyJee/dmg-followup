# Production Follow-Up System — Brief for Claude Code
**File:** `Suivi_Production_SOURCE.xlsx`  
**Purpose:** Track manufacturing orders from receipt to shipment, auto-calculating operation deadlines backwards from the required delivery date, accounting for Quebec holidays and construction vacations.

---

## Overview

The workbook is the central database for production planning. It links **orders → part recipes → operation schedules → live status tracking**. All deadline calculations use working days only, excluding a defined list of Quebec statutory holidays and construction vacation weeks.

---

## Sheets Summary

### 1. `Commande` — Orders Table
The master list of active orders.

| Column | Description |
|---|---|
| `No Bon Commande` | Purchase order number (e.g. `5002340270`) |
| `No Item` | Line item within the PO (e.g. `00001`) |
| `CommandeItem` | Unique key = `No Bon Commande` + `-` + `No Item` (e.g. `5002340270-00001`) |
| `Client` | Customer name (dropdown: Airbus Atlantique, Airbus Canada, Satair, Bombardier, Placeteco, Hutchison, Apex Précision, Bombardier Red Oak) |
| `No Pièce` | Part number — used to look up the recipe |
| `Date Requise` | Required delivery date |
| `Quantité` | Quantity ordered |
| `No Bon Travail` | Work order number |
| `Statut - Opération en cours` | Current active operation name (computed from Suivi) |
| `Détail Opération en cours` | Detail of current operation (e.g. machine number) |
| `Date Limite OP en cours` | Deadline for the current active operation |
| `Commentaire` | Free-text comment |

**Current data:** 7 order lines across 3 POs, all for Airbus Canada + 2 test entries (Bombardier).

---

### 2. `Recette` — Part Recipes
Defines the sequence of operations for each part number.

| Column | Description |
|---|---|
| `No Pièce` | Part number (primary key, links to Commande) |
| `Template Référence` | Which template was used (Standard, Custom, Plastique, etc.) |
| `Opération 1–16` | Operation name for each step |
| `Détail OP 1–16` | Detail / sub-info (e.g. machine number, subcontractor) |
| `Délais OP 1–16` | Lead time in **working days** for that operation |
| `Nb Ops` | Total number of active operations (ignores empty slots) |

**Important:** A recipe row with `Template Référence` ending in `*` means it was manually overridden from the template.  
**Special row:** The last row `"Garder cette Ligne!!!!"` is a placeholder/template anchor — do NOT delete it.

**Current parts:**
| No Pièce | Template | Nb Ops |
|---|---|---|
| C01645411-N0003 | Standard | 8 |
| C00814156-101-01 | Standard+Assy | 10 |
| C00814121-101-01 | Plastique | 7 |
| C00814159-101-01 | Plastique | 7 |
| C00814125-101-01 | Standard | 8 |
| 1234 | Standard | 8 |
| 2450 | Custom | 13 |

---

### 3. `Template Recette` — Operation Templates
Reusable templates that pre-fill a recipe's operations and default lead times.

| Template ID | Description |
|---|---|
| `Custom` | Blank — fully manual |
| `Standard` | 8 ops: Commande Matière (20d) → Coupe (3d) → Usinage (5d) → Ébavurage (3d) → Sous-traitance/Finition TNM (12d) → Inspection (2d) → Identification (0d) → Expédition (2d) |
| `Standard*` | Same as Standard (manual override version) |
| `Standard+Assy` | Standard + Assemblage (4d) + extra Inspection (2d) = 10 ops |
| `Standard+Assy*` | Same as Standard+Assy (manual override version) |
| `Plastique` | 7 ops: Commande Matière (25d) → Coupe (3d) → Usinage (5d) → Ébavurage (3d) → Inspection (2d) → Identification (0d) → Expédition (2d) |
| `Plastique*` | Same as Plastique (manual override version) |

**Note:** `***Info Requise***` in `Détail OP 3` (Usinage) is a placeholder — must be replaced with the actual machine number in the recipe.

---

### 4. `Planification` — Deadline Calculations
Auto-calculated sheet. **Read-only** from a logic standpoint — derived entirely from Commande + Recette.

| Column | Description |
|---|---|
| `CommandeItem` | Links to Commande |
| `No Pièce` | Links to Recette |
| `Date Requise` | Delivery date |
| `OP1 Date Limite` → `OP16 Date Limite` | Deadline for each operation |

**Calculation logic:** Deadlines are computed **backwards** from `Date Requise`. Each operation's deadline = next operation's deadline minus that operation's lead time in working days, excluding holidays.

Example for part C01645411-N0003 (required 2026-08-03):
- OP1 (Commande Matière, 20d): 2026-05-11
- OP2 (Coupe, 3d): 2026-06-09
- OP3 (Usinage, 5d): 2026-06-12
- ...
- OP7 (Identification, 0d): 2026-07-16
- OP8 (Expédition, 2d): 2026-08-03

---

### 5. `Suivi` — Live Tracking Table
The main operational view. Managers manually update operation statuses here.

**Columns A–E:** Basic order info pulled from Commande (CommandeItem, No Pièce, Date Requise, Quantité, No Bon Travail)

**Columns F–I:** Active operation (computed)
| Column | Description |
|---|---|
| `OP Active` | First operation not yet marked "Complété" (e.g. `Op3`, `Op7`) |
| `Description` | Operation type (e.g. Usinage, Inspection) |
| `Détail` | Sub-detail (e.g. Machine #7) |
| `Date Limite` | Deadline from Planification for the active op |

**Columns J–Y:** `Statut OP1` through `Statut OP16` — manually set by the production manager.

**Possible status values:**
- `Requis` — Not started
- `En cours` — In progress
- `Bloqué` — Blocked
- `Complété` — Done
- `N/A` — Not applicable (operation slot unused)
- (blank) — Operation doesn't exist for this part

**Column Z:** `Nb Ops` — number of operations for this part (from Recette)

**Current state snapshot:**
| CommandeItem | OP Active | Description | Deadline |
|---|---|---|---|
| 5002340270-00001 | Op7 | Identification | 2026-07-16 |
| 5002344711-00001 | Op6 | Inspection | 2025-12-01 |
| 5002344711-00002 | Op4 | Ébavurage | 2025-12-04 |
| 5002344711-00003 | Op3 | Usinage (Machine #7) | 2025-11-27 |
| 5002344711-4 | Op6 | Inspection | 2025-12-09 |
| New-1 | Op6 | Inspection | 2027-02-18 |
| 1234-1 | Op3 | Sous-traitance (Heat Treat VAC Aero) | 2026-10-26 |

---

### 6. `Menu déroulant` — Dropdown Lists & Reference Data

| Column | Values |
|---|---|
| `Fériers` | Holiday dates (date objects) |
| `Fériers (description)` | Holiday names |
| `Client` | Airbus Atlantique, Airbus Canada, Satair, Bombardier, Placeteco, Hutchison, Apex Précision, Bombardier Red Oak |
| `Suivi` | Matière à commander, En attente matière, Coupe, Traitement Thermique, En attente programmation, En attente Dessin, Usinage |
| `Statut Étape` | Requis, En cours, Bloqué, Complété |
| `Statut Matière` | À commander, En attente, Stock |
| `Opérations` | Commande Matière, Coupe, Usinage, Ébavurage, Sous-traitance, Assemblage, Inspection, Identification, Expédition |
| `Détails (Opérations)` | Traitement Thermique, Finition (TNM), Finition (Ultraspec), Finition (IVD), Machine #2 through Machine #15 |

**Holidays defined (used in working-day calculations):**

2026: Jan 1, Apr 3 (Good Friday), May 18 (Patriots), Jun 24 (St-Jean), Jul 1 (Canada), Jul 20–31 (Construction vacation), Sep 7 (Labour), Oct 12 (Thanksgiving), Dec 25

2027: Jan 1, Mar 26 (Good Friday), May 24 (Patriots), Jun 24 (St-Jean), Jul 1 (Canada), Jul 19–30 (Construction vacation), Sep 6 (Labour), Oct 11 (Thanksgiving), Dec 25

---

## Key Relationships

```
Commande.No Pièce
    → Recette.No Pièce          (recipe lookup)
    → Recette.Délais OP 1–16    (used by Planification)

Commande.CommandeItem
    → Planification.CommandeItem (deadline lookup)
    → Suivi.CommandeItem         (status tracking)

Template Recette.Template ID
    → Recette.Template Référence (auto-fill operations)

Menu déroulant.Fériers
    → Planification               (holiday exclusions in WORKDAY calculations)
```

---

## Notes for Automation

- **Unique key** for all cross-sheet lookups: `CommandeItem` = `No Bon Commande` + `-` + `No Item`
- **Deadline logic** is backward-chaining WORKDAY: start from `Date Requise`, subtract each op's delay going left (OP16 → OP1)
- **Active operation** = first op index where status ≠ `Complété` and index ≤ `Nb Ops`
- **New orders** need: a row in `Commande`, a row in `Recette` (with valid `No Pièce`), and the `Planification`/`Suivi` sheets will auto-populate via Excel formulas
- **Templates with `*`** suffix indicate the recipe was manually customized after being loaded from the base template
- Dates are stored as Excel date serials; use ISO format `YYYY-MM-DD` when writing back
