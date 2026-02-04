# Guide de correction - Demande OS 2026

## Probleme

Le fichier 2026 a des conflits de fusion (merge conflicts) en co-edition (SharePoint/OneDrive) que le fichier 2025 n'avait pas.

## Cause racine

Le refactoring 2026 a introduit un pattern **Unprotect -> Action -> Protect** dans chaque fonction utilitaire (`modUI`, `modData`, `modSecurity`). En co-edition, 2 utilisateurs actifs simultanement declenchent ces cycles en parallele, ce qui cree des conflits de fusion.

Les 3 mecanismes responsables :

1. **modUI.bas** (`CreateValidatedDropdown`) : chaque liste deroulante fait Unprotect WB -> Unprotect Sheet -> Ecrit TempLists -> Cree Named Range -> Protect
2. **modSecurity.bas** (`g_SecurityDone` + `modFilters.EnforceAllowedFilterDropdowns`) : applique filtres par colonne avec flag global et cycles Unprotect/Protect supplementaires
3. **Feuil1.cls** (`g_ActivateDone` + header shields) : setup complexe a chaque activation avec cycles Protect/Unprotect

En 2025, les listes etaient creees avec `Formula1:="Oui,Non"` ou `CreateDropdown` inline, sans toucher a la protection.

---

## Objectif

Garder l'architecture modulaire 2026, en corrigeant seulement les 4 modules qui causent les conflits. Les 2 fonctionnalites 2026 (Reset Filtre + Mode Consultation) restent intactes.

---

## APPROCHE RECOMMANDEE : Correction chirurgicale (4 modules)

Au lieu de tout supprimer et reimporter du 2025, on corrige **uniquement les 4 fichiers responsables**. Tous les autres modules (modData, modDocument, modEmail, modConsultation, modFilters, etc.) restent INCHANGES.

### Principe de la correction

**Pourquoi ca marchait en 2025 ?** Parce que les feuilles etaient protegees avec `UserInterfaceOnly:=True`. Ce parametre permet au VBA de modifier les cellules, validations et filtres SANS avoir besoin de deproteger/reproteger. Le code 2026 fait ces cycles Unprotect/Protect inutilement.

### Les 4 fichiers a remplacer

Les fichiers corriges sont fournis dans le depot :

| Fichier | Ce qui change | Fichier corrige |
|---|---|---|
| `modUI.bas` | Suppression de TOUS les Unprotect/Protect dans CreateValidatedDropdown + utilisation de Formula1 directe quand la liste < 255 chars | `FIX_modUI.bas` |
| `modSecurity.bas` | Suppression de g_SecurityDone + suppression de l'appel a EnforceAllowedFilterDropdowns + simplification de EnforceSecurity | `FIX_modSecurity.bas` |
| `Feuil1.cls` | Suppression de g_ActivateDone + header shields + Worksheet_Activate simplifie (Formula1:="Oui,Non" direct) | `FIX_Feuil1.cls` |
| `ThisWorkbook.cls` | Suppression du scroll auto dans Workbook_Open + suppression du Me.Save auto dans BeforeClose | `FIX_ThisWorkbook.cls` |

### Modules qui NE CHANGENT PAS

- `modData.bas` - inchange
- `modDocument.bas` - inchange
- `modEmail.bas` - inchange
- `modConsultation.bas` - inchange
- `modFilters.bas` - inchange (les fonctions CanUseAnyFilters/CanSort qu'il appelle sont gardees dans modSecurity)
- `modReparation.bas` - inchange
- `modDashboard.bas` - inchange
- `modDateNotif.bas` - inchange
- `modResetAnnuel.bas` - inchange
- `Module1.bas` a `Module4.bas` (2026) - inchanges

---

## Procedure dans l'editeur VBA

1. Ouvrir le fichier 2026 -> **Alt+F11**
2. Double-clic sur **ThisWorkbook** -> Remplacer TOUT le code par le contenu de `FIX_ThisWorkbook.cls`
3. Double-clic sur **Feuil1** (Demande OS) -> Remplacer TOUT le code par le contenu de `FIX_Feuil1.cls`
4. Double-clic sur **modSecurity** -> Remplacer TOUT le code par le contenu de `FIX_modSecurity.bas`
5. Double-clic sur **modUI** -> Remplacer TOUT le code par le contenu de `FIX_modUI.bas`
6. **Ctrl+S** pour sauvegarder
7. Fermer et rouvrir le fichier pour tester

---

## Resume des corrections

| Correction | Pourquoi |
|---|---|
| Suppression Unprotect/Protect dans modUI | `UserInterfaceOnly:=True` rend ces cycles inutiles. 2 users faisant Unprotect/Protect en meme temps = conflit |
| Suppression g_SecurityDone | Flag qui empechait la re-initialisation et desynchronisait les sessions |
| Suppression g_ActivateDone | Meme probleme que g_SecurityDone |
| Suppression header shields (Shapes) | Manipulation de Shapes sur l'en-tete qui entrait en conflit entre sessions |
| Suppression EnforceAllowedFilterDropdowns au demarrage | Cette fonction faisait un cycle Unprotect/Protect par colonne (31 colonnes = 31 cycles) |
| Suppression scroll auto + Me.Save | En co-edition, le scroll et la sauvegarde sont geres par SharePoint/OneDrive |
| Formula1:="Oui,Non" direct | Au lieu de passer par un Named Range VL_OuiNon qui necessite TempLists |
