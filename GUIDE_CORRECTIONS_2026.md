# Guide de correction - Demande OS 2026

## Probleme

Le fichier 2026 a des conflits de fusion (merge conflicts) en co-edition (SharePoint/OneDrive) que le fichier 2025 n'avait pas.

## Cause racine

Le refactoring 2026 a introduit un pattern **Unprotect → Action → Protect** dans chaque fonction utilitaire (`modUI`, `modData`, `modSecurity`). En co-edition, 2 utilisateurs actifs simultanement declenchent ces cycles en parallele, ce qui cree des conflits de fusion.

Les 3 mecanismes responsables :

1. **modUI.bas** (`CreateValidatedDropdown`) : chaque liste deroulante fait Unprotect WB → Unprotect Sheet → Ecrit TempLists → Cree Named Range → Protect
2. **modSecurity.bas** (`g_SecurityDone` + filter enforcement) : applique filtres par colonne avec flag global
3. **Feuil1.cls** : appelle modUI/modData/modFilters avec cycles Protect/Unprotect constants

En 2025, les listes etaient creees avec `Formula1:="Oui,Non"` ou `CreateDropdown` inline, sans toucher a la protection.

---

## Objectif

Revenir au fonctionnement 2025 dans le fichier 2026, en gardant :
- Le **bouton Reset Filtre** (`modFilters.bas`)
- Le **mode Consultation** (`modConsultation.bas`)

---

## Plan d'action detaille

### A. MODULES A SUPPRIMER (8 modules)

| Module | Raison |
|---|---|
| `modUI.bas` | Moteur centralise de validation/TempLists - COUPABLE PRINCIPAL |
| `modData.bas` | Remplace par Module1 de 2025 |
| `modDocument.bas` | Remplace par Module17 de 2025 |
| `modEmail.bas` | Remplace par Module9 de 2025 |
| `modResetAnnuel.bas` | N'existait pas en 2025 |
| `modReparation.bas` | N'existait pas en 2025 |
| `modDashboard.bas` | N'existait pas en 2025 |
| `modDateNotif.bas` | N'existait pas en 2025 |

### B. MODULES A CONSERVER

| Module | Action |
|---|---|
| `modConsultation.bas` | GARDER TEL QUEL |
| `modFilters.bas` | GARDER mais remplacer `modSecurity.CanUseAnyFilters()` et `modSecurity.CanSort()` par `True` |

### C. MODULES A REMPLACER (3 fichiers)

#### 1. ThisWorkbook.cls

Remplacer par la version 2025 + ajout nettoyage consultation :

```vb
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    modSecurity.EnforceSecurity
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    modSecurity.RehideIfUnauthorized Sh
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    On Error Resume Next
    modSecurity.HandleNewSheet Sh
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modConsultation.SupprimerVuesTemporaires
    modSecurity.CancelAutoRehide
End Sub
```

#### 2. Feuil1.cls

Remplacer par la version 2025 complete. Seule adaptation :
- Chemin dossier dans `Worksheet_BeforeRightClick` : `...\2026\020 Suivi plan de travaux 2026`
- Appel `ProcessusCompletPourLigne` reste sur Module17

#### 3. modSecurity.bas

Revenir a la version 2025 avec 2 ajouts :
- `AllowedVisibleNames` : ajouter `"Tableau de Bord"` si necessaire
- `IsAllowed` : ajouter detection `FILTRE_PERSO` pour le mode consultation :
  ```vb
  If Left$(sheetName, 12) = "FILTRE_PERSO" Then
      IsAllowed = True
      Exit Function
  End If
  ```

### D. MODULES 2025 A REIMPORTER

| Module | Contenu |
|---|---|
| Module1 | GetComptesMatches, GetImputationMatches, GetCategorieMatches, GetUFSMatches, etc. |
| Module2 | TestCorrespondance |
| Module3 | ClearAllValidations |
| Module4 | ApplyAllDropdownsFixed |
| Module5 | ClearCompteValidation + ClearCompteIPCat |
| Module6 | CreateDropdownFromRange |
| Module7 | CreateDropdownRange |
| Module8 | ApplyDropdownToColumns + ApplyDropdownSorted |
| Module9 | PrepareEmail (version sans ConfigEmails) |
| Module10 | FormatLignesProfessionnel |
| Module11 | FormatTablePersonnalise |
| Module12 | BtnAjoutEntreprise_Click + BtnAjoutCategorie_Click |
| Module13 | RemplirEtFormaterDossiers |
| Module14 | ActualiserReferents |
| Module15 | ApplyCodeProjetDropdown |
| Module16 | UpdateValidationColumnW |
| Module17 | ProcessusCompletPourLigne |

### E. FEUILLES CACHEES

- `TempLists` : SUPPRIMER
- `ConfigEmails` : SUPPRIMER (emails hardcodes dans Module9)
- `ErrorLog` / `Sys` : Peuvent rester

### F. MODULES 2026 A SUPPRIMER

Les Module1 a Module4 actuels du 2026 (NettoyerLeVide, NettoyerNomsValidList, PACK_CLEAN, Refresh_CodesProjets) sont des utilitaires specifiques au refactoring 2026. Ils peuvent etre supprimes car ils dependent de modUI/modData.

---

## Procedure dans l'editeur VBA

1. Ouvrir le fichier 2026 → Alt+F11
2. Supprimer les 8 modules listes en A
3. Supprimer les 4 Module1-4 du 2026
4. Importer les 17 modules du 2025
5. Remplacer le code de ThisWorkbook, Feuil1, modSecurity
6. Adapter modFilters (remplacer CanUseAnyFilters/CanSort par True)
7. Supprimer les feuilles TempLists et ConfigEmails
8. Sauvegarder et tester en co-edition
