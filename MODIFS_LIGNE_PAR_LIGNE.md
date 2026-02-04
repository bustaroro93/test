# Modifications ligne par ligne - Fichier 2026

Ouvre l'editeur VBA (Alt+F11) et fais ces modifications dans l'ordre.
Chaque modification est numerotee. Utilise Ctrl+H (Rechercher/Remplacer) pour aller vite.

---

## MODULE 1 : ThisWorkbook (4 modifs)

### MODIF 1.1 - Supprimer le corps de Workbook_Open (scroll auto)

**CHERCHE** tout le bloc `Workbook_Open` et remplace-le :

SUPPRIMER :
```vb
Private Sub Workbook_Open()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' 1. SECURITE AU DEMARRAGE
    On Error Resume Next
    modSecurity.EnforceSecurity
    On Error GoTo 0

    ' 2. OUVERTURE SUR "DEMANDE OS"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Demande OS")
    ws.Activate

    ws.ScrollArea = "A1:AE1500"

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    If lastRow < 2 Then lastRow = 2

    Application.EnableEvents = False
    Application.GoTo Reference:=ws.Cells(lastRow, 1), Scroll:=True
    Application.EnableEvents = True

    If ActiveWindow.ScrollRow > 10 Then
        ActiveWindow.ScrollRow = ActiveWindow.ScrollRow - 10
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```

REMPLACER PAR :
```vb
Private Sub Workbook_Open()
    On Error Resume Next
    modSecurity.EnforceSecurity
End Sub
```

**Pourquoi** : Le scroll auto + les 3 lignes Application.* causent des conflits quand 2 users ouvrent en meme temps.

---

### MODIF 1.2 - Simplifier Workbook_BeforeClose

SUPPRIMER :
```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 1. NETTOYAGE
    On Error Resume Next
    modConsultation.SupprimerVuesTemporaires
    On Error GoTo 0

    ' 2. SAUVEGARDE SECURISEE
    If Not Me.Saved Then
        On Error GoTo ErrSave
        Me.Save
    End If
    Exit Sub

ErrSave:
    If MsgBox("Impossible de sauvegarder automatiquement." & vbCrLf & _
              "Voulez-vous fermer quand meme ?", vbYesNo + vbExclamation) = vbNo Then
        Cancel = True
    End If
End Sub
```

REMPLACER PAR :
```vb
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modConsultation.SupprimerVuesTemporaires
    modSecurity.CancelAutoRehide
End Sub
```

**Pourquoi** : `Me.Save` en co-edition cree des conflits. SharePoint/OneDrive gere la sauvegarde tout seul.

---

## MODULE 2 : modSecurity (5 modifs)

### MODIF 2.1 - Supprimer g_SecurityDone

CHERCHE la ligne :
```vb
Private g_SecurityDone As Boolean
```
SUPPRIMER cette ligne.

---

### MODIF 2.2 - Supprimer le bloc Enum + Constantes colonnes

SUPPRIMER tout ce bloc (lignes 11-21) :
```vb
Public Enum eFilterLevel
    flNone = 0
    flLimited = 1
    flDTM = 2
    flAdmin = 3
End Enum

' Colonnes filtrables
Private Const COL_REFERENT As Long = 2          ' B
Private Const COL_VALIDATION_DTM As Long = 25   ' Y  <<< change si besoin
Private Const COL_TRAVAUX_TERMINES As Long = 29 ' AC
```

---

### MODIF 2.3 - Simplifier EnforceSecurity

SUPPRIMER tout le Sub `EnforceSecurity()` actuel et REMPLACER PAR :

```vb
Public Sub EnforceSecurity()
    AllowedVisibleNames = Array("Demande OS", "Liste des UFS pour les techs", "Tableau de Bord")

    Application.ScreenUpdating = False
    On Error Resume Next

    ThisWorkbook.Unprotect Password:=WB_PWD

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsAllowed(ws.Name) Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' S'assurer que TempLists existe
    Dim wsTemp As Worksheet
    Set wsTemp = Nothing
    On Error Resume Next
    Set wsTemp = ThisWorkbook.Worksheets("TempLists")
    On Error GoTo 0
    If wsTemp Is Nothing Then
        On Error Resume Next
        Set wsTemp = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Not wsTemp Is Nothing Then
            wsTemp.Name = "TempLists"
            wsTemp.Visible = xlSheetVeryHidden
        End If
        On Error GoTo 0
    End If

    ' Protection Demande OS - SIMPLE
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Demande OS")
    On Error GoTo 0

    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Unprotect Password:="cidalianeves"
        If Not ws.AutoFilterMode Then ws.Range("A1:AE1").AutoFilter
        If ws.FilterMode Then ws.ShowAllData

        ws.Protect Password:="cidalianeves", _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                   AllowFiltering:=True, AllowSorting:=True, _
                   UserInterfaceOnly:=True
        On Error GoTo 0
    End If

    ThisWorkbook.Protect Password:=WB_PWD, Structure:=True, Windows:=False
    Application.ScreenUpdating = True
End Sub
```

**Ce qui change** :
- Plus de `If g_SecurityDone Then Exit Sub` au debut
- Plus de `modFilters.EnforceAllowedFilterDropdowns ws` (faisait 31 cycles Unprotect/Protect)
- Plus de variables `allowF`/`allowS` (on met `True` directement)
- Plus de `g_SecurityDone = True` a la fin
- Ajout creation TempLists pendant que la structure est deprotegee

---

### MODIF 2.4 - Simplifier AdminHide

CHERCHE :
```vb
Public Sub AdminHide()
    CancelAutoRehide
    g_AdminMode = False
    g_SecurityDone = False
    EnforceSecurity
    MsgBox "Mode Admin desactive : feuilles sensibles remasquees.", vbInformation
End Sub
```

REMPLACER PAR :
```vb
Public Sub AdminHide()
    CancelAutoRehide
    g_AdminMode = False
    EnforceSecurity
    MsgBox "Mode Admin desactive : feuilles sensibles remasquees.", vbInformation
End Sub
```

(On enleve juste la ligne `g_SecurityDone = False`)

---

### MODIF 2.5 - Supprimer CurrentFilterLevel

SUPPRIMER tout ce bloc (c'est mort code, plus appele nulle part) :
```vb
Public Function CurrentFilterLevel() As eFilterLevel
    If g_AdminMode Then
        CurrentFilterLevel = flAdmin
        Exit Function
    End If

    Dim u As String
    u = UCase$(Environ$("USERNAME"))

    Select Case u
        Case "R.LETELLIER", "F.WEINTZEM"
            CurrentFilterLevel = flAdmin
        Case "JC.HONART"
            CurrentFilterLevel = flDTM
        Case Else
            CurrentFilterLevel = flLimited
    End Select
End Function
```

**Note** : Les fonctions `CanUseAnyFilters`, `CanSort`, `CanFilterColumn`, `CanUseFilters` sont GARDEES telles quelles (elles retournent True et sont appelees par modFilters).

---

## MODULE 3 : Feuil1 (Demande OS) (5 modifs)

### MODIF 3.1 - Supprimer les constantes header shield et le flag

SUPPRIMER ces 4 lignes (vers le haut du module) :
```vb
Private Const HEADER_ADDR As String = "A1:AE1"
Private Const SHIELD_NAME As String = "shHeaderShield"
Private Const SHIELD_PREFIX As String = "shHeaderShieldCell_"
Private Const FILTER_GAP_PTS As Double = 22#
```

Et SUPPRIMER aussi :
```vb
Private g_ActivateDone As Boolean
```

**Note** : Les constantes UF (UF_VENTILE, PCT_8191, etc.) sont GARDEES.

---

### MODIF 3.2 - Supprimer les 4 fonctions header shield

SUPPRIMER ces 4 Subs/Functions EN ENTIER :

```vb
Private Sub EnsureHeaderRowLocked()
    Me.Range(HEADER_ADDR).Locked = True
End Sub
```

```vb
Private Sub DeleteHeaderShields()
    On Error Resume Next
    Dim i As Long, shp As Shape
    For i = Me.Shapes.count To 1 Step -1
        Set shp = Me.Shapes(i)
        If shp.Name = SHIELD_NAME Or Left$(shp.Name, Len(SHIELD_PREFIX)) = SHIELD_PREFIX Then
            shp.Delete
        End If
    Next i
    On Error GoTo 0
End Sub
```

```vb
Private Sub BringRealButtonsToFront()
    Dim s As Shape
    For Each s In Me.Shapes
        If Len(s.OnAction) > 0 Then
            If s.Name <> SHIELD_NAME And Left$(s.Name, Len(SHIELD_PREFIX)) <> SHIELD_PREFIX Then
                s.ZOrder msoBringToFront
            End If
        End If
    Next s
End Sub
```

```vb
Private Sub EnsureHeaderShield(ByVal allowFilters As Boolean)
    DeleteHeaderShields
    BringRealButtonsToFront
End Sub
```

---

### MODIF 3.3 - Simplifier Worksheet_Activate (LE PLUS IMPORTANT)

SUPPRIMER tout le Sub `Worksheet_Activate` actuel et REMPLACER PAR :

```vb
Private Sub Worksheet_Activate()
    On Error Resume Next

    ' Validation Oui/Non DIRECTE
    With Me.Range("AC2:AC1500").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="Oui,Non"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    Me.ScrollArea = "A1:AE1500"

    On Error GoTo 0
End Sub
```

**Ce qui est supprime** :
- `If g_ActivateDone Then Exit Sub`
- `Me.Unprotect Password:="cidalianeves"`
- `allowF = modSecurity.CanUseAnyFilters()` / `allowS = modSecurity.CanSort()`
- `EnsureHeaderRowLocked`
- `modFilters.EnforceAllowedFilterDropdowns Me`
- `EnsureHeaderShield True`
- Le bloc `CleanUp:` avec `Me.Protect`
- `g_ActivateDone = True`
- Le `Formula1:="=VL_OuiNon"` remplace par `Formula1:="Oui,Non"`

**Pourquoi** : La feuille est DEJA protegee par `EnforceSecurity` avec `UserInterfaceOnly:=True`. Pas besoin de re-deproteger/reproteger ici. Et `Formula1:="Oui,Non"` evite de passer par TempLists + Named Range.

---

### MODIF 3.4 - Simplifier Worksheet_SelectionChange (blocage ligne 1)

CHERCHE :
```vb
    ' Bloquer acces a la ligne 1 pour les non-autorises (clavier)
    If Target.row = 1 Then
        If Not modSecurity.CanFilterColumn(Target.Column) Then
            Application.EnableEvents = False
            Me.Cells(2, Target.Column).Select
            GoTo ExitHere
        Else
            Exit Sub
        End If
    End If
```

SUPPRIMER ce bloc entier (8 lignes). On garde le reste du Sub tel quel.

**Pourquoi** : `CanFilterColumn` retourne toujours True donc ce bloc ne fait jamais rien. Et il appelle modSecurity inutilement.

---

### MODIF 3.5 - Supprimer CalculerMontantTTC (sub morte)

SUPPRIMER ce bloc entier (cette sub n'est appelee nulle part dans Feuil1, le calcul TTC est fait inline) :

```vb
Private Sub CalculerMontantTTC(ByVal rowNumber As Long)
    On Error Resume Next
    Dim montantHT As Double, tvaRate As Double
    montantHT = CDbl(Me.Cells(rowNumber, "R").Value)
    tvaRate = CDbl(Me.Cells(rowNumber, "S").Value)

    If montantHT > 0 And tvaRate >= 0 Then
        Me.Cells(rowNumber, "T").Value = montantHT * (1 + tvaRate)
    Else
        Me.Cells(rowNumber, "T").ClearContents
    End If
    On Error GoTo 0
End Sub
```

---

## MODULE 4 : modUI (3 modifs)

### MODIF 4.1 - Supprimer Unprotect/Protect dans CreateValidatedDropdown

Dans le Sub `CreateValidatedDropdown`, CHERCHE ces 4 lignes (vers le debut) :
```vb
    Application.ScreenUpdating = False

    ' Deverrouillage
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=WB_PWD
    targetCell.Parent.Unprotect Password:=SHEET_PWD
    On Error GoTo ErrorHandler
```

REMPLACER PAR :
```vb
    Application.ScreenUpdating = False
```

(On supprime les 4 lignes d'Unprotect)

---

Ensuite dans le MEME Sub, CHERCHE le bloc `ReProtectAndExit:` (vers la fin) :
```vb
ReProtectAndExit:
    ' >>> FIX : Separer allowF et allowS
    Dim allowF As Boolean, allowS As Boolean
    allowF = (targetCell.Parent.Name = "Demande OS" And modSecurity.CanUseAnyFilters())
    allowS = (targetCell.Parent.Name = "Demande OS" And modSecurity.CanSort())

    targetCell.Parent.Protect Password:=SHEET_PWD, _
                              DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                              AllowFiltering:=allowF, AllowSorting:=allowS, _
                              UserInterfaceOnly:=True

    ThisWorkbook.Protect Password:=WB_PWD, Structure:=True, Windows:=False

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Dim allowF2 As Boolean, allowS2 As Boolean
    allowF2 = (targetCell.Parent.Name = "Demande OS" And modSecurity.CanUseAnyFilters())
    allowS2 = (targetCell.Parent.Name = "Demande OS" And modSecurity.CanSort())

    targetCell.Parent.Protect Password:=SHEET_PWD, _
                              DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                              AllowFiltering:=allowF2, AllowSorting:=allowS2, _
                              UserInterfaceOnly:=True
    ThisWorkbook.Protect Password:=WB_PWD, Structure:=True, Windows:=False
    Application.ScreenUpdating = True
    Debug.Print "Erreur CreateValidatedDropdown : " & Err.Description
End Sub
```

REMPLACER PAR :
```vb
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print "Erreur CreateValidatedDropdown : " & Err.Description
End Sub
```

Et le `GoTo ReProtectAndExit` dans le garde-fou "si aucune valeur" (ligne ~73), remplacer par :
```vb
        GoTo ReProtectAndExit
```
devient :
```vb
        Application.ScreenUpdating = True
        Exit Sub
```

---

### MODIF 4.2 - Supprimer Unprotect/Protect dans ClearValidation

CHERCHE :
```vb
Public Sub ClearValidation(targetCell As Range)
    On Error Resume Next
    targetCell.Parent.Unprotect Password:=SHEET_PWD
    targetCell.Validation.Delete
    targetCell.ClearContents
    Dim allowF As Boolean, allowS As Boolean
    allowF = (targetCell.Parent.Name = "Demande OS" And modSecurity.CanUseAnyFilters())
    allowS = (targetCell.Parent.Name = "Demande OS" And modSecurity.CanSort())

    targetCell.Parent.Protect Password:=SHEET_PWD, _
                              DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                              AllowFiltering:=allowF, AllowSorting:=allowS, _
                              UserInterfaceOnly:=True
End Sub
```

REMPLACER PAR :
```vb
Public Sub ClearValidation(targetCell As Range)
    On Error Resume Next
    targetCell.Validation.Delete
    targetCell.ClearContents
End Sub
```

---

### MODIF 4.3 - Supprimer Unprotect/Protect dans GetOrCreateTempSheet

CHERCHE :
```vb
Private Function GetOrCreateTempSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateTempSheet = ThisWorkbook.Worksheets(TEMP_SHEET_NAME)

    If GetOrCreateTempSheet Is Nothing Then
        ThisWorkbook.Unprotect Password:=WB_PWD
        Set GetOrCreateTempSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        GetOrCreateTempSheet.Name = TEMP_SHEET_NAME
        GetOrCreateTempSheet.Visible = xlSheetVeryHidden
        ThisWorkbook.Protect Password:=WB_PWD, Structure:=True, Windows:=False
    End If
    On Error GoTo 0
End Function
```

REMPLACER PAR :
```vb
Private Function GetOrCreateTempSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateTempSheet = ThisWorkbook.Worksheets(TEMP_SHEET_NAME)

    If GetOrCreateTempSheet Is Nothing Then
        Set GetOrCreateTempSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Not GetOrCreateTempSheet Is Nothing Then
            GetOrCreateTempSheet.Name = TEMP_SHEET_NAME
            GetOrCreateTempSheet.Visible = xlSheetVeryHidden
        End If
    End If
    On Error GoTo 0
End Function
```

(On enleve `Unprotect`/`Protect`. La feuille TempLists est maintenant creee dans `EnforceSecurity` au demarrage quand le classeur est deja deprotege.)

---

### MODIF 4.4 - Supprimer Unprotect/Protect dans CleanupTempLists

CHERCHE :
```vb
    If Not ws Is Nothing Then
        ThisWorkbook.Unprotect Password:=WB_PWD
        ws.Cells.Clear
        ThisWorkbook.Protect Password:=WB_PWD, Structure:=True, Windows:=False
    End If
```

REMPLACER PAR :
```vb
    If Not ws Is Nothing Then
        ws.Cells.Clear
    End If
```

---

## RESUME

| # | Module | Ce que tu fais | Temps |
|---|---|---|---|
| 1.1 | ThisWorkbook | Simplifier Workbook_Open (supprimer scroll) | 30 sec |
| 1.2 | ThisWorkbook | Simplifier BeforeClose (supprimer Me.Save) | 30 sec |
| 2.1 | modSecurity | Supprimer `g_SecurityDone` | 5 sec |
| 2.2 | modSecurity | Supprimer Enum + constantes colonnes | 10 sec |
| 2.3 | modSecurity | Remplacer EnforceSecurity | 1 min |
| 2.4 | modSecurity | Supprimer `g_SecurityDone = False` dans AdminHide | 5 sec |
| 2.5 | modSecurity | Supprimer CurrentFilterLevel | 10 sec |
| 3.1 | Feuil1 | Supprimer constantes shield + g_ActivateDone | 10 sec |
| 3.2 | Feuil1 | Supprimer les 4 subs header shield | 20 sec |
| 3.3 | Feuil1 | Remplacer Worksheet_Activate | 1 min |
| 3.4 | Feuil1 | Supprimer blocage ligne 1 dans SelectionChange | 10 sec |
| 3.5 | Feuil1 | Supprimer CalculerMontantTTC (sub morte) | 10 sec |
| 4.1 | modUI | Supprimer Unprotect/Protect dans CreateValidatedDropdown | 1 min |
| 4.2 | modUI | Supprimer Unprotect/Protect dans ClearValidation | 15 sec |
| 4.3 | modUI | Supprimer Unprotect/Protect dans GetOrCreateTempSheet | 15 sec |
| 4.4 | modUI | Supprimer Unprotect/Protect dans CleanupTempLists | 10 sec |

**Total : 16 modifications dans 4 modules. Rien a ajouter, rien a importer.**
