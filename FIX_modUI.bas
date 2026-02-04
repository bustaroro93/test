Attribute VB_Name = "modUI"
Option Explicit

' Constantes centralisees
Private Const WB_PWD As String = "C23j0813"
Private Const SHEET_PWD As String = "cidalianeves"
Private Const TEMP_SHEET_NAME As String = "TempLists"

' ==============================================================================
' MOTEUR DE VALIDATION - VERSION CORRIGEE POUR CO-EDITION
' ==============================================================================
' CORRECTION PRINCIPALE : On ne fait PLUS de Unprotect/Protect ici.
' Les feuilles sont protegees avec UserInterfaceOnly:=True, ce qui permet
' au VBA de modifier les validations et cellules SANS deproteger.
' C'est ca qui causait les conflits de fusion en co-edition.
' ==============================================================================
Public Sub CreateValidatedDropdown(ByVal targetCell As Range, _
                                   ByVal sourceData As Variant, _
                                   Optional ByVal listName As String = "", _
                                   Optional ByVal sortData As Boolean = False)
    On Error GoTo ErrorHandler

    ' Securite : Si la source est vide, on ne fait rien
    If IsEmpty(sourceData) Then Exit Sub
    If Not IsArray(sourceData) Then Exit Sub

    ' --- TENTATIVE 1 : Chaine directe Formula1 (pas de feuille cachee) ---
    ' Si la liste tient dans 255 caracteres, on l'utilise directement
    Dim csvList As String
    Dim i As Long
    csvList = ""
    For i = LBound(sourceData) To UBound(sourceData)
        If CStr(sourceData(i)) <> "" Then
            If csvList <> "" Then csvList = csvList & ","
            csvList = csvList & CStr(sourceData(i))
        End If
    Next i

    ' Trier si demande (tri simple)
    If sortData Then
        Dim arr() As String
        Dim parts() As String
        parts = Split(csvList, ",")
        Dim j As Long, tmp As String
        For i = LBound(parts) To UBound(parts) - 1
            For j = i + 1 To UBound(parts)
                If UCase$(parts(j)) < UCase$(parts(i)) Then
                    tmp = parts(i)
                    parts(i) = parts(j)
                    parts(j) = tmp
                End If
            Next j
        Next i
        csvList = Join(parts, ",")
    End If

    If Len(csvList) <= 255 And Len(csvList) > 0 Then
        ' CAS SIMPLE : La liste tient dans Formula1 => pas besoin de TempLists
        On Error Resume Next
        targetCell.Validation.Delete
        targetCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=csvList
        targetCell.Validation.IgnoreBlank = True
        targetCell.Validation.InCellDropdown = True
        On Error GoTo 0
        Exit Sub
    End If

    ' --- TENTATIVE 2 : Liste trop longue, utiliser TempLists + Named Range ---
    ' (Rare : seulement si > 255 caracteres)
    ' On utilise On Error Resume Next au lieu de Unprotect/Protect

    Dim wsTemp As Worksheet
    Set wsTemp = GetOrCreateTempSheet()
    If wsTemp Is Nothing Then
        ' Si on ne peut pas creer la feuille, fallback sur les 255 premiers chars
        On Error Resume Next
        targetCell.Validation.Delete
        targetCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=Left$(csvList, 255)
        On Error GoTo 0
        Exit Sub
    End If

    Dim tempRangeName As String
    If listName <> "" Then
        tempRangeName = "VL_" & CleanNameString(listName)
    Else
        tempRangeName = "VL_Col" & targetCell.Cells(1, 1).Column
    End If

    On Error Resume Next
    ThisWorkbook.Names(tempRangeName).Delete
    On Error GoTo ErrorHandler

    Dim targetCol As Long
    targetCol = GetOrCreateColumnForList(wsTemp, tempRangeName)

    wsTemp.Cells(1, targetCol).Resize(2000, 1).ClearContents
    wsTemp.Cells(1, targetCol).Value = tempRangeName

    Dim r As Long
    r = 2
    For i = LBound(sourceData) To UBound(sourceData)
        If CStr(sourceData(i)) <> "" Then
            wsTemp.Cells(r, targetCol).Value = sourceData(i)
            r = r + 1
        End If
    Next i

    If r <= 2 Then
        targetCell.Validation.Delete
        Exit Sub
    End If

    If sortData And r > 3 Then
        wsTemp.Range(wsTemp.Cells(2, targetCol), wsTemp.Cells(r - 1, targetCol)).Sort _
            Key1:=wsTemp.Cells(2, targetCol), Order1:=xlAscending, Header:=xlNo
    End If

    If r > 2 Then
        ThisWorkbook.Names.Add Name:=tempRangeName, _
                               RefersTo:=wsTemp.Range(wsTemp.Cells(2, targetCol), wsTemp.Cells(r - 1, targetCol))
    End If

    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & tempRangeName
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0

    Exit Sub

ErrorHandler:
    Debug.Print "Erreur CreateValidatedDropdown : " & Err.Description
End Sub

' Nettoie une chaine pour en faire un nom valide
Private Function CleanNameString(s As String) As String
    Dim result As String
    result = s
    result = Replace(result, " ", "_")
    result = Replace(result, "-", "_")
    result = Replace(result, "'", "")
    result = Replace(result, """", "")
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    If Len(result) > 30 Then result = Left(result, 30)
    CleanNameString = result
End Function

' Trouve ou cree une colonne pour une liste donnee
Private Function GetOrCreateColumnForList(ws As Worksheet, listName As String) As Long
    Dim lastCol As Long, col As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1
    For col = 1 To lastCol
        If ws.Cells(1, col).Value = listName Then
            GetOrCreateColumnForList = col
            Exit Function
        End If
    Next col
    GetOrCreateColumnForList = lastCol + 1
End Function

' Obtenir ou creer la feuille TempLists
' CORRECTION : On utilise On Error Resume Next au lieu de Unprotect/Protect
Private Function GetOrCreateTempSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateTempSheet = ThisWorkbook.Worksheets(TEMP_SHEET_NAME)

    If GetOrCreateTempSheet Is Nothing Then
        ' Tenter de creer sans Unprotect (marchera si structure pas protegee)
        Set GetOrCreateTempSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Not GetOrCreateTempSheet Is Nothing Then
            GetOrCreateTempSheet.Name = TEMP_SHEET_NAME
            GetOrCreateTempSheet.Visible = xlSheetVeryHidden
        End If
    End If
    On Error GoTo 0
End Function

' ==============================================================================
' FONCTIONS RELAIS (avec noms stables)
' ==============================================================================
Public Sub ApplyDropdownSorted(sheetName As String, sourceColumn As String, targetCell As Range)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, sourceColumn).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim dataArray As Variant
    dataArray = ws.Range(sourceColumn & "2:" & sourceColumn & lastRow).Value

    Dim values() As Variant
    ReDim values(1 To UBound(dataArray, 1))
    Dim i As Long
    For i = 1 To UBound(dataArray, 1)
        values(i) = dataArray(i, 1)
    Next i

    CreateValidatedDropdown targetCell, values, sheetName, True
End Sub

Public Sub ApplyDropdownToColumns()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Demande OS")

    ' CORRECTION : Utiliser des chaines directes au lieu de Named Ranges
    ' Ces listes sont courtes, elles tiennent dans 255 caracteres
    Dim listeStatuts As Variant: listeStatuts = Array("En attente", "Valid" & ChrW(233), "Refus" & ChrW(233))
    Dim listeSoumis As Variant: listeSoumis = Array("Soumis", "Refus" & ChrW(233))

    CreateValidatedDropdown ws.Range("Y2:Y1500"), listeStatuts, "Statuts", False
    CreateValidatedDropdown ws.Range("W2:W1500"), listeSoumis, "Soumis", False
End Sub

Public Sub ClearValidation(targetCell As Range)
    On Error Resume Next
    ' CORRECTION : Plus de Unprotect/Protect ici
    ' UserInterfaceOnly:=True permet au VBA de modifier
    targetCell.Validation.Delete
    targetCell.ClearContents
End Sub

' ==============================================================================
' FONCTIONS D'AJOUT (BOUTONS) - Inchangees
' ==============================================================================
Public Sub BtnAjoutEntreprise_Click()
    On Error GoTo ErrHandler
    Dim wsMain As Worksheet, ws As Worksheet, wsComptes As Worksheet
    Dim lastRow As Long, lastRowComptes As Long, newEntreprise As Variant

    Set wsMain = ThisWorkbook.Worksheets("Demande OS")
    Set ws = ThisWorkbook.Worksheets("Entreprises")
    Set wsComptes = ThisWorkbook.Worksheets("Comptes")

    newEntreprise = Application.InputBox("Veuillez saisir le nom de la nouvelle entreprise :", "Ajouter une nouvelle entreprise", Type:=2)
    If TypeName(newEntreprise) = "Boolean" Or Trim(CStr(newEntreprise)) = "" Then Exit Sub
    newEntreprise = UCase(Trim(CStr(newEntreprise)))

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If Not ws.Range("A2:A" & lastRow).Find(What:=newEntreprise, LookIn:=xlValues, LookAt:=xlWhole) Is Nothing Then
        MsgBox "Cette entreprise existe d" & ChrW(233) & "j" & ChrW(224) & " !", vbExclamation: Exit Sub
    End If

    Application.ScreenUpdating = False
    ws.Cells(lastRow + 1, "A").Value = newEntreprise
    ws.Range("A2:A" & lastRow + 1).Sort Key1:=ws.Cells(1), Order1:=xlAscending, Header:=xlNo

    lastRowComptes = wsComptes.Cells(wsComptes.Rows.Count, "A").End(xlUp).Row
    Dim dictComptes As Object, i As Long, dataComptes As Variant
    Set dictComptes = CreateObject("Scripting.Dictionary")

    If lastRowComptes >= 2 Then
        dataComptes = wsComptes.Range("A2:C" & lastRowComptes).Value
        For i = 1 To UBound(dataComptes, 1)
            If Not IsEmpty(dataComptes(i, 3)) Then
                If Not dictComptes.Exists(CStr(dataComptes(i, 3))) Then dictComptes.Add CStr(dataComptes(i, 3)), CStr(dataComptes(i, 2))
            End If
        Next i
    End If

    Dim compte As Variant
    For Each compte In dictComptes.Keys
        lastRowComptes = wsComptes.Cells(wsComptes.Rows.Count, "A").End(xlUp).Row
        wsComptes.Cells(lastRowComptes + 1, "A").Value = newEntreprise
        wsComptes.Cells(lastRowComptes + 1, "B").Value = dictComptes(compte)
        wsComptes.Cells(lastRowComptes + 1, "C").Value = compte
    Next compte

    wsComptes.Range("A2:C" & wsComptes.Cells(wsComptes.Rows.Count, "A").End(xlUp).Row).Sort Key1:=wsComptes.Columns(1), Order1:=xlAscending, Header:=xlNo
    Application.ScreenUpdating = True

    ApplyDropdownSorted "Entreprises", "A", wsMain.Range("C2:C1500")
    MsgBox "L'entreprise '" & newEntreprise & "' a " & ChrW(233) & "t" & ChrW(233) & " ajout" & ChrW(233) & "e.", vbInformation
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

Public Sub BtnAjoutCategorie_Click()
    On Error GoTo ErrHandler
    Dim wsCat As Worksheet, wsMain As Worksheet
    Dim lastRowCat As Long, newCategorie As Variant
    Dim entreprise As String, classe As String, compte As String, currentRow As Long

    Set wsMain = ThisWorkbook.Worksheets("Demande OS")
    Set wsCat = ThisWorkbook.Worksheets("CategorieHomogene")

    If ActiveSheet.Name <> wsMain.Name Then wsMain.Activate
    currentRow = ActiveCell.Row

    If currentRow < 2 Or currentRow > 1500 Then
        MsgBox "Placez-vous sur une ligne de donn" & ChrW(233) & "es (2 " & ChrW(224) & " 1500) avant d'ajouter une op" & ChrW(233) & "ration.", vbExclamation
        Exit Sub
    End If

    entreprise = Trim(wsMain.Cells(currentRow, "C").Value)
    classe = Trim(wsMain.Cells(currentRow, "E").Value)
    compte = Trim(wsMain.Cells(currentRow, "G").Value)

    If InStr(1, classe, "Classe 2", vbTextCompare) = 0 Then
        MsgBox "L'ajout de codes op" & ChrW(233) & "ration est r" & ChrW(233) & "serv" & ChrW(233) & " aux dossiers Classe 2.", vbExclamation
        Exit Sub
    End If

    If entreprise = "" Or compte = "" Then
        MsgBox "S" & ChrW(233) & "lectionnez d'abord une entreprise et un compte.", vbExclamation
        Exit Sub
    End If

    newCategorie = Application.InputBox("Nom de la nouvelle op" & ChrW(233) & "ration :", "Ajouter une op" & ChrW(233) & "ration (Classe 2)", Type:=2)
    If TypeName(newCategorie) = "Boolean" Or Trim(CStr(newCategorie)) = "" Then Exit Sub
    newCategorie = UCase(Trim(CStr(newCategorie)))

    lastRowCat = wsCat.Cells(wsCat.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRowCat
        Dim catExist As String, clsExist As String, entExist As String
        catExist = UCase$(Trim$(wsCat.Cells(i, "D").Value))
        clsExist = UCase$(Trim$(wsCat.Cells(i, "B").Value))
        entExist = Trim$(wsCat.Cells(i, "A").Value)

        If catExist = newCategorie Then
            If (clsExist = "CLASSE 2" Or entExist = "*") And _
               (entExist = "*" Or UCase$(entExist) = UCase$(entreprise)) Then
                MsgBox "Ce code op" & ChrW(233) & "ration existe d" & ChrW(233) & "j" & ChrW(224) & " !", vbExclamation
                Exit Sub
            End If
        End If
    Next i

    Dim choix As VbMsgBoxResult
    choix = MsgBox("Cr" & ChrW(233) & "er ce code op" & ChrW(233) & "ration pour :" & vbCrLf & vbCrLf & _
                   "OUI = TOUTES les entreprises (Classe 2)" & vbCrLf & _
                   "NON = Seulement " & entreprise, _
                   vbYesNoCancel + vbQuestion, "Port" & ChrW(233) & "e du code op" & ChrW(233) & "ration")

    If choix = vbCancel Then Exit Sub

    Application.ScreenUpdating = False

    Dim entToSave As String, cptToSave As String
    If choix = vbYes Then
        entToSave = "*"
        cptToSave = "*"
    Else
        entToSave = entreprise
        cptToSave = compte
    End If

    With wsCat
        .Cells(lastRowCat + 1, "A").Value = entToSave
        .Cells(lastRowCat + 1, "B").Value = "Classe 2"
        .Cells(lastRowCat + 1, "C").Value = cptToSave
        .Cells(lastRowCat + 1, "D").Value = newCategorie
        .Cells(lastRowCat + 1, "F").Value = "OUI"
    End With

    wsCat.Range("A2:F" & lastRowCat + 1).Sort Key1:=wsCat.Columns(1), Order1:=xlAscending, Header:=xlNo

    Dim categories As Variant
    categories = modData.GetCategorieMatches(entreprise, classe, compte)
    If Not IsEmpty(categories) Then
        If UBound(categories) >= 1 Then
            CreateValidatedDropdown wsMain.Cells(currentRow, "I"), categories, "Cat_Row" & currentRow, False
        End If
    End If

    wsMain.Cells(currentRow, "I").Value = newCategorie
    Application.ScreenUpdating = True

    MsgBox "Code op" & ChrW(233) & "ration ajout" & ChrW(233) & " !", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

Public Sub CleanupTempLists()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("TempLists")

    If Not ws Is Nothing Then
        ws.Cells.Clear
    End If

    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        If Left(nm.Name, 3) = "VL_" Or Left(nm.Name, 10) = "ValidList_" Then
            nm.Delete
        End If
    Next nm
End Sub
