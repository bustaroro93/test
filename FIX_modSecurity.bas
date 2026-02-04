Attribute VB_Name = "modSecurity"
Option Explicit

Public Const WB_PWD As String = "C23j0813"
Public Const ADMIN_PWD As String = "C23j0813"
Public g_AdminMode As Boolean
Private AllowedVisibleNames As Variant
Private scheduledTime As Date
' CORRECTION : g_SecurityDone SUPPRIME
' Ce flag empechait la re-initialisation et rendait le comportement
' imprevisible quand plusieurs utilisateurs ouvrent le fichier.

' CORRECTION : eFilterLevel / CurrentFilterLevel SUPPRIMES
' Tout le monde a les memes droits de filtrage = moins de complexity = moins de conflits

' ==============================================================
' SECURITE PRINCIPALE - VERSION CORRIGEE POUR CO-EDITION
' ==============================================================
' CORRECTION PRINCIPALE :
' - Plus de g_SecurityDone (la securite est appliquee a chaque ouverture)
' - Plus d'appel a modFilters.EnforceAllowedFilterDropdowns (qui faisait Unprotect/Protect)
' - On s'assure que TempLists existe pendant qu'on a le classeur deprotege
' - La protection "Demande OS" est faite UNE SEULE FOIS avec UserInterfaceOnly:=True
' ==============================================================
Public Sub EnforceSecurity()
    AllowedVisibleNames = Array("Demande OS", "Liste des UFS pour les techs", "Tableau de Bord")

    Application.ScreenUpdating = False
    On Error Resume Next

    ' 1. Deverrouiller la structure du classeur
    ThisWorkbook.Unprotect Password:=WB_PWD

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsAllowed(ws.Name) Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' 2. S'assurer que TempLists existe (pendant que la structure est deprotegee)
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

    ' 3. Protection de "Demande OS" - SIMPLE, UNE SEULE FOIS
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Demande OS")
    On Error GoTo 0

    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Unprotect Password:="cidalianeves"

        ' Activer les filtres simplement
        If Not ws.AutoFilterMode Then ws.Range("A1:AE1").AutoFilter
        If ws.FilterMode Then ws.ShowAllData

        ' CORRECTION : Protection SIMPLE, pas de boucle EnforceAllowedFilterDropdowns
        ws.Protect Password:="cidalianeves", _
                   DrawingObjects:=True, _
                   Contents:=True, _
                   Scenarios:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True, _
                   UserInterfaceOnly:=True
        On Error GoTo 0
    End If

    ' 4. Reverrouiller la structure du classeur
    ThisWorkbook.Protect Password:=WB_PWD, Structure:=True, Windows:=False

    Application.ScreenUpdating = True
End Sub

' ==============================================================
' WATCHDOG : EMPECHE L'ACCES AUX FEUILLES INTERDITES
' ==============================================================
Public Sub RehideIfUnauthorized(ByVal Sh As Object)
    On Error GoTo ExitClean
    If Sh Is Nothing Then Exit Sub
    If g_AdminMode Then Exit Sub

    If Not IsAllowed(Sh.Name) Then
        Application.EnableEvents = False
        ThisWorkbook.Unprotect Password:=WB_PWD

        Sh.Visible = xlSheetVeryHidden
        ThisWorkbook.Worksheets("Demande OS").Activate

        ThisWorkbook.Protect Password:=WB_PWD, Structure:=True, Windows:=False
    End If
ExitClean:
    Application.EnableEvents = True
End Sub

' ==============================================================
' EMPECHE LA CREATION DE NOUVELLES FEUILLES PAR L'UTILISATEUR
' ==============================================================
Public Sub HandleNewSheet(ByVal Sh As Object)
    On Error Resume Next
    If Sh Is Nothing Then Exit Sub
    If Not IsAllowed(Sh.Name) Then
        ThisWorkbook.Unprotect Password:=WB_PWD
        Sh.Visible = xlSheetVeryHidden
        ThisWorkbook.Protect Password:=WB_PWD, Structure:=True, Windows:=False
    End If
End Sub

' ==============================================================
' VERIFIE SI UNE FEUILLE A LE DROIT D'ETRE VUE
' ==============================================================
Private Function IsAllowed(ByVal sheetName As String) As Boolean
    Dim v

    ' 1. Les vues temporaires de consultation sont TOUJOURS autorisees
    If Left$(sheetName, 12) = "FILTRE_PERSO" Then
        IsAllowed = True
        Exit Function
    End If

    ' 2. Verification dans la liste
    If IsEmpty(AllowedVisibleNames) Then _
        AllowedVisibleNames = Array("Demande OS", "Liste des UFS pour les techs", "Tableau de Bord")

    For Each v In AllowedVisibleNames
        If StrComp(sheetName, CStr(v), vbTextCompare) = 0 Then
            IsAllowed = True
            Exit Function
        End If
    Next v
End Function

' ==============================================================
' MODE ADMIN (DEVERROUILLAGE TOTAL TEMPORAIRE)
' ==============================================================
Public Sub AdminShow()
    Dim p As String, ws As Worksheet
    p = InputBox("Mot de passe administrateur ?", "Mode Admin")
    If p <> ADMIN_PWD Then
        MsgBox "Acces refuse.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ThisWorkbook.Unprotect Password:=WB_PWD

    For Each ws In ThisWorkbook.Worksheets
        If Not IsAllowed(ws.Name) Then ws.Visible = xlSheetVisible
    Next ws

    ThisWorkbook.Protect Password:=WB_PWD, Structure:=True, Windows:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    g_AdminMode = True
    ArmAutoRehide Now + TimeSerial(0, 10, 0)

    MsgBox "Mode Admin actif : les feuilles sensibles sont visibles temporairement." & vbCrLf & _
           "Elles seront remasquees automatiquement dans 10 minutes.", vbInformation
End Sub

Public Sub AdminHide()
    CancelAutoRehide
    g_AdminMode = False
    EnforceSecurity
    MsgBox "Mode Admin desactive : feuilles sensibles remasquees.", vbInformation
End Sub

' --- Minuterie auto re-hide ---
Public Sub ArmAutoRehide(ByVal runAt As Date)
    scheduledTime = runAt
    On Error Resume Next
    Application.OnTime EarliestTime:=scheduledTime, Procedure:="modSecurity.AdminHide", Schedule:=True
End Sub

Public Sub CancelAutoRehide()
    If scheduledTime <> 0 Then
        On Error Resume Next
        Application.OnTime EarliestTime:=scheduledTime, Procedure:="modSecurity.AdminHide", Schedule:=False
        scheduledTime = 0
    End If
End Sub

' ==============================================================
' FONCTIONS DE COMPATIBILITE
' Ces fonctions sont appelees par modFilters et Feuil1
' On les garde pour ne pas casser les references, mais simplifiees
' ==============================================================
Public Function CanUseAnyFilters() As Boolean
    CanUseAnyFilters = True
End Function

Public Function CanSort() As Boolean
    CanSort = True
End Function

Public Function CanFilterColumn(ByVal colIndex As Long) As Boolean
    CanFilterColumn = True
End Function

Public Function CanUseFilters() As Boolean
    CanUseFilters = True
End Function
