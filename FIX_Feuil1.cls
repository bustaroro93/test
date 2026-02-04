Attribute VB_Name = "Feuil1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'========================================================================================
' FEUILLE "DEMANDE OS" - VERSION CORRIGEE POUR CO-EDITION
'
' CORRECTIONS :
'   - g_ActivateDone SUPPRIME (causait des etats desynchronises entre co-editeurs)
'   - Worksheet_Activate SIMPLIFIE : Formula1:="Oui,Non" direct (pas de VL_OuiNon)
'   - Header shields SUPPRIMES (complexite inutile)
'   - Plus d'appel a modFilters.EnforceAllowedFilterDropdowns dans Activate
'   - Tout le reste (modData, modUI, modDocument, modEmail) INCHANGE
'========================================================================================
Option Explicit

' CORRECTION : Plus de HEADER_ADDR, SHIELD_NAME, SHIELD_PREFIX, FILTER_GAP_PTS
' CORRECTION : Plus de g_ActivateDone

' ===== VENTILATION UF 9025 (GARDE) =====
Private Const UF_VENTILE As String = "9025"
Private Const PCT_8191 As Double = 0.55
Private Const PCT_8335 As Double = 0.36
Private Const PCT_8336 As Double = 0.09

' CORRECTION : Plus de EnsureHeaderRowLocked, DeleteHeaderShields,
' BringRealButtonsToFront, EnsureHeaderShield
' Ces fonctions manipulaient des Shapes sur l'en-tete et causaient des conflits

' ===== ACTIVATION DE LA FEUILLE =====
' CORRECTION MAJEURE : Version simplifiee comme en 2025
Private Sub Worksheet_Activate()
    On Error Resume Next

    ' Validation Oui/Non DIRECTE (pas de Named Range)
    With Me.Range("AC2:AC1500").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="Oui,Non"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' Limiter la zone de defilement
    Me.ScrollArea = "A1:AE1500"

    On Error GoTo 0
End Sub


' ===== GESTION DES CHANGEMENTS =====
' Structure IDENTIQUE au 2026 actuel, seuls les appels restent les memes
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Count > 1 Then Exit Sub
    If Target.Row < 2 Then Exit Sub

    On Error GoTo CleanUp

    Application.EnableEvents = False

    Dim rowNumber As Long
    rowNumber = Target.Row

    ' ===== 1. PROTECTION DES COLONNES AUTO-REMPLIES =====
    If Not Intersect(Target, Me.Range("A:A,F:F,J:J,K:K,L:L,M:M,N:N,S:S,T:T,V:V")) Is Nothing Then
        GoTo CleanUp
    End If

    ' ===== 2. VALIDATION HYPERLIEN COLONNE U =====
    If Target.Column = 21 Then
        If Target.Hyperlinks.Count = 0 And Len(Target.Value) > 0 Then
            Target.ClearContents
            MsgBox "Cette colonne n'accepte que des liens hypertexte." & vbCrLf & _
                   "Utilisez le clic droit pour inserer un lien.", vbExclamation
            GoTo CleanUp
        End If
    End If

    ' ===== 3. VALIDATION COMPTE (COLONNE G) =====
    If Target.Column = 7 Then
        On Error Resume Next
        If Not Target.Validation.Value Then
            Target.ClearContents
            MsgBox "Veuillez utiliser la liste deroulante pour selectionner un compte.", vbExclamation
            GoTo CleanUp
        End If
        On Error GoTo CleanUp
    End If

    ' ===== 4. STANDARDISATION DATE DEVIS (COLONNE P) =====
    If Target.Column = 16 Then
        Call StandardiserDateDevis(Target)
    End If

    ' ===== 5. GESTION COLONNE Y (STATUT + HORODATAGE) =====
    If Target.Column = 25 Then
        Dim statutValue As String
        statutValue = LCase(Trim(CStr(Target.Value)))

        If statutValue = "en attente" Then
            Dim classeValTemp As String
            classeValTemp = LCase(Trim(CStr(Me.Cells(rowNumber, "E").Value)))
            If classeValTemp = "classe 2" Then
                MsgBox "Ce dossier est en Classe 2. Veuillez penser a preparer le PV.", vbInformation
            End If
        End If

        If statutValue = "valid" & ChrW(233) Then
            If IsEmpty(Me.Cells(rowNumber, "AE").Value) Or Not IsDate(Me.Cells(rowNumber, "AE").Value) Then
                Me.Cells(rowNumber, "AE").Value = Now
                Me.Cells(rowNumber, "AE").NumberFormat = "dd/mm/yyyy HH:mm"
            End If
        End If
    End If

    ' ===== 6. VERIFICATION DUPLICATION NUMERO DEVIS (COLONNE O) =====
    If Target.Column = 15 Then
        Call VerifierDuplicationDevis(Target, rowNumber)
    End If

    ' ===== 7. VERIFICATION DUPLICATION NUMERO OS (COLONNE AA) =====
    If Target.Column = 27 Then
        Call VerifierDuplicationOS(Target, rowNumber)
    End If

    ' ===== 8. TRAITEMENT TRAVAUX TERMINES (COLONNE AC) =====
    If Target.Column = 29 Then
        Call TraiterTravauxTermines(Target)
    End If

    ' ===== 9. GESTION UF (COLONNE D) =====
    If Target.Column = 4 Then
        Call TraiterModificationUF(rowNumber)
    End If

    ' ===== 10. GESTION ENTREPRISE/CLASSE (COLONNES C et E) =====
    If Target.Column = 3 Or Target.Column = 5 Then
        Call TraiterModificationEntrepriseClasse(rowNumber)
    End If

    ' ===== 11. GESTION COMPTE (COLONNE G) =====
    If Target.Column = 7 Then
        Call TraiterModificationCompte(rowNumber)
    End If

    ' ===== 12. VALIDATION MONTANTS (COLONNE R) =====
    If Target.Column = 18 Then
        Call ValiderMontantHT(Target)

        If IsEmpty(Target.Value) Or Trim$(CStr(Target.Value)) = "" Then
            Me.Cells(rowNumber, "T").ClearContents
            GoTo CleanUp
        End If

        ' ===== CAS SPECIAL : UF 9025 => ventilation =====
        If Trim$(CStr(Me.Cells(rowNumber, "D").Value)) = UF_VENTILE Then

            If Not IsNumeric(Target.Value) Then
                Target.ClearContents
                MsgBox "Veuillez saisir un montant HT numerique.", vbExclamation
                GoTo CleanUp
            End If

            Dim total As Double
            total = CDbl(Target.Value)

            If total <= 0 Then
                Target.ClearContents
                Me.Cells(rowNumber, "T").ClearContents
                GoTo CleanUp
            End If

            Dim m8191 As Double, m8335 As Double, m8336 As Double
            m8191 = Round(total * PCT_8191, 2)
            m8335 = Round(total * PCT_8335, 2)
            m8336 = Round(total - m8191 - m8335, 2)

            Me.Cells(rowNumber, "R").NumberFormat = "@"
            Me.Cells(rowNumber, "R").Value = _
                "8191 : " & Format(m8191, "#,##0.00") & Chr(10) & _
                "8335 : " & Format(m8335, "#,##0.00") & Chr(10) & _
                "8336 : " & Format(m8336, "#,##0.00")
            Me.Cells(rowNumber, "R").WrapText = True

            Dim tvaRate As Double
            On Error Resume Next
            tvaRate = CDbl(Me.Cells(rowNumber, "S").Value)
            On Error GoTo CleanUp

            If tvaRate >= 0 Then
                Me.Cells(rowNumber, "T").Value = total * (1 + tvaRate)
                Me.Cells(rowNumber, "T").NumberFormat = "#,##0.00"
            Else
                Me.Cells(rowNumber, "T").ClearContents
            End If

            If IsEmpty(Me.Cells(rowNumber, "V").Value) Then
                Me.Cells(rowNumber, "V").Value = Date
            End If

            Me.Rows(rowNumber).AutoFit
            GoTo CleanUp
        End If

        ' ===== CAS NORMAL (UF != 9025) =====
        Dim montantHT As Double, tvaRateN As Double
        On Error Resume Next
        montantHT = CDbl(Me.Cells(rowNumber, "R").Value)
        tvaRateN = CDbl(Me.Cells(rowNumber, "S").Value)
        On Error GoTo CleanUp

        If montantHT > 0 And tvaRateN >= 0 Then
            Me.Cells(rowNumber, "T").Value = montantHT * (1 + tvaRateN)
        ElseIf IsEmpty(Me.Cells(rowNumber, "R").Value) Then
            Me.Cells(rowNumber, "T").ClearContents
        End If

        Me.Cells(rowNumber, "R").NumberFormat = "#,##0.00"
        Me.Cells(rowNumber, "S").NumberFormat = "0.00"
        Me.Cells(rowNumber, "T").NumberFormat = "#,##0.00"

        If IsEmpty(Me.Cells(rowNumber, "V").Value) And Not IsEmpty(Target.Value) Then
            Me.Cells(rowNumber, "V").Value = Date
        End If
    End If

CleanUp:
    Application.EnableEvents = True
End Sub

' ===== CLIC DROIT POUR AJOUTER LIEN DOSSIER (COLONNE U) =====
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Column = 21 Then
        Cancel = True

        Dim cheminRacine As String
        cheminRacine = "L:\Services\SERVICE_DES_TRAVAUX\_NEW SERVEUR DPT\2000_PLAN TVX COURANT cl2-cl6\2026\020 Suivi plan de travaux 2026"

        Dim selectedPath As String
        selectedPath = GetFolder(cheminRacine)

        If selectedPath <> "" Then
            ActiveSheet.Hyperlinks.Add _
                Anchor:=Target, _
                Address:=selectedPath, _
                TextToDisplay:=selectedPath
        End If
    End If
End Sub

' ===== DOUBLE CLIC - PROTECTION =====
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Me.Range("A:A,F:F,J:J,K:K,L:L,M:M,N:N,S:S,T:T,V:V")) Is Nothing Then
        Cancel = True
        MsgBox "Cette colonne se remplit automatiquement et ne peut pas etre modifiee.", vbExclamation
    End If
End Sub

' ===== SELECTION - AUTO-REFERENT =====
' CORRECTION : Suppression du blocage ligne 1 via modSecurity.CanFilterColumn
' (c'etait une source de conflits car ca touchait a la navigation)
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo ExitHere
    If Target Is Nothing Then Exit Sub
    If Target.CountLarge > 1 Then Exit Sub

    ' Auto-referent uniquement en B2:B1500
    If Intersect(Target, Me.Range("B2:B1500")) Is Nothing Then Exit Sub
    If Len(Trim$(CStr(Target.Value))) > 0 Then Exit Sub

    Dim refName As String
    refName = ResolveCurrentUserReferent()
    If Len(refName) = 0 Then Exit Sub
    If Not IsReferentInList(refName) Then Exit Sub

    Application.EnableEvents = False
    Target.Value = refName

ExitHere:
    Application.EnableEvents = True
End Sub


' ===== FONCTIONS UTILITAIRES PRIVEES =====
' (Toutes INCHANGEES par rapport au 2026 actuel)

Private Sub StandardiserDateDevis(ByVal Target As Range)
    On Error GoTo ExitProc
    If IsEmpty(Target.Value) Then Exit Sub

    If IsDate(Target.Value) Then
        Target.NumberFormat = "dd/mm/yyyy"
        Exit Sub
    End If

    Dim dateSaisie As String
    dateSaisie = Trim(CStr(Target.Value))
    dateSaisie = Replace(Replace(Replace(dateSaisie, "/", ""), "-", ""), " ", "")

    If IsNumeric(dateSaisie) And (Len(dateSaisie) = 6 Or Len(dateSaisie) = 8) Then
        Dim jour As Integer, mois As Integer, annee As Integer

        jour = CInt(Left(dateSaisie, 2))
        mois = CInt(Mid(dateSaisie, 3, 2))

        If Len(dateSaisie) = 6 Then
            annee = CInt(Right(dateSaisie, 2))
            If annee >= 0 And annee <= 30 Then
                annee = annee + 2000
            Else
                annee = annee + 1900
            End If
        Else
            annee = CInt(Right(dateSaisie, 4))
        End If

        If jour > 0 And jour <= 31 And mois > 0 And mois <= 12 Then
            Target.Value = DateSerial(annee, mois, jour)
            Target.NumberFormat = "dd/mm/yyyy"
        Else
            Target.ClearContents
            MsgBox "Date invalide. Format attendu : JJMMAA ou JJ/MM/AA", vbExclamation
        End If
    Else
        Target.ClearContents
        MsgBox "Format de date incorrect. Utilisez : JJMMAA ou JJ/MM/AA", vbExclamation
    End If

ExitProc:
    On Error GoTo 0
End Sub

Private Sub VerifierDuplicationDevis(ByVal Target As Range, ByVal rowNumber As Long)
    Dim devisVal As String
    devisVal = Trim(CStr(Target.Value))
    If devisVal = "" Then Exit Sub

    Dim lastRowO As Long
    lastRowO = Me.Cells(Me.Rows.Count, "O").End(xlUp).Row
    If lastRowO < 2 Then Exit Sub

    Dim matchRowDevis As Variant
    matchRowDevis = Application.Match(devisVal, Me.Range("O2:O" & lastRowO), 0)

    If Not IsError(matchRowDevis) Then
        Dim foundLineDevis As Long
        foundLineDevis = matchRowDevis + 1
        If foundLineDevis <> rowNumber Then
            Dim foundDossierDevis As String
            foundDossierDevis = CStr(Me.Cells(foundLineDevis, "A").Value)
            MsgBox "Le numero de devis '" & devisVal & "' est deja utilise a la ligne " & foundLineDevis & _
                   " (Dossier : " & foundDossierDevis & ").", vbExclamation
            Target.ClearContents
        End If
    End If
End Sub

Private Sub VerifierDuplicationOS(ByVal Target As Range, ByVal rowNumber As Long)
    Dim osVal As Variant
    osVal = Target.Value
    If IsEmpty(osVal) Then Exit Sub

    Dim lastRowOS As Long
    lastRowOS = Me.Cells(Me.Rows.Count, "AA").End(xlUp).Row
    If lastRowOS < 2 Then Exit Sub

    Dim matchRowOS As Variant
    matchRowOS = Application.Match(osVal, Me.Range("AA2:AA" & lastRowOS), 0)

    If Not IsError(matchRowOS) Then
        Dim foundLineOS As Long
        foundLineOS = matchRowOS + 1
        If foundLineOS <> rowNumber Then
            Dim foundDossierOS As String
            foundDossierOS = CStr(Me.Cells(foundLineOS, "A").Value)
            MsgBox "Le numero d'OS " & osVal & " est deja utilise a la ligne " & foundLineOS & _
                   " (Dossier : " & foundDossierOS & ").", vbExclamation
            Target.ClearContents
        End If
    End If
End Sub

Private Sub TraiterTravauxTermines(ByVal Target As Range)
    Dim cell As Range
    For Each cell In Intersect(Target, Me.Range("AC2:AC1500"))
        Select Case LCase(Trim(cell.Value))
            Case "oui": cell.Interior.Color = RGB(198, 239, 206)
            Case "non": cell.Interior.Color = RGB(255, 199, 206)
            Case Else: cell.Interior.Color = vbWhite
        End Select

        If LCase(Trim(cell.Value)) = "oui" Then
            Dim rep As VbMsgBoxResult
            rep = MsgBox("Travaux termines pour la ligne " & cell.Row & "." & vbCrLf & vbCrLf & _
                         "Generer le P.V. (si Classe 2) et preparer l'e-mail ?", _
                         vbQuestion + vbYesNo)

            If rep = vbYes Then
                Call modDocument.ProcessusCompletPourLigne(Me, cell.Row)
            End If
        End If
    Next cell
End Sub

Private Sub TraiterModificationUF(ByVal rowNumber As Long)
    Dim ufCell As Range
    Set ufCell = Me.Cells(rowNumber, "D")

    Dim UFVal As String
    UFVal = Trim(CStr(ufCell.Value))

    If UFVal = "" Then
        Call ClearUFDependentFields(rowNumber)
        Exit Sub
    End If

    Dim matches As Variant
    matches = modData.GetUFSMatches(UFVal)

    If UBound(matches, 1) < 1 Then
        MsgBox "UF introuvable. Merci de saisir un UF valide.", vbExclamation
        ufCell.ClearContents
        Call ClearUFDependentFields(rowNumber)
        Exit Sub
    End If

    If UBound(matches, 1) = 1 Then
        Call FillUFData(rowNumber, matches(1, 1), matches(1, 2), matches(1, 3), _
                       matches(1, 4), matches(1, 5), matches(1, 6), matches(1, 7), matches(1, 8))
    Else
        MsgBox "Plusieurs correspondances trouvees pour cet UF.", vbExclamation
    End If
End Sub

Private Sub TraiterModificationEntrepriseClasse(ByVal rowNumber As Long)
    Dim entrepriseCell As Range, classeCell As Range, compteCell As Range
    Set entrepriseCell = Me.Cells(rowNumber, "C")
    Set classeCell = Me.Cells(rowNumber, "E")
    Set compteCell = Me.Cells(rowNumber, "G")

    Dim EntrepriseVal As String, ClasseVal As String
    EntrepriseVal = Trim(CStr(entrepriseCell.Value))
    ClasseVal = Trim(CStr(classeCell.Value))

    modUI.ClearValidation compteCell
    modData.ClearCompteIPCat rowNumber

    If EntrepriseVal = "" Or ClasseVal = "" Then Exit Sub

    Dim comptes As Variant
    comptes = modData.GetComptesMatches(EntrepriseVal, ClasseVal)

    If UBound(comptes) < 1 Then
        MsgBox "Aucun compte disponible pour cette entreprise et cette classe.", vbExclamation
        Exit Sub
    End If

    If UBound(comptes) = 1 Then
        compteCell.Value = comptes(1)
        Call GenerateDependencies(rowNumber, EntrepriseVal, ClasseVal, compteCell.Value)
    Else
        compteCell.ClearContents
        Call modUI.CreateValidatedDropdown(compteCell, comptes, "Compte_Row" & rowNumber, False)
        MsgBox "Merci de choisir le compte budgetaire dans la liste deroulante.", vbExclamation
    End If
End Sub

Private Sub TraiterModificationCompte(ByVal rowNumber As Long)
    Dim entrepriseCell As Range, classeCell As Range, compteCell As Range
    Set entrepriseCell = Me.Cells(rowNumber, "C")
    Set classeCell = Me.Cells(rowNumber, "E")
    Set compteCell = Me.Cells(rowNumber, "G")

    Dim EntrepriseVal As String, ClasseVal As String, CompteVal As String
    EntrepriseVal = Trim(CStr(entrepriseCell.Value))
    ClasseVal = Trim(CStr(classeCell.Value))
    CompteVal = Trim(CStr(compteCell.Value))

    If CompteVal = "" Then
        Me.Cells(rowNumber, "H").ClearContents
        Me.Cells(rowNumber, "I").ClearContents
        Exit Sub
    End If

    Call GenerateDependencies(rowNumber, EntrepriseVal, ClasseVal, CompteVal)
End Sub

Private Sub GenerateDependencies(rowNumber As Long, EntrepriseVal As String, ClasseVal As String, CompteVal As String)
    Dim ipCell As Range, catCell As Range
    Set ipCell = Me.Cells(rowNumber, "H")
    Set catCell = Me.Cells(rowNumber, "I")

    If CompteVal = "" Then
        ipCell.ClearContents
        catCell.ClearContents
        Exit Sub
    End If

    Dim imputations As Variant
    imputations = modData.GetImputationMatches(EntrepriseVal, ClasseVal, CompteVal)

    If UBound(imputations) = 1 Then
        ipCell.Value = imputations(1)
    ElseIf UBound(imputations) > 1 Then
        ipCell.ClearContents
        Call modUI.CreateValidatedDropdown(ipCell, imputations, "IP_Row" & rowNumber, False)
    End If

    Dim categories As Variant
    categories = modData.GetCategorieMatches(EntrepriseVal, ClasseVal, CompteVal)

    If UBound(categories) = 1 Then
        catCell.Value = categories(1)
    ElseIf UBound(categories) > 1 Then
        catCell.ClearContents
        Call modUI.CreateValidatedDropdown(catCell, categories, "Cat_Row" & rowNumber, False)
        MsgBox "Veuillez choisir un code operation dans la liste deroulante.", vbInformation
    End If
End Sub

Private Sub ValiderMontantHT(ByVal Target As Range)
    If Not IsEmpty(Target.Value) Then
        If Not IsNumeric(Target.Value) Then
            Target.ClearContents
            MsgBox "Veuillez entrer une valeur numerique.", vbExclamation
        ElseIf Target.Value < 0 Then
            Target.ClearContents
            MsgBox "Le montant ne peut pas etre negatif.", vbExclamation
        End If
    End If
End Sub

Private Sub ClearUFDependentFields(ByVal rowNumber As Long)
    Me.Cells(rowNumber, "F").ClearContents
    Me.Cells(rowNumber, "S").ClearContents
    Me.Cells(rowNumber, "J").ClearContents
    Me.Cells(rowNumber, "K").ClearContents
    Me.Cells(rowNumber, "L").ClearContents
    modUI.ClearValidation Me.Cells(rowNumber, "M")
    modUI.ClearValidation Me.Cells(rowNumber, "N")
End Sub

Private Sub FillUFData(ByVal rowNumber As Long, ByVal uf As String, ByVal Ville As String, _
                      ByVal adresse As String, ByVal Secteur As String, ByVal Site As String, _
                      ByVal Service As String, ByVal Budget As String, ByVal TVA As Variant)
    Me.Cells(rowNumber, "J").Value = Ville
    Me.Cells(rowNumber, "K").Value = adresse
    Me.Cells(rowNumber, "L").Value = Secteur
    Me.Cells(rowNumber, "M").Value = Site
    Me.Cells(rowNumber, "N").Value = Service
    Me.Cells(rowNumber, "F").Value = Budget
    Me.Cells(rowNumber, "S").Value = TVA

    On Error Resume Next
    Me.Cells(rowNumber, "M").Validation.Delete
    Me.Cells(rowNumber, "N").Validation.Delete
    On Error GoTo 0
End Sub

Private Function GetFolder(initialPath As String) As String
    With Application.FileDialog(4) ' msoFileDialogFolderPicker
        .InitialFileName = initialPath & "\"
        .Title = "Selectionner le dossier"
        If .Show = -1 Then
            GetFolder = .SelectedItems(1)
        Else
            GetFolder = ""
        End If
    End With
End Function

Private Function ResolveCurrentUserReferent() As String
    On Error GoTo Done
    Dim login As String
    login = Environ$("USERNAME")
    If Len(Trim$(login)) = 0 Then login = Application.UserName

    Dim ws As Worksheet, lastRow As Long, f As Range
    Set ws = ThisWorkbook.Worksheets("UsersMap")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set f = ws.Range("A1:A" & lastRow).Find(What:=login, LookIn:=xlValues, _
                                            LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then
        ResolveCurrentUserReferent = Trim$(CStr(f.Offset(0, 1).Value))
    End If
Done:
End Function

Private Function IsReferentInList(ByVal nameRef As String) As Boolean
    On Error GoTo Nope
    Dim ws As Worksheet, lastRow As Long, pos As Variant
    Set ws = ThisWorkbook.Worksheets("ReferentsDPT")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    pos = Application.Match(nameRef, ws.Range("A2:A" & lastRow), 0)
    IsReferentInList = Not IsError(pos)
    Exit Function
Nope:
    IsReferentInList = False
End Function
