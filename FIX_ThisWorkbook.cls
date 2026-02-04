Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
'========================================================================================
' ThisWorkbook - VERSION CORRIGEE POUR CO-EDITION
'
' CORRECTIONS :
'   - Workbook_Open SIMPLIFIE : plus de scroll automatique (causait des conflits
'     quand 2 utilisateurs ouvraient le fichier en meme temps)
'   - Workbook_BeforeClose : garde le nettoyage consultation, supprime le Me.Save
'     automatique (en co-edition, la sauvegarde est geree par SharePoint/OneDrive)
'========================================================================================

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
    ' Nettoyage des vues consultation temporaires
    modConsultation.SupprimerVuesTemporaires
    ' Annuler la minuterie admin si active
    modSecurity.CancelAutoRehide
End Sub
