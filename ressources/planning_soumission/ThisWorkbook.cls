Option Explicit
Private modificationEnCours As Boolean

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If modificationEnCours Then Exit Sub
    If Sh.Name <> "Planning soumission" Then Exit Sub
    On Error GoTo Erreur
    modificationEnCours = True
    HorizonModifier Sh, Target
Sortie:
    modificationEnCours = False
    Exit Sub
Erreur:
    MsgBox "La saisie de la periode n'a pas pu etre memorisee : " & Err.Description, vbExclamation, "Planning de soumission"
    Resume Sortie
End Sub

Private Sub Workbook_Open()
    On Error GoTo Sortie
    modificationEnCours = True
    HorizonLegendesToutes ThisWorkbook.Worksheets("Planning soumission"), False
Sortie:
    modificationEnCours = False
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If modificationEnCours Then Exit Sub
    On Error GoTo Erreur
    modificationEnCours = True
    HorizonLegendesToutes ThisWorkbook.Worksheets("Planning soumission"), True
Sortie:
    modificationEnCours = False
    Exit Sub
Erreur:
    Cancel = True
    MsgBox "Le texte du planning n'a pas pu etre memorise : " & Err.Description, vbExclamation, "Planning de soumission"
    Resume Sortie
End Sub
