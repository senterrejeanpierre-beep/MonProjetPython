Option Explicit

Public Function HorizonChoisir(ByVal feuille As Object, ByVal ligne As Long) As Long
    Dim nombre As Long, periode As Long, i As Long, premiere As Long
    Dim texte As String
    nombre = CLng(feuille.Cells(1, 159).Value2)
    premiere = 160 + 2 * nombre
    periode = CLng(feuille.Cells(ligne, 159).Value2)
    texte = CStr(feuille.Cells(ligne, 4).Value2)
    For i = 1 To nombre
        If texte = CStr(feuille.Cells(ligne, premiere + i - 1).Value2) Then
            HorizonChoisir = i
            Exit Function
        End If
    Next i
    If IsNumeric(texte) Then
        If CDbl(texte) >= 1 And CDbl(texte) <= nombre Then
            If CDbl(texte) = Fix(CDbl(texte)) Then
                HorizonChoisir = CLng(texte)
                Exit Function
            End If
        End If
    End If
    If texte = "" Then
        feuille.Cells(ligne, premiere + periode - 1).Value2 = periode
    Else
        feuille.Cells(ligne, premiere + periode - 1).Value2 = texte
    End If
    HorizonChoisir = periode
End Function

Public Sub HorizonAfficher(ByVal feuille As Object, ByVal ligne As Long)
    Dim periode As Long, colonne As Long, reference As String, i As Long, nombre As Long
    periode = CLng(feuille.Cells(ligne, 159).Value2)
    nombre = CLng(feuille.Cells(1, 159).Value2)
    colonne = 160 + (periode - 1) * 2
    feuille.Cells(ligne, 4).Value2 = feuille.Cells(ligne, 160 + 2 * nombre + periode - 1).Value2
    For i = 0 To 1
        reference = feuille.Cells(ligne, colonne + i).Address(False, False)
        feuille.Cells(ligne, 5 + i).Formula = "=IF(" & reference & "="""",""""," & reference & ")"
    Next i
End Sub

Public Sub HorizonModifier(ByVal feuille As Object, ByVal cible As Object)
    Dim zone As Object, ligne As Object, touche As Object
    Dim r As Long, periode As Long, i As Long, derniere As Long, nombre As Long
    derniere = CLng(feuille.Cells(2, 159).Value2)
    nombre = CLng(feuille.Cells(1, 159).Value2)
    Set touche = Application.Intersect(cible, feuille.Range("G4"))
    If Not touche Is Nothing Then HorizonLegendesToutes feuille, True
    Set zone = Application.Intersect(cible, feuille.Range("D9:F" & derniere))
    If zone Is Nothing Then Exit Sub
    For Each ligne In zone.Rows
        r = ligne.Row
        Set touche = Application.Intersect(cible, feuille.Cells(r, 4))
        HorizonLireLegendes feuille, r, (touche Is Nothing)
        periode = HorizonChoisir(feuille, r)
        If periode < 1 Or periode > nombre Then
            HorizonAfficher feuille, r
            MsgBox "Choisissez une periode dans la liste.", vbExclamation, "Planning de soumission"
        Else
            feuille.Cells(r, 159).Value2 = periode
            For i = 0 To 1
                Set touche = Application.Intersect(cible, feuille.Cells(r, 5 + i))
                If Not touche Is Nothing Then
                    feuille.Cells(r, 160 + (periode - 1) * 2 + i).Value2 = feuille.Cells(r, 5 + i).Value2
                End If
            Next i
            HorizonAfficher feuille, r
            HorizonLegendesLigne feuille, r
        End If
    Next ligne
    feuille.Calculate
End Sub

Public Function HorizonTexteLegende(ByVal label As Variant) As String
    If VarType(label) = vbString Then HorizonTexteLegende = CStr(label)
End Function

Public Function HorizonForme(ByVal feuille As Object, ByVal nom As String) As Object
    On Error Resume Next
    Set HorizonForme = feuille.Shapes(nom)
    On Error GoTo 0
End Function

Public Sub HorizonLireLegendes(ByVal feuille As Object, ByVal r As Long, ByVal afficher As Boolean)
    Dim n As Long, nombre As Long, boite As Object, texte As String, ancien As String
    nombre = CLng(feuille.Cells(1, 159).Value2)
    For n = 1 To nombre
        Set boite = HorizonForme(feuille, "HorizonLegende_" & r & "_" & n)
        If Not boite Is Nothing Then
            texte = boite.TextFrame.Characters.Text
            ancien = CStr(feuille.Cells(r, 160 + 3 * nombre + n - 1).Value2)
            If texte <> ancien Then
                If texte = "" Then
                    feuille.Cells(r, 160 + 2 * nombre + n - 1).Value2 = n
                Else
                    feuille.Cells(r, 160 + 2 * nombre + n - 1).Value2 = texte
                End If
                feuille.Cells(r, 160 + 3 * nombre + n - 1).Value2 = texte
                If afficher And n = CLng(feuille.Cells(r, 159).Value2) Then
                    feuille.Cells(r, 4).Value2 = feuille.Cells(r, 160 + 2 * nombre + n - 1).Value2
                End If
            End If
        End If
    Next n
End Sub

Public Sub HorizonLegendesLigne(ByVal feuille As Object, ByVal r As Long)
    Dim n As Long, m As Long, j As Long, nombre As Long, premier As Long, dernier As Long
    Dim boite As Object, cellule As Object
    Dim texte As String, debut As Variant, fin As Variant, autre As Variant
    Dim jour As Double, gauche As Double, largeur As Double, limite As Double, hauteur As Double, souhaite As Double
    nombre = CLng(feuille.Cells(1, 159).Value2)
    hauteur = feuille.Cells(r, 7).Height
    jour = CDbl(feuille.Range("G4").Value2)
    For n = 1 To nombre
        Set boite = HorizonForme(feuille, "HorizonLegende_" & r & "_" & n)
        If Not boite Is Nothing Then
            texte = HorizonTexteLegende(feuille.Cells(r, 160 + 2 * nombre + n - 1).Value2)
            boite.TextFrame.Characters.Text = texte
            feuille.Cells(r, 160 + 3 * nombre + n - 1).Value2 = texte
            boite.Width = 0.1
            boite.Height = 0.1
            boite.Left = feuille.Cells(r, 7).Left
            boite.Top = feuille.Cells(r, 7).Top + 1
            debut = feuille.Cells(r, 160 + 2 * (n - 1)).Value2
            fin = feuille.Cells(r, 161 + 2 * (n - 1)).Value2
            premier = 0
            dernier = 0
            If texte <> "" And IsNumeric(debut) And IsNumeric(fin) Then
                If debut >= 1 And fin > debut - 1 Then
                    For j = 7 To 146
                        If jour + j - 7 >= debut And jour + j - 7 < fin + 1 Then
                            If premier = 0 Then premier = j
                            dernier = j
                        End If
                    Next j
                End If
            End If
            If premier > 0 Then
                Set cellule = feuille.Cells(r, premier)
                gauche = cellule.Left + 1
                limite = feuille.Cells(r, 146).Left + feuille.Cells(r, 146).Width
                For m = 1 To nombre
                    If m <> n Then
                        autre = feuille.Cells(r, 160 + 2 * (m - 1)).Value2
                        If IsNumeric(autre) Then
                            If autre > jour + premier - 7 And autre <= jour + 139 Then
                                If HorizonTexteLegende(feuille.Cells(r, 160 + 2 * nombre + m - 1).Value2) <> "" Then
                                    j = 7 + CLng(-Int(-CDbl(autre))) - CLng(jour)
                                    If feuille.Cells(r, j).Left - 1 < limite Then limite = feuille.Cells(r, j).Left - 1
                                End If
                            End If
                        End If
                    End If
                Next m
                largeur = feuille.Cells(r, dernier).Left + feuille.Cells(r, dernier).Width - gauche
                souhaite = Len(texte) * feuille.Cells(r, 2).Font.Size * 0.55 + 4
                If souhaite > largeur Then largeur = souhaite
                If largeur > limite - gauche Then largeur = limite - gauche
                If largeur < 0.1 Then largeur = 0.1
                boite.Left = gauche
                boite.Width = largeur
                boite.Height = hauteur - 2
            End If
        End If
    Next n
End Sub

Public Sub HorizonLegendesToutes(ByVal feuille As Object, ByVal memoriser As Boolean)
    Dim r As Long
    For r = 9 To CLng(feuille.Cells(2, 159).Value2)
        If memoriser Then HorizonLireLegendes feuille, r, True
        HorizonLegendesLigne feuille, r
    Next r
End Sub
