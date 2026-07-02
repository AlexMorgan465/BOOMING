VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFBesoins 
   Caption         =   "UserForm1"
   ClientHeight    =   11370
   ClientLeft      =   -30
   ClientTop       =   -105
   ClientWidth     =   13065
   OleObjectBlob   =   "UFBesoins.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFBesoins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_ligneSelectionnee As Long
Private m_modeAjout As Boolean

Private Sub UserForm_Initialize()
    ' Projets de renfort
    cboProjet.Clear
    Dim Projets As Variant
    Projets = Array("EBRA PRESSE", "EBRA PRESS", "COFIT ITALY", "GOOGLE LEADS", _
                    "AFEDIM", "GLF", "TLV", "FACTO", "DAC")
    Dim p As Variant
    For Each p In Projets
        cboProjet.AddItem p
    Next p

    ' Jours
    cboJour.Clear
    Dim jours As Variant
    jours = Array("Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche")
    Dim j As Variant
    For Each j In jours
        cboJour.AddItem j
    Next j

    ' Semaine courante par défaut
    txtSemaine.Text = CStr(Application.WorksheetFunction.WeekNum(Date, 2))
    txtHDebut.Text = "08:00"
    txtHFin.Text = "18:00"
    txtNbAgents.Text = "1"

    m_ligneSelectionnee = 0
    m_modeAjout = False

    ChargerListe
    VerrouillerFormulaire True
End Sub

' --- Charger la liste ---
Private Sub ChargerListe()
    lstBesoins.Clear
    If Not FeuilleExiste("BESOINS") Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BESOINS")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) <> "" Then
            lstBesoins.AddItem ws.Cells(i, 1).Value
            lstBesoins.List(lstBesoins.ListCount - 1, 1) = CStr(ws.Cells(i, 2).Value)
            lstBesoins.List(lstBesoins.ListCount - 1, 2) = CStr(ws.Cells(i, 3).Value)
            lstBesoins.List(lstBesoins.ListCount - 1, 3) = CStr(ws.Cells(i, 4).Value)
            lstBesoins.List(lstBesoins.ListCount - 1, 4) = CStr(ws.Cells(i, 5).Value)
            lstBesoins.List(lstBesoins.ListCount - 1, 5) = CStr(ws.Cells(i, 6).Value)
        End If
    Next i
End Sub

' --- Sélection ---
Private Sub lstBesoins_Click()
    If lstBesoins.ListIndex < 0 Then Exit Sub
    m_modeAjout = False
    m_ligneSelectionnee = lstBesoins.ListIndex + 2

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BESOINS")

    cboProjet.Text = CStr(ws.Cells(m_ligneSelectionnee, 1).Value)
    txtSemaine.Text = CStr(ws.Cells(m_ligneSelectionnee, 2).Value)
    cboJour.Text = CStr(ws.Cells(m_ligneSelectionnee, 3).Value)
    txtHDebut.Text = CStr(ws.Cells(m_ligneSelectionnee, 4).Value)
    txtHFin.Text = CStr(ws.Cells(m_ligneSelectionnee, 5).Value)
    txtNbAgents.Text = CStr(ws.Cells(m_ligneSelectionnee, 6).Value)

    ' Résultats (lecture seule)
    txtAgents.Text = CStr(ws.Cells(m_ligneSelectionnee, 7).Value)
    txtNbDispo.Text = CStr(ws.Cells(m_ligneSelectionnee, 8).Value)
    txtStatut.Text = CStr(ws.Cells(m_ligneSelectionnee, 9).Value)

    VerrouillerFormulaire False
End Sub

' --- Boutons semaine rapide ---
Private Sub cmdSemCourante_Click()
    txtSemaine.Text = CStr(Application.WorksheetFunction.WeekNum(Date, 2))
End Sub

Private Sub cmdSemSuivante_Click()
    txtSemaine.Text = CStr(Application.WorksheetFunction.WeekNum(Date + 7, 2))
End Sub

' --- Nouveau ---
Private Sub cmdNouveau_Click()
    m_modeAjout = True
    m_ligneSelectionnee = 0
    ViderFormulaire
    VerrouillerFormulaire False
    cboProjet.SetFocus
End Sub

' --- Supprimer ---
Private Sub cmdSupprimer_Click()
    If lstBesoins.ListIndex < 0 Then
        MsgBox "Sélectionnez d'abord un besoin.", vbExclamation: Exit Sub
    End If
    Dim rep As Integer
    rep = MsgBox("Supprimer ce besoin ?", vbYesNo + vbWarning, "Confirmation")
    If rep = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BESOINS")
    ws.Rows(m_ligneSelectionnee).Delete

    ViderFormulaire
    VerrouillerFormulaire True
    m_ligneSelectionnee = 0
    ChargerListe
    MsgBox "Besoin supprimé.", vbInformation
End Sub

' --- Sauvegarder ---
Private Sub cmdSauver_Click()
    ' Validations
    If Trim(cboProjet.Text) = "" Then
        MsgBox "Le projet est obligatoire.", vbExclamation: cboProjet.SetFocus: Exit Sub
    End If
    If Not IsNumeric(txtSemaine.Text) Then
        MsgBox "Le numéro de semaine doit être un nombre.", vbExclamation: txtSemaine.SetFocus: Exit Sub
    End If
    If Trim(cboJour.Text) = "" Then
        MsgBox "Le jour est obligatoire.", vbExclamation: cboJour.SetFocus: Exit Sub
    End If
    If Trim(txtHDebut.Text) = "" Or Trim(txtHFin.Text) = "" Then
        MsgBox "Les heures sont obligatoires (format HH:MM).", vbExclamation: Exit Sub
    End If
    If Not IsNumeric(txtNbAgents.Text) Or CInt(txtNbAgents.Text) < 1 Then
        MsgBox "Le nombre d'agents doit être ≥ 1.", vbExclamation: txtNbAgents.SetFocus: Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("BESOINS")

    ' S'assurer que l'en-tête existe
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "Projet"
        ws.Cells(1, 2).Value = "Semaine"
        ws.Cells(1, 3).Value = "Jour"
        ws.Cells(1, 4).Value = "H Debut"
        ws.Cells(1, 5).Value = "H Fin"
        ws.Cells(1, 6).Value = "N°Agent"
        ws.Cells(1, 7).Value = "Agents proposes"
        ws.Cells(1, 8).Value = "Nb disponibles"
        ws.Cells(1, 9).Value = "Statut"
        With ws.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If

    Dim lr As Long
    If m_modeAjout Then
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Else
        lr = m_ligneSelectionnee
        ' Effacer les résultats précédents car le besoin a changé
        ws.Cells(lr, 7).Value = ""
        ws.Cells(lr, 8).Value = ""
        ws.Cells(lr, 9).Value = ""
        ws.Rows(lr).Interior.ColorIndex = xlNone
    End If

    ws.Cells(lr, 1).Value = Trim(cboProjet.Text)
    ws.Cells(lr, 2).Value = CInt(txtSemaine.Text)
    ws.Cells(lr, 3).Value = Trim(cboJour.Text)
    ws.Cells(lr, 4).Value = Trim(txtHDebut.Text)
    ws.Cells(lr, 5).Value = Trim(txtHFin.Text)
    ws.Cells(lr, 6).Value = CInt(txtNbAgents.Text)

    ' Effacer les champs résultats si nouveau ou modifié
    txtAgents.Text = ""
    txtNbDispo.Text = ""
    txtStatut.Text = "(sera mis à jour à la prochaine génération)"

    ChargerListe
    VerrouillerFormulaire True
    m_modeAjout = False
    MsgBox "Besoin " & IIf(lr = m_ligneSelectionnee, "modifié", "ajouté") & " avec succ�s !", vbInformation
End Sub

Private Sub cmdAnnuler_Click()
    If m_ligneSelectionnee > 0 Then
        lstBesoins_Click
    Else
        ViderFormulaire
    End If
    m_modeAjout = False
    VerrouillerFormulaire True
End Sub

Private Sub cmdFermer_Click()
    Unload Me
End Sub

' --- Helpers ---
Private Sub ViderFormulaire()
    cboProjet.Text = ""
    txtSemaine.Text = CStr(Application.WorksheetFunction.WeekNum(Date, 2))
    cboJour.Text = ""
    txtHDebut.Text = "08:00"
    txtHFin.Text = "18:00"
    txtNbAgents.Text = "1"
    txtAgents.Text = ""
    txtNbDispo.Text = ""
    txtStatut.Text = ""
End Sub

Private Sub VerrouillerFormulaire(bVerrouille As Boolean)
    cboProjet.Enabled = Not bVerrouille
    txtSemaine.Enabled = Not bVerrouille
    cmdSemCourante.Enabled = Not bVerrouille
    cmdSemSuivante.Enabled = Not bVerrouille
    cboJour.Enabled = Not bVerrouille
    txtHDebut.Enabled = Not bVerrouille
    txtHFin.Enabled = Not bVerrouille
    txtNbAgents.Enabled = Not bVerrouille
    cmdSauver.Enabled = Not bVerrouille
    cmdAnnuler.Enabled = Not bVerrouille
    cmdSupprimer.Enabled = Not bVerrouille And (lstBesoins.ListIndex >= 0)
End Sub

Private Function FeuilleExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not (ws Is Nothing)
End Function

Public Sub Ouvrir()
    UFBesoins.Show
End Sub


