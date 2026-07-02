VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFGenerer 
   Caption         =   "G�n�rer le Planning"
   ClientHeight    =   9450.001
   ClientLeft      =   -120
   ClientTop       =   -615
   ClientWidth     =   8715.001
   OleObjectBlob   =   "UFGenerer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFGenerer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    OptionSemCourante.Value = True
    txtDatePerso.Enabled = False
    txtDatePerso.Text = Format(Date, "dd/mm/yyyy")
    lblInfo.Caption = "Laissez vide ou entrez n'importe quelle date de la semaine souhait�."
    AfficherApercu
End Sub

Private Sub OptionSemCourante_Click()
    txtDatePerso.Enabled = False
    AfficherApercu
End Sub

Private Sub OptionSemSuivante_Click()
    txtDatePerso.Enabled = False
    AfficherApercu
End Sub

Private Sub OptionSemPerso_Click()
    txtDatePerso.Enabled = True
    txtDatePerso.SetFocus
    AfficherApercu
End Sub

Private Sub txtDatePerso_Change()
    AfficherApercu
End Sub

Private Sub AfficherApercu()
    Dim lundi As Date
    On Error Resume Next
    lundi = ObtenirLundi()
    On Error GoTo 0

    If lundi = 0 Then
        lblSemNum.Caption = "Date invalide"
        lblDates.Caption = ""
        Exit Sub
    End If

    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(lundi, 2)
    lblSemNum.Caption = "Semaine " & sem & " - " & Year(lundi)
    lblDates.Caption = "Du " & Format(lundi, "dddd dd/mm/yyyy") & _
                       " au " & Format(lundi + 6, "dddd dd/mm/yyyy")
End Sub

Private Function ObtenirLundi() As Date
    Dim refDate As Date

    If OptionSemCourante.Value Then
        refDate = Date
    ElseIf OptionSemSuivante.Value Then
        refDate = Date + 7
    Else
        If Not IsDate(txtDatePerso.Text) Then
            ObtenirLundi = 0: Exit Function
        End If
        refDate = CDate(txtDatePerso.Text)
    End If

    Dim wd As Integer
    wd = Weekday(refDate, vbMonday)
    ObtenirLundi = refDate - (wd - 1)
End Function

Private Sub cmdGenerer_Click()
    Dim lundi As Date
    On Error Resume Next
    lundi = ObtenirLundi()
    On Error GoTo 0

    If lundi = 0 Then
        MsgBox "Date invalide. Veuillez saisir une date au format jj/mm/aaaa.", vbExclamation
        Exit Sub
    End If

    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(lundi, 2)
    Dim rep As Integer
    rep = MsgBox("Générer le planning pour la semaine " & sem & " ?" & Chr(10) & _
                 "Du " & Format(lundi, "dd/mm/yyyy") & " au " & Format(lundi + 6, "dd/mm/yyyy"), _
                 vbYesNo + vbQuestion, "Confirmation")
    If rep = vbNo Then Exit Sub

    ' Définir la semaine cible dans le module principal
    g_LundiCible = lundi
    Me.Hide

    ' Lancer la génération
    Call GenererPlanning

    MsgBox "Planning Semaine " & sem & " généré avec succ�s !", vbInformation, "Terminé"
    Unload Me
End Sub

Private Sub cmdFermer_Click()
    Unload Me
End Sub


