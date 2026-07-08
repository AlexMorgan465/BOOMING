Attribute VB_Name = "ModulePlanningTf"

Option Explicit

'--------------------------------------------------------------------
' Structure regroupant toutes les infos calculees pour un agent / semaine
'--------------------------------------------------------------------
Public Type TF_AgentInfo
    Zone As String
    Nom As String
    EntreeH(1 To 7) As Integer
    SortieH(1 To 7) As Integer
    IsOff(1 To 7) As Boolean
    Comment(1 To 7) As String
    PauseDebut(1 To 7) As String   ' "" si OFF
    PauseFin(1 To 7) As String
    RowBDD As Long
    EquipeLabel As String
    JourRepoTournant As Integer
    JourCourt As Integer           ' 0 = pas de jour raccourci
End Type


'=====================================================================================
' GENERATEUR DE PLANNING - Projet "TF"  (v2)
'=====================================================================================
' CHANGEMENTS PAR RAPPORT A LA V1 :
'
' 1. Deux feuilles generees au lieu d'une :
'      - Feuille "<Projet> S<semaine>"       : planning des SHIFTS (Debut/Fin de shift)
'      - Feuille "<Projet> Pause S<semaine>" : planning des PAUSES DEJEUNER (DP/FP)
'
' 2. Separation Manager / Collaborateur, basee sur la colonne BDD "MANAGER" (OUI/NON) :
'      - Bloc "Collaborateurs" : lignes normales (agents avec MANAGER <> OUI)
'      - Bloc "Manager"        : lignes des agents avec MANAGER = OUI, affichees dans
'        un tableau separe en bas de chaque feuille, au format "detaille" (avec OFF,
'        NB heures planifiees, TT, Commentaires), comme dans ModulePlanningTlv.
'
' 3. Le tableau "Collaborateurs" n'affiche plus les colonnes OFF / NB heures / TT par
'    ligne. A la place, un RECAP JOURNALIER est ajoute en bas de la feuille SHIFTS :
'      NB Ouverture / NB Middle / NB Fermeture / NB OFF, calcule du MARDI au SAMEDI.
'      - "Ouverture" = shift 8h-18h
'      - "Fermeture" = shift 10h-20h
'      - "Middle"    = tout autre horaire travaille (ex : jour raccourci, cf point 4)
'
' 4. Journee raccourcie (1 jour/semaine/collaborateur) pour ajuster le total hebdo :
'      - 1 jour travaille (hors repos), choisi de facon tournante, est raccourci :
'          Equipe A (8h-18h)  -> ce jour-la : 8h-17h
'          Equipe B (10h-20h) -> ce jour-la : 10h-17h
'      Objectif vise par l'utilisateur : total hebdo ~44h. A NOTER : avec la regle
'      telle que demandee, la reduction n'est pas symetrique entre equipe A (-1h) et
'      equipe B (-3h) : a ajuster si besoin une fois la regle definitive validee.
'
' 5. FI / DAC / LEADS (colonnes "groupe" visibles sur le planning pause dejeuner
'    d'origine) : NON traites dans cette version (sujet mis de cote, traite
'    separement). Tous les agents suivent la regle Equipe A/B "generique".
'
' Priorite des regles (du + fort au + faible) : Contrat > Maladie > Conge > Repos
' hebdomadaire (dimanche + jour tournant) > Horaire equipe A/B (avec jour raccourci)
' puis annotation Teletravail si applicable et que le jour est travaille.
'=====================================================================================

Public Const NOM_FEUILLE_BDD As String = "BDD"
Public Const PART_EQUIPE_A As Double = 0.45   ' 45% des agents sur l'equipe 8h-18h

'--------------------------------------------------------------------
' POINT D'ENTREE
'--------------------------------------------------------------------
Sub GenererPlanningTF()

    Dim wsBDD As Worksheet, wsShift As Worksheet, wsPause As Worksheet
    Dim projectName As String, weekStartStr As String
    Dim weekStart As Date
    Dim lastRow As Long, r As Long
    Dim headers As Object

    On Error GoTo ErrHandler

    If Not SheetExists(NOM_FEUILLE_BDD) Then
        MsgBox "La feuille '" & NOM_FEUILLE_BDD & "' est introuvable.", vbCritical
        Exit Sub
    End If
    Set wsBDD = ThisWorkbook.Sheets(NOM_FEUILLE_BDD)

    projectName = InputBox("Nom du projet / de l'activite a generer (ex: TF) :", _
                            "Generation du planning", "TF")
    If Trim(projectName) = "" Then Exit Sub

    weekStartStr = InputBox("Date du LUNDI de la semaine a generer (jj/mm/aaaa) :", _
                             "Generation du planning", _
                             Format(Date - Weekday(Date, vbMonday) + 1, "dd/mm/yyyy"))
    If Trim(weekStartStr) = "" Then Exit Sub
    If Not IsDate(weekStartStr) Then
        MsgBox "Date invalide.", vbExclamation
        Exit Sub
    End If
    weekStart = CDate(weekStartStr)
    weekStart = weekStart - Weekday(weekStart, vbMonday) + 1 ' recale sur le lundi

    Set headers = GetHeaderMap(wsBDD)

    lastRow = wsBDD.Cells(wsBDD.Rows.Count, GetCol(headers, "MATRICULE")).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Aucune donnee trouvee dans la BDD.", vbExclamation
        Exit Sub
    End If

    Dim colActivite As Long, colManager As Long
    colActivite = GetCol(headers, "ACTIVITE")
    colManager = GetCol(headers, "MANAGER")

    ' --- Separation Collaborateurs / Manager (colonne BDD "MANAGER" = OUI/NON) -----
    Dim collabRows() As Long, nCollab As Long
    Dim managerRows() As Long, nManager As Long
    nCollab = 0: nManager = 0
    ReDim collabRows(1 To lastRow)
    ReDim managerRows(1 To lastRow)

    For r = 2 To lastRow
        Dim actVal As String, mgrVal As String
        actVal = Trim(wsBDD.Cells(r, colActivite).Value)
        If StrComp(actVal, projectName, vbTextCompare) = 0 Then
            mgrVal = Trim(wsBDD.Cells(r, colManager).Value)
            If StrComp(mgrVal, "OUI", vbTextCompare) = 0 Then
                nManager = nManager + 1
                managerRows(nManager) = r
            Else
                nCollab = nCollab + 1
                collabRows(nCollab) = r
            End If
        End If
    Next r

    If nCollab = 0 And nManager = 0 Then
        MsgBox "Aucune ligne trouvee pour l'activite '" & projectName & "' dans la BDD.", vbExclamation
        Exit Sub
    End If

    Dim weekNum As Long
    On Error Resume Next
    weekNum = Application.WorksheetFunction.IsoWeekNum(weekStart)
    On Error GoTo 0

    Dim nEquipeA As Long
    nEquipeA = CLng(Int(nCollab * PART_EQUIPE_A + 0.5))

    ' --- Pre-calcul des infos (equipe / repos / jour raccourci) + horaires jour par jour
    Dim collabInfo() As TF_AgentInfo
    ReDim collabInfo(1 To nCollab)
    Dim i As Long
    For i = 1 To nCollab
        collabInfo(i) = BuildTF_AgentInfo(wsBDD, headers, collabRows(i), weekStart, _
                                            i - 1, nCollab, nEquipeA, weekNum)
    Next i

    Dim managerInfo() As TF_AgentInfo
    ReDim managerInfo(1 To nManager)
    For i = 1 To nManager
        ' Les managers restent sur une equipe fixe "large" (existant deja dans la BDD) :
        ' on reutilise la meme mecanique Equipe A/B pour rester coherent, sans jour raccourci.
        managerInfo(i) = BuildTF_AgentInfo(wsBDD, headers, managerRows(i), weekStart, _
                                             i - 1, IIf(nManager = 0, 1, nManager), _
                                             IIf(nManager = 0, 0, nManager), weekNum, _
                                             applyShortDay:=False)
    Next i

    ' === Ecriture BDD (horaires reels) pour tous les agents (collab + manager) ========
    For i = 1 To nCollab
        WriteAgentHoursToBDD wsBDD, headers, collabInfo(i), weekStart
    Next i
    For i = 1 To nManager
        WriteAgentHoursToBDD wsBDD, headers, managerInfo(i), weekStart
    Next i

    ' === Feuille SHIFTS ================================================================
    Set wsShift = PreparePlanningSheet(projectName & " S" & weekNum)
    Dim outRow As Long
    outRow = WriteDayHeader(wsShift, 1, weekStart, "S" & weekNum, "Collaborateurs", True)
    For i = 1 To nCollab
        outRow = WriteShiftRow(wsShift, collabInfo(i), outRow)
    Next i

    outRow = outRow + 2
    outRow = WriteDailyRecap(wsShift, outRow, collabInfo, nCollab)

    If nManager > 0 Then
        outRow = outRow + 2
        outRow = WriteDayHeader(wsShift, outRow, weekStart, "Manager", "Manager", False)
        For i = 1 To nManager
            outRow = WriteManagerRow(wsShift, managerInfo(i), outRow)
        Next i
    End If
    wsShift.Columns.AutoFit

    ' === Feuille PAUSES =================================================================
    Set wsPause = PreparePlanningSheet(projectName & " Pause S" & weekNum)
    outRow = WriteDayHeader(wsPause, 1, weekStart, "Planning pause dejeuner", "Collaborateur", False, _
                             pauseMode:=True)
    For i = 1 To nCollab
        outRow = WritePauseRow(wsPause, collabInfo(i), outRow)
    Next i

    If nManager > 0 Then
        outRow = outRow + 1
        outRow = WriteDayHeader(wsPause, outRow, weekStart, "Manager", "Manager", False)
        For i = 1 To nManager
            outRow = WriteManagerRow(wsPause, managerInfo(i), outRow)
        Next i
    End If
    wsPause.Columns.AutoFit

    MsgBox "Planning genere avec succes." & vbCrLf & _
           "Feuille shifts : '" & wsShift.Name & "'" & vbCrLf & _
           "Feuille pauses : '" & wsPause.Name & "'" & vbCrLf & vbCrLf & _
           "Collaborateurs : " & nCollab & " (Equipe A : " & nEquipeA & _
           " / Equipe B : " & (nCollab - nEquipeA) & ")" & vbCrLf & _
           "Managers : " & nManager & vbCrLf & vbCrLf & _
           "Rappel : la logique FI/DAC/LEADS n'est pas traitee dans cette version " & _
           "(sujet mis de cote).", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub


'--------------------------------------------------------------------
' Calcule equipe / repos / jour raccourci / horaires jour par jour / pauses
' pour un agent donne, sur la semaine "weekStart".
'--------------------------------------------------------------------
Function BuildTF_AgentInfo(wsBDD As Worksheet, headers As Object, rowBDD As Long, _
                             weekStart As Date, ByVal pos As Long, ByVal total As Long, _
                             ByVal nEquipeA As Long, ByVal weekNum As Long, _
                             Optional ByVal applyShortDay As Boolean = True) As TF_AgentInfo

    Dim info As TF_AgentInfo
    info.RowBDD = rowBDD

    Dim nomComplet As String
    nomComplet = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "NOMCOMPLET")).Value)
    If nomComplet = "" Then
        nomComplet = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "NOM")).Value & " " & _
                           wsBDD.Cells(rowBDD, GetCol(headers, "PRENOM")).Value)
    End If
    info.Nom = nomComplet
    info.Zone = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "ZONES")).Value)

    Dim entreeDefaut As Integer, sortieDefaut As Integer, equipeLabel As String
    Dim jourRepoTournant As Integer, pauseLabelDummy As String
    AffecterEquipeEtRepos pos, total, nEquipeA, weekNum, entreeDefaut, sortieDefaut, _
                           equipeLabel, jourRepoTournant, pauseLabelDummy
    info.EquipeLabel = equipeLabel
    info.JourRepoTournant = jourRepoTournant

    Dim jourCourt As Integer
    jourCourt = 0
    If applyShortDay Then jourCourt = ChoisirJourCourt(pos, weekNum, jourRepoTournant)
    info.JourCourt = jourCourt

    Dim dayIndex As Integer
    For dayIndex = 1 To 7
        Dim dayDate As Date
        dayDate = weekStart + (dayIndex - 1)

        Dim d As Variant
        d = GetDayInfo(wsBDD, headers, rowBDD, dayDate, dayIndex, entreeDefaut, sortieDefaut, _
                        jourRepoTournant, jourCourt)

        info.EntreeH(dayIndex) = d(0)
        info.SortieH(dayIndex) = d(1)
        info.IsOff(dayIndex) = d(2)
        info.Comment(dayIndex) = d(3)

        If info.IsOff(dayIndex) Then
            info.PauseDebut(dayIndex) = "OFF"
            info.PauseFin(dayIndex) = "OFF"
        Else
            Dim pd As String, pf As String
            CalculerPause info.EntreeH(dayIndex), pos, weekNum, pd, pf
            info.PauseDebut(dayIndex) = pd
            info.PauseFin(dayIndex) = pf
        End If
    Next dayIndex

    BuildTF_AgentInfo = info
End Function

'--------------------------------------------------------------------
' Choisit, parmi les jours travailles (hors jour de repos tournant et
' hors dimanche), un jour "raccourci" tournant selon pos+weekNum.
'--------------------------------------------------------------------
Function ChoisirJourCourt(ByVal pos As Long, ByVal weekNum As Long, _
                           ByVal jourRepoTournant As Integer) As Integer
    Dim candidats() As Integer
    ReDim candidats(1 To 5)
    Dim n As Integer, d As Integer
    n = 0
    For d = 1 To 6 ' Lundi a Samedi (Dimanche jamais travaille de toute facon)
        If d <> jourRepoTournant Then
            n = n + 1
            candidats(n) = d
        End If
    Next d
    ChoisirJourCourt = candidats(((pos + weekNum) Mod n) + 1)
End Function

'--------------------------------------------------------------------
' Determine le creneau de pause dejeuner (1h) en fonction de l'heure
' d'entree du jour (8h -> pause A, 10h -> pause B), en alternant selon
' l'agent/la semaine pour etaler les pauses.
'--------------------------------------------------------------------
Sub CalculerPause(ByVal entreeH As Integer, ByVal pos As Long, ByVal weekNum As Long, _
                   ByRef pauseDebut As String, ByRef pauseFin As String)
    Dim alt As Boolean
    alt = ((pos + weekNum) Mod 2 = 0)
    Select Case entreeH
        Case 8
            If alt Then pauseDebut = "12:00": pauseFin = "13:00" Else pauseDebut = "13:00": pauseFin = "14:00"
        Case 10
            If alt Then pauseDebut = "14:00": pauseFin = "15:00" Else pauseDebut = "15:00": pauseFin = "16:00"
        Case Else
            ' Horaire non standard (ex : futur groupe FI/DAC/LEADS) -> pause par defaut
            pauseDebut = "13:00": pauseFin = "14:00"
    End Select
End Sub

'--------------------------------------------------------------------
' Equipe (A=8h-18h / B=10h-20h) + jour de repos tournant, bases sur la
' position de l'agent dans la liste filtree et le numero de semaine ISO.
'--------------------------------------------------------------------
Sub AffecterEquipeEtRepos(ByVal pos As Long, ByVal total As Long, ByVal nEquipeA As Long, _
                          ByVal weekNum As Long, ByRef entreeH As Integer, _
                          ByRef sortieH As Integer, ByRef equipeLabel As String, _
                          ByRef jourRepoTournant As Integer, ByRef pauseLabel As String)

    Dim posRotatif As Long
    If total <= 0 Then total = 1
    posRotatif = (pos + weekNum) Mod total

    If posRotatif < nEquipeA Then
        entreeH = 8: sortieH = 18: equipeLabel = "8h-18h"
    Else
        entreeH = 10: sortieH = 20: equipeLabel = "10h-20h"
    End If
    pauseLabel = "" ' non utilise ici (cf CalculerPause, recalcule par jour)

    ' Jour de repos tournant : UNIQUEMENT entre Mercredi (dayIndex=3) et Samedi (=6).
    Select Case (pos + weekNum) Mod 4
        Case 0: jourRepoTournant = 3 ' Mercredi
        Case 1: jourRepoTournant = 4 ' Jeudi
        Case 2: jourRepoTournant = 5 ' Vendredi
        Case Else: jourRepoTournant = 6 ' Samedi
    End Select
End Sub

'--------------------------------------------------------------------
' Calcule l'horaire/l'etat d'un jour donne pour une ligne BDD, en tenant
' compte du jour raccourci ("jourCourt") : ce jour-la, la sortie est
' avancee (Equipe A : 18h->17h / Equipe B : 20h->17h), pour ajuster le
' total hebdomadaire.
' Retourne un tableau : (entreeHeure, sortieHeure, isOff, commentaire)
'--------------------------------------------------------------------
Function GetDayInfo(wsBDD As Worksheet, headers As Object, rowBDD As Long, _
                     dayDate As Date, dayIndex As Integer, _
                     entreeDefaut As Integer, sortieDefaut As Integer, _
                     jourRepoTournant As Integer, jourCourt As Integer) As Variant

    Dim entreeH As Integer, sortieH As Integer, isOff As Boolean, comment As String
    isOff = False
    comment = ""

    ' 1) Contrat / statut ------------------------------------------------
    Dim dEmbauche As Variant, dSortie As Variant, typeContrat As String
    dEmbauche = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDEMBAUCHE")).Value
    dSortie = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDESORTIE")).Value
    typeContrat = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TYPEDECONTRAT")).Value)

    If IsDate(dEmbauche) Then
        If dayDate < CDate(dEmbauche) Then
            isOff = True: comment = "Pas encore embauche"
        End If
    End If
    If Not isOff And IsDate(dSortie) Then
        If dayDate >= CDate(dSortie) Then
            isOff = True: comment = "Contrat termine"
        End If
    End If
    If Not isOff And (StrComp(typeContrat, "Termine", vbTextCompare) = 0 _
                       Or StrComp(typeContrat, "Sorti", vbTextCompare) = 0) Then
        isOff = True: comment = "Contrat termine"
    End If

    ' 2) Maladie ----------------------------------------------------------
    If Not isOff Then
        Dim maladieVal As String, dArret As Variant, dRepr As Variant
        maladieVal = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "MALADIE")).Value)
        dArret = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDARRET")).Value
        dRepr = wsBDD.Cells(rowBDD, GetCol(headers, "DATEDEREPRISE")).Value
        If maladieVal <> "" And StrComp(maladieVal, "NON", vbTextCompare) <> 0 Then
            Dim okStartM As Boolean, okEndM As Boolean
            okStartM = (Not IsDate(dArret)) Or (dayDate >= CDate(dArret))
            okEndM = (Not IsDate(dRepr)) Or (dayDate <= CDate(dRepr))
            If okStartM And okEndM Then
                isOff = True: comment = "Maladie"
            End If
        End If
    End If

    ' 3) Conge --------------------------------------------------------------
    If Not isOff Then
        Dim congeVal As String, cD As Variant, cF As Variant, typeConge As String
        congeVal = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "CONGE")).Value)
        cD = wsBDD.Cells(rowBDD, GetCol(headers, "CONGED")).Value
        cF = wsBDD.Cells(rowBDD, GetCol(headers, "CONGEF")).Value
        typeConge = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TYPEDECONGE")).Value)
        If congeVal <> "" And StrComp(congeVal, "NON", vbTextCompare) <> 0 Then
            Dim okStartC As Boolean, okEndC As Boolean
            okStartC = (Not IsDate(cD)) Or (dayDate >= CDate(cD))
            okEndC = (Not IsDate(cF)) Or (dayDate <= CDate(cF))
            If okStartC And okEndC Then
                isOff = True
                comment = IIf(typeConge <> "", typeConge, "Conge")
            End If
        End If
    End If

    ' 4) Repos hebdomadaire par defaut : Dimanche fixe + jour tournant -------
    If Not isOff Then
        If dayIndex = 7 Then
            isOff = True
            comment = "Repos hebdomadaire (Dimanche)"
        ElseIf dayIndex = jourRepoTournant Then
            isOff = True
            comment = "Repos hebdomadaire"
        Else
            entreeH = entreeDefaut
            sortieH = sortieDefaut
            ' Jour raccourci : sortie avancee (regle demandee : 8h-18h -> 8h-17h ;
            ' 10h-20h -> 10h-17h), pour ajuster le total hebdomadaire.
            If dayIndex = jourCourt Then
                sortieH = 17
                If comment = "" Then comment = "Journee raccourcie"
            End If
        End If
        If comment = "" Then comment = "RAS"
    Else
        If comment = "" Then comment = "RAS"
    End If

    ' 5) Teletravail : annotation seule, horaire inchange ---------------------
    If Not isOff Then
        Dim ttVal As String, ttD As Variant, ttF As Variant
        ttVal = Trim(wsBDD.Cells(rowBDD, GetCol(headers, "TT")).Value)
        ttD = wsBDD.Cells(rowBDD, GetCol(headers, "TTD")).Value
        ttF = wsBDD.Cells(rowBDD, GetCol(headers, "TTF")).Value
        If ttVal <> "" And StrComp(ttVal, "NON", vbTextCompare) <> 0 Then
            Dim okStartT As Boolean, okEndT As Boolean
            okStartT = (Not IsDate(ttD)) Or (dayDate >= CDate(ttD))
            okEndT = (Not IsDate(ttF)) Or (dayDate <= CDate(ttF))
            If okStartT And okEndT Then
                If StrComp(comment, "RAS", vbTextCompare) = 0 Then
                    comment = "Teletravail"
                Else
                    comment = comment & " / Teletravail"
                End If
            End If
        End If
    End If

    GetDayInfo = Array(entreeH, sortieH, isOff, comment)
End Function

'--------------------------------------------------------------------
' Classe un jour travaille : "Ouverture" (8h-18h) / "Fermeture" (10h-20h)
' / "Middle" (tout autre horaire, ex : jour raccourci).
'--------------------------------------------------------------------
Function ClassifierShift(ByVal entreeH As Integer, ByVal sortieH As Integer) As String
    If entreeH = 8 And sortieH = 18 Then
        ClassifierShift = "Ouverture"
    ElseIf entreeH = 10 And sortieH = 20 Then
        ClassifierShift = "Fermeture"
    Else
        ClassifierShift = "Middle"
    End If
End Function

'--------------------------------------------------------------------
' Ecrit les horaires calcules dans la BDD (colonnes LUN Entree...DIM Sortie)
'--------------------------------------------------------------------
Sub WriteAgentHoursToBDD(wsBDD As Worksheet, headers As Object, info As TF_AgentInfo, weekStart As Date)
    Dim dayIndex As Integer
    For dayIndex = 1 To 7
        Dim colEntreeBDD As Long, colSortieBDD As Long
        colEntreeBDD = GetCol(headers, DayColKey(dayIndex, True))
        colSortieBDD = GetCol(headers, DayColKey(dayIndex, False))

        If info.IsOff(dayIndex) Then
            wsBDD.Cells(info.RowBDD, colEntreeBDD).Value = "OFF"
            wsBDD.Cells(info.RowBDD, colSortieBDD).Value = "OFF"
        Else
            wsBDD.Cells(info.RowBDD, colEntreeBDD).Value = TimeSerial(info.EntreeH(dayIndex), 0, 0)
            wsBDD.Cells(info.RowBDD, colEntreeBDD).NumberFormat = "hh""H"""
            wsBDD.Cells(info.RowBDD, colSortieBDD).Value = TimeSerial(info.SortieH(dayIndex), 0, 0)
            wsBDD.Cells(info.RowBDD, colSortieBDD).NumberFormat = "hh""H"""
        End If
    Next dayIndex
End Sub

'--------------------------------------------------------------------
' Ecrit une ligne "collaborateur" dans la feuille SHIFTS
' (Zones | Nom | 7x(Debut,Fin) | Commentaires)
'--------------------------------------------------------------------
Function WriteShiftRow(ws As Worksheet, info As TF_AgentInfo, outRow As Long) As Long
    ws.Cells(outRow, 1).Value = info.Zone
    ws.Cells(outRow, 2).Value = info.Nom
    ws.Cells(outRow, 1).Font.Bold = True
    ws.Cells(outRow, 2).Font.Bold = True

    Dim dayIndex As Integer, comments As Object
    Set comments = CreateObject("Scripting.Dictionary")

    For dayIndex = 1 To 7
        Dim c1 As Long, c2 As Long
        c1 = 3 + (dayIndex - 1) * 2: c2 = c1 + 1
        If info.IsOff(dayIndex) Then
            ws.Cells(outRow, c1).Value = "OFF"
            ws.Cells(outRow, c2).Value = "OFF"
            ws.Range(ws.Cells(outRow, c1), ws.Cells(outRow, c2)).Font.Color = RGB(200, 0, 0)
        Else
            ws.Cells(outRow, c1).Value = TimeSerial(info.EntreeH(dayIndex), 0, 0)
            ws.Cells(outRow, c1).NumberFormat = "h:mm"
            ws.Cells(outRow, c2).Value = TimeSerial(info.SortieH(dayIndex), 0, 0)
            ws.Cells(outRow, c2).NumberFormat = "h:mm"
        End If
        If info.Comment(dayIndex) <> "" And StrComp(info.Comment(dayIndex), "RAS", vbTextCompare) <> 0 Then
            If Not comments.Exists(info.Comment(dayIndex)) Then comments.Add info.Comment(dayIndex), True
        End If
    Next dayIndex

    Dim k As Variant, txt As String
    If comments.Count = 0 Then
        txt = "RAS"
    Else
        For Each k In comments.Keys
            txt = txt & IIf(txt = "", "", " / ") & k
        Next k
    End If
    ws.Cells(outRow, 17).Value = txt

    WriteShiftRow = outRow + 1
End Function

'--------------------------------------------------------------------
' Ecrit une ligne "collaborateur" dans la feuille PAUSES
' (Zones | Nom | 7x(DP,FP), pas de colonne Commentaires)
'--------------------------------------------------------------------
Function WritePauseRow(ws As Worksheet, info As TF_AgentInfo, outRow As Long) As Long
    ws.Cells(outRow, 1).Value = info.Zone
    ws.Cells(outRow, 2).Value = info.Nom
    ws.Cells(outRow, 1).Font.Bold = True
    ws.Cells(outRow, 2).Font.Bold = True

    Dim dayIndex As Integer
    For dayIndex = 1 To 7
        Dim c1 As Long, c2 As Long
        c1 = 3 + (dayIndex - 1) * 2: c2 = c1 + 1
        ws.Cells(outRow, c1).Value = info.PauseDebut(dayIndex)
        ws.Cells(outRow, c2).Value = info.PauseFin(dayIndex)
        If info.IsOff(dayIndex) Then
            ws.Range(ws.Cells(outRow, c1), ws.Cells(outRow, c2)).Font.Color = RGB(200, 0, 0)
        End If
    Next dayIndex

    WritePauseRow = outRow + 1
End Function

'--------------------------------------------------------------------
' Ecrit une ligne "Manager" (format detaille, identique pour les 2 feuilles) :
' Zones | Nom | 7x(Debut,Fin) | OFF | NB heures planifiees | TT | Commentaires
'--------------------------------------------------------------------
Function WriteManagerRow(ws As Worksheet, info As TF_AgentInfo, outRow As Long) As Long
    ws.Cells(outRow, 1).Value = info.Zone
    ws.Cells(outRow, 2).Value = info.Nom
    ws.Cells(outRow, 1).Font.Bold = True
    ws.Cells(outRow, 2).Font.Bold = True

    Dim dayIndex As Integer, offCount As Long, totalHeures As Double
    Dim comments As Object
    Set comments = CreateObject("Scripting.Dictionary")
    offCount = 0: totalHeures = 0

    For dayIndex = 1 To 7
        Dim c1 As Long, c2 As Long
        c1 = 3 + (dayIndex - 1) * 2: c2 = c1 + 1
        If info.IsOff(dayIndex) Then
            ws.Cells(outRow, c1).Value = "OFF"
            ws.Cells(outRow, c2).Value = "OFF"
            ws.Range(ws.Cells(outRow, c1), ws.Cells(outRow, c2)).Font.Color = RGB(200, 0, 0)
            offCount = offCount + 1
        Else
            ws.Cells(outRow, c1).Value = TimeSerial(info.EntreeH(dayIndex), 0, 0)
            ws.Cells(outRow, c1).NumberFormat = "h:mm"
            ws.Cells(outRow, c2).Value = TimeSerial(info.SortieH(dayIndex), 0, 0)
            ws.Cells(outRow, c2).NumberFormat = "h:mm"
            totalHeures = totalHeures + (info.SortieH(dayIndex) - info.EntreeH(dayIndex) - 1)
        End If
        If info.Comment(dayIndex) <> "" And StrComp(info.Comment(dayIndex), "RAS", vbTextCompare) <> 0 Then
            If Not comments.Exists(info.Comment(dayIndex)) Then comments.Add info.Comment(dayIndex), True
        End If
    Next dayIndex

    ws.Cells(outRow, 17).Value = offCount
    ws.Cells(outRow, 18).Value = totalHeures / 24
    ws.Cells(outRow, 18).NumberFormat = "[h]:mm:ss"
    ws.Cells(outRow, 19).Value = "N"

    Dim k As Variant, txt As String
    If comments.Count = 0 Then
        txt = "RAS"
    Else
        For Each k In comments.Keys
            txt = txt & IIf(txt = "", "", " / ") & k
        Next k
    End If
    ws.Cells(outRow, 20).Value = txt

    WriteManagerRow = outRow + 1
End Function

'--------------------------------------------------------------------
' Recap journalier (Mardi -> Samedi) pour les COLLABORATEURS uniquement :
' NB Ouverture / NB Middle / NB Fermeture / NB OFF
'--------------------------------------------------------------------
Function WriteDailyRecap(ws As Worksheet, startRow As Long, collabInfo() As TF_AgentInfo, _
                          ByVal nCollab As Long) As Long
    Dim labels As Variant
    labels = Array("Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi") ' dayIndex 2..6

    Dim r As Long
    r = startRow

    ' En-tete
    ws.Cells(r, 2).Value = ""
    Dim c As Integer
    For c = 0 To 4
        ws.Cells(r, 3 + c).Value = labels(c)
    Next c
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 7))
        .Interior.Color = RGB(217, 226, 243)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    Dim rowNames As Variant
    rowNames = Array("NB Ouverture", "NB Middle", "NB Fermeture", "NB OFF")
    Dim ri As Integer
    For ri = 0 To 3
        ws.Cells(r + 1 + ri, 2).Value = rowNames(ri)
        ws.Cells(r + 1 + ri, 2).Font.Bold = True
    Next ri

    For c = 0 To 4 ' Mardi(dayIndex=2) a Samedi(dayIndex=6)
        Dim dayIndex As Integer
        dayIndex = c + 2
        Dim nOuv As Long, nMid As Long, nFerm As Long, nOff As Long
        nOuv = 0: nMid = 0: nFerm = 0: nOff = 0

        Dim i As Long
        For i = 1 To nCollab
            If collabInfo(i).IsOff(dayIndex) Then
                nOff = nOff + 1
            Else
                Select Case ClassifierShift(collabInfo(i).EntreeH(dayIndex), collabInfo(i).SortieH(dayIndex))
                    Case "Ouverture": nOuv = nOuv + 1
                    Case "Fermeture": nFerm = nFerm + 1
                    Case Else: nMid = nMid + 1
                End Select
            End If
        Next i

        ws.Cells(r + 1, 3 + c).Value = nOuv
        ws.Cells(r + 2, 3 + c).Value = nMid
        ws.Cells(r + 3, 3 + c).Value = nFerm
        ws.Cells(r + 4, 3 + c).Value = nOff
    Next c

    With ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + 4, 7))
        .HorizontalAlignment = xlCenter
    End With

    WriteDailyRecap = r + 6
End Function

'--------------------------------------------------------------------
' Ecrit les 2 lignes d'en-tete (titre bloc + libelles colonnes) d'un bloc
' (Collaborateurs ou Manager) sur une feuille (Shift ou Pause).
'   - shiftMode = True  -> libelles "Debut de shift/Fin de shift" + colonne Commentaires
'   - shiftMode = False & pauseMode = True -> libelles "DP/FP", pas de Commentaires
'   - shiftMode = False & pauseMode = False -> bloc Manager (detaille, comme avant)
'--------------------------------------------------------------------
Function WriteDayHeader(ws As Worksheet, startRow As Long, weekStart As Date, _
                         titleCell As String, roleLabel As String, _
                         ByVal shiftMode As Boolean, _
                         Optional ByVal pauseMode As Boolean = False) As Long

    Dim r1 As Long, r2 As Long
    r1 = startRow: r2 = startRow + 1

    Dim headerFill As Long, headerFont As Long
    headerFill = RGB(31, 73, 125)
    headerFont = RGB(255, 255, 255)

    With ws.Range(ws.Cells(r1, 1), ws.Cells(r1, 2))
        .Merge
        .Value = titleCell
        .HorizontalAlignment = xlCenter
        .Interior.Color = headerFill
        .Font.Color = headerFont
        .Font.Bold = True
    End With

    Dim dayIndex As Integer
    For dayIndex = 1 To 7
        Dim c1 As Long, c2 As Long
        c1 = 3 + (dayIndex - 1) * 2: c2 = c1 + 1
        Dim dayDate As Date
        dayDate = weekStart + (dayIndex - 1)
        With ws.Range(ws.Cells(r1, c1), ws.Cells(r1, c2))
            .Merge
            .Value = DayLabel(dayIndex) & " " & Format(dayDate, "dd mmmm yyyy")
            .HorizontalAlignment = xlCenter
            .Interior.Color = headerFill
            .Font.Color = headerFont
            .Font.Bold = True
        End With
        If pauseMode Then
            ws.Cells(r2, c1).Value = "DP"
            ws.Cells(r2, c2).Value = "FP"
        Else
            ws.Cells(r2, c1).Value = "Debut de shift"
            ws.Cells(r2, c2).Value = "Fin de shift"
        End If
    Next dayIndex

    ws.Cells(r2, 1).Value = "Zones"
    ws.Cells(r2, 2).Value = roleLabel

    Dim lastCol As Long
    If shiftMode Then
        ws.Cells(r2, 17).Value = "Commentaires"
        lastCol = 17
    ElseIf pauseMode Then
        lastCol = 16
    Else ' bloc Manager detaille
        ws.Cells(r2, 17).Value = "OFF"
        ws.Cells(r2, 18).Value = "NB heures planifiees"
        ws.Cells(r2, 19).Value = "TT"
        ws.Cells(r2, 20).Value = "Commentaires"
        lastCol = 20
    End If

    With ws.Range(ws.Cells(r2, 1), ws.Cells(r2, lastCol))
        .Interior.Color = RGB(217, 226, 243)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    WriteDayHeader = r2 + 1
End Function

'--------------------------------------------------------------------
' Cle de colonne BDD normalisee pour un jour/sens donne (ex: "LUNENTREE")
'--------------------------------------------------------------------
Function DayColKey(dayIndex As Integer, isEntree As Boolean) As String
    Dim prefixes As Variant
    prefixes = Array("LUN", "MAR", "MER", "JEU", "VEN", "SAM", "DIM")
    DayColKey = prefixes(dayIndex - 1) & IIf(isEntree, "ENTREE", "SORTIE")
End Function

Function DayLabel(dayIndex As Integer) As String
    Dim labels As Variant
    labels = Array("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche")
    DayLabel = labels(dayIndex - 1)
End Function

'--------------------------------------------------------------------
' Prepare (recree) une feuille au nom donne
'--------------------------------------------------------------------
Function PreparePlanningSheet(ByVal sheetNameWanted As String) As Worksheet
    Dim sheetName As String
    sheetName = CleanSheetName(sheetNameWanted)

    Application.DisplayAlerts = False
    If SheetExists(sheetName) Then
        ThisWorkbook.Sheets(sheetName).Delete
    End If
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName
    Set PreparePlanningSheet = ws
End Function

Function CleanSheetName(ByVal s As String) As String
    Dim r As String
    r = s
    Dim badChars As Variant, ch As Variant
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For Each ch In badChars
        r = Replace(r, ch, "")
    Next ch
    If Len(r) > 31 Then r = Left(r, 31)
    CleanSheetName = r
End Function

Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'--------------------------------------------------------------------
' Construit un dictionnaire {en-tete normalise -> numero de colonne}
' en lisant la ligne 1 de la feuille BDD.
'--------------------------------------------------------------------
Function GetHeaderMap(wsBDD As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long, c As Long
    lastCol = wsBDD.Cells(1, wsBDD.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        Dim key As String
        key = NormalizeHeader(CStr(wsBDD.Cells(1, c).Value))
        If key <> "" And Not dict.Exists(key) Then
            dict.Add key, c
        End If
    Next c

    Set GetHeaderMap = dict
End Function

Function GetCol(headers As Object, ByVal key As String) As Long
    Dim normKey As String
    normKey = NormalizeHeader(key)
    If headers.Exists(normKey) Then
        GetCol = headers(normKey)
    Else
        Dim k As Variant, listeEntetes As String
        For Each k In headers.Keys
            listeEntetes = listeEntetes & IIf(listeEntetes = "", "", ", ") & k
        Next k
        Err.Raise vbObjectError + 1, , "Colonne introuvable dans la BDD pour la cle : " & key & _
                  " (normalisee : " & normKey & ")." & vbCrLf & _
                  "En-tetes detectes en ligne 1 de la BDD : " & listeEntetes
    End If
End Function

'--------------------------------------------------------------------
' Normalise un en-tete : majuscules, sans accents, sans espaces/points/apostrophes
'--------------------------------------------------------------------
Function NormalizeHeader(ByVal s As String) As String
    Dim r As String
    r = UCase(Trim(s))
    r = Replace(r, "É", "E"): r = Replace(r, "È", "E"): r = Replace(r, "Ê", "E"): r = Replace(r, "Ë", "E")
    r = Replace(r, "À", "A"): r = Replace(r, "Â", "A")
    r = Replace(r, "Ô", "O")
    r = Replace(r, "Î", "I"): r = Replace(r, "Ï", "I")
    r = Replace(r, "Ù", "U"): r = Replace(r, "Û", "U")
    r = Replace(r, "Ç", "C")
    r = Replace(r, ".", "")
    r = Replace(r, "'", "")
    r = Replace(r, "-", "")
    r = Replace(r, " ", "")
    NormalizeHeader = r
End Function
