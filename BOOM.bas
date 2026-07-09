Sub ApplyScheduling(ws As Worksheet, weekNum As Integer)
    ' This subroutine now supports auto-rotation if input tables are empty.
    ' It will infer shift types and OFF days from the source sheet (previous week)
    ' and apply a rotation to generate a varied schedule.

    Dim inputSheet As Worksheet
    On Error Resume Next
    Set inputSheet = ThisWorkbook.Worksheets("PlanningInputs")
    On Error GoTo 0

    Dim useAutoRotate As Boolean
    useAutoRotate = (inputSheet Is Nothing) Or (Application.CountA(inputSheet.Range("A:A")) = 0)

    ' Locate main roster data
    Dim mainStartRow As Long
    mainStartRow = FindRow(ws, "Collaborateur", 3) + 1
    If mainStartRow = 1 Then
        MsgBox "Main roster header not found.", vbCritical
        Exit Sub
    End If

    Dim mainEndRow As Long
    mainEndRow = mainStartRow
    Do While ws.Cells(mainEndRow, 3).Value <> ""
        mainEndRow = mainEndRow + 1
    Loop
    mainEndRow = mainEndRow - 1

    ' Prepare dictionaries
    Dim shiftDict As Object, offDict As Object, absDict As Object
    Set shiftDict = CreateObject("Scripting.Dictionary")
    Set offDict = CreateObject("Scripting.Dictionary")
    Set absDict = CreateObject("Scripting.Dictionary")

    If useAutoRotate Then
        ' ---- AUTO-ROTATION: infer from the source sheet (ws is the new copy, but we need the *source* before it was copied)
        ' Since we are in the new sheet, we need to get the source sheet name. We stored it earlier? We can pass it.
        ' For simplicity, we will assume the source sheet is the one before ws in the sheet order.
        Dim srcSheet As Worksheet
        Set srcSheet = ws.Previous  ' This might not be reliable if sheets reordered; better to store name during generation.
        ' We'll modify the main GenerateNextWeek to pass the srcSheet as a global or parameter.
        ' For now, we'll use a workaround: we'll look for the sheet with name "S" & Format(weekNum - 1, "00")
        Dim prevWeekCode As String
        prevWeekCode = "S" & Format(weekNum - 1, "00")
        On Error Resume Next
        Set srcSheet = ThisWorkbook.Worksheets(prevWeekCode)
        On Error GoTo 0
        If srcSheet Is Nothing Then
            MsgBox "Cannot find previous week sheet for auto-rotation. Please provide inputs.", vbExclamation
            Exit Sub
        End If

        ' Infer shift types and OFF days from srcSheet
        Dim srcMainStart As Long
        srcMainStart = FindRow(srcSheet, "Collaborateur", 3) + 1
        If srcMainStart = 1 Then
            MsgBox "Cannot find roster in source sheet.", vbCritical
            Exit Sub
        End If
        Dim srcRow As Long
        Dim name As String
        Dim dayCol As Long
        Dim startVal As Variant
        Dim matinCount As Integer, soirCount As Integer
        Dim offDays As Collection

        For srcRow = srcMainStart To srcMainStart + (mainEndRow - mainStartRow)
            name = Trim(srcSheet.Cells(srcRow, 3).Value)
            If name = "" Then Exit For
            ' Count start times for days 1-6 (Mon-Sat) to determine shift type
            matinCount = 0: soirCount = 0
            For dayCol = 4 To 14 Step 2   ' D to N (Mon-Sat)
                startVal = srcSheet.Cells(srcRow, dayCol).Value
                If IsNumeric(startVal) Then
                    If startVal >= TimeValue("07:00") And startVal <= TimeValue("08:30") Then
                        matinCount = matinCount + 1
                    ElseIf startVal >= TimeValue("10:00") And startVal <= TimeValue("11:30") Then
                        soirCount = soirCount + 1
                    End If
                End If
            Next dayCol
            ' Decide type
            If matinCount > soirCount Then
                shiftDict(name) = "soir"   ' flip to soir for next week
            ElseIf soirCount > matinCount Then
                shiftDict(name) = "matin"  ' flip to matin
            Else
                ' If tie, default to matin? We'll keep previous? We'll set to matin.
                shiftDict(name) = "matin"
            End If

            ' Infer OFF days: find days where both start and end are "OFF" (or "Congé", etc.)
            Set offDays = New Collection
            For dayCol = 4 To 16 Step 2   ' D to P (Mon-Sun)
                If srcSheet.Cells(srcRow, dayCol).Value = "OFF" And srcSheet.Cells(srcRow, dayCol + 1).Value = "OFF" Then
                    ' This day is OFF. Rotate by +1 day: (dayCol-4)/2 + 1 gives 1..7, then add 1 mod 7
                    Dim dayIndex As Integer
                    dayIndex = (dayCol - 4) / 2 + 1
                    dayIndex = dayIndex + 1
                    If dayIndex > 7 Then dayIndex = dayIndex - 7
                    offDays.Add dayIndex
                End If
            Next dayCol
            ' If we found two OFF days, use them; otherwise, default to Monday+Tuesday or Tuesday+Wednesday? We'll use found ones.
            If offDays.Count < 2 Then
                ' Fallback: set OFF to Monday and Tuesday
                offDays.Add 1
                offDays.Add 2
            End If
            Set offDict(name) = offDays
        Next srcRow

        ' Absences: we could carry them from srcSheet, but we'll skip for auto-rotate (assume no absences unless specified in inputs).
        ' If you want to carry absences, we could detect "Congé", "Maladie", etc. and copy them.

    Else
        ' ---- MANUAL INPUT MODE ----
        ' Read shift types (columns A-B)
        Dim lastRow As Long
        lastRow = inputSheet.Cells(inputSheet.Rows.Count, 1).End(xlUp).Row
        Dim i As Long
        For i = 2 To lastRow
            name = Trim(inputSheet.Cells(i, 1).Value)
            If name <> "" Then
                shiftDict(name) = Trim(inputSheet.Cells(i, 2).Value)
            End If
        Next i

        ' Read OFF days (columns E-G)
        Dim offRow As Long
        offRow = 2
        Do While inputSheet.Cells(offRow, 5).Value <> ""
            name = Trim(inputSheet.Cells(offRow, 5).Value)
            If name <> "" Then
                Dim offDaysCol As Collection
                Set offDaysCol = New Collection
                Dim d1 As Variant, d2 As Variant
                d1 = inputSheet.Cells(offRow, 6).Value
                If IsNumeric(d1) And d1 >= 1 And d1 <= 7 Then offDaysCol.Add CInt(d1)
                d2 = inputSheet.Cells(offRow, 7).Value
                If IsNumeric(d2) And d2 >= 1 And d2 <= 7 Then offDaysCol.Add CInt(d2)
                Set offDict(name) = offDaysCol
            End If
            offRow = offRow + 1
        Loop

        ' Read absences (columns H-K)
        Dim absRow As Long
        absRow = 2
        Do While inputSheet.Cells(absRow, 8).Value <> ""
            name = Trim(inputSheet.Cells(absRow, 8).Value)
            If name <> "" Then
                Dim absCode As String
                absCode = Trim(inputSheet.Cells(absRow, 9).Value)
                Dim startDate As Date
                startDate = inputSheet.Cells(absRow, 10).Value
                Dim endDate As Variant
                endDate = inputSheet.Cells(absRow, 11).Value
                If Not absDict.Exists(name) Then Set absDict(name) = New Collection
                absDict(name).Add Array(absCode, startDate, endDate)
            End If
            absRow = absRow + 1
        Loop
    End If

    ' ---- Apply schedule to the new sheet ----
    Dim mondayDate As Date
    mondayDate = ws.Range("D2").Value
    If Not IsDate(mondayDate) Then
        MsgBox "Invalid Monday date in D2.", vbCritical
        Exit Sub
    End If

    ' Define base times
    Dim baseStartMatin As Date, baseEndMatin As Date
    Dim baseStartSoir As Date, baseEndSoir As Date
    baseStartMatin = TimeValue("07:00:00")
    baseEndMatin = TimeValue("17:00:00")
    baseStartSoir = TimeValue("10:00:00")
    baseEndSoir = TimeValue("20:00:00")

    Dim row As Long
    Dim personName As String
    Dim shiftType As String
    Dim offDaysSet As Collection
    Dim absencesCol As Collection
    Dim dayIndex As Integer, col As Long
    Dim currentDate As Date
    Dim startTime As Date, endTime As Date
    Dim isOff As Boolean, isAbsent As Boolean
    Dim absenceCode As String
    Dim workingCountDict As Object
    Set workingCountDict = CreateObject("Scripting.Dictionary")

    For row = mainStartRow To mainEndRow
        personName = Trim(ws.Cells(row, 3).Value)
        If personName = "" Then Exit For

        shiftType = ""
        If shiftDict.Exists(personName) Then shiftType = LCase(Trim(shiftDict(personName)))

        Set offDaysSet = New Collection
        If offDict.Exists(personName) Then Set offDaysSet = offDict(personName)

        Set absencesCol = New Collection
        If absDict.Exists(personName) Then Set absencesCol = absDict(personName)

        workingCountDict(personName) = 0

        For dayIndex = 1 To 7
            col = 4 + (dayIndex - 1) * 2
            currentDate = mondayDate + dayIndex - 1

            isOff = False
            Dim offDay As Variant
            For Each offDay In offDaysSet
                If offDay = dayIndex Then isOff = True: Exit For
            Next offDay

            isAbsent = False
            absenceCode = ""
            Dim absItem As Variant
            For Each absItem In absencesCol
                Dim absStart As Date, absEnd As Variant
                absStart = absItem(1)
                absEnd = absItem(2)
                If IsDate(absEnd) Then
                    If currentDate >= absStart And currentDate <= absEnd Then
                        isAbsent = True: absenceCode = absItem(0): Exit For
                    End If
                ElseIf absEnd = "ongoing" Or absEnd = "" Then
                    If currentDate >= absStart Then
                        isAbsent = True: absenceCode = absItem(0): Exit For
                    End If
                End If
            Next absItem

            If isAbsent Then
                ws.Cells(row, col).Value = absenceCode
                ws.Cells(row, col + 1).Value = absenceCode
            ElseIf isOff Then
                ws.Cells(row, col).Value = "OFF"
                ws.Cells(row, col + 1).Value = "OFF"
            Else
                If shiftType <> "" Then
                    If shiftType = "matin" Then
                        startTime = baseStartMatin
                        endTime = baseEndMatin
                    ElseIf shiftType = "soir" Then
                        startTime = baseStartSoir
                        endTime = baseEndSoir
                    Else
                        GoTo NextDay
                    End If

                    ' Apply variant on second working day
                    Dim cnt As Integer
                    cnt = workingCountDict(personName)
                    If Not isOff And Not isAbsent Then
                        cnt = cnt + 1
                        workingCountDict(personName) = cnt
                        If cnt = 2 Then
                            If shiftType = "matin" Then
                                endTime = endTime - TimeSerial(1, 0, 0)
                            ElseIf shiftType = "soir" Then
                                startTime = startTime + TimeSerial(1, 0, 0)
                            End If
                        End If
                    End If

                    ws.Cells(row, col).Value = startTime
                    ws.Cells(row, col + 1).Value = endTime
                Else
                    ' No shift type – leave existing (do nothing)
                End If
            End If
NextDay:
        Next dayIndex
    Next row

    ' ---- IOBSP: if auto-rotate, we can clear or keep existing? We'll clear and leave empty.
    ' If manual, we read from input sheet.
    If Not useAutoRotate Then
        ' Write IOBSP from input sheet (columns L-O)
        Dim iobspHeaderRow As Long
        iobspHeaderRow = FindRow(ws, "IOBSP", 2)
        If iobspHeaderRow <> 0 Then
            Dim iobspDataStart As Long
            iobspDataStart = iobspHeaderRow + 2
            Dim iobspLastRow As Long
            iobspLastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
            If iobspLastRow >= iobspDataStart Then
                ws.Rows(iobspDataStart & ":" & iobspLastRow).ClearContents
            End If
            Dim iobspRow As Long
            iobspRow = 2
            Dim writeRow As Long
            writeRow = iobspDataStart
            Do While inputSheet.Cells(iobspRow, 12).Value <> ""   ' column L = 12
                Dim iobspName As String
                iobspName = Trim(inputSheet.Cells(iobspRow, 12).Value)
                If iobspName <> "" Then
                    ws.Cells(writeRow, 3).Value = iobspName
                    ws.Cells(writeRow, 4).Value = inputSheet.Cells(iobspRow, 13).Value
                    ws.Cells(writeRow, 5).Value = inputSheet.Cells(iobspRow, 14).Value
                    ws.Cells(writeRow, 6).Value = inputSheet.Cells(iobspRow, 15).Value
                    writeRow = writeRow + 1
                End If
                iobspRow = iobspRow + 1
            Loop
        End If
    Else
        ' Auto-rotate: clear IOBSP (or keep existing from copy? We'll leave it blank)
        ' We'll not clear it; user can manually edit.
    End If

    If useAutoRotate Then
        MsgBox "Auto-rotation applied (shift types flipped, OFF days rotated). Please review the generated schedule.", vbInformation
    End If
End Sub
