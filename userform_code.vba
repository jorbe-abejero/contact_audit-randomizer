Private Sub ComboBox1_Change()
    ' Populate ComboBox2 (Processes) based on selected week
    Dim wsSource As Worksheet
    Dim processKey As Variant
    Dim weekKey As String
    Dim i As Long

    ' Set wsSource to the active sheet or specify the sheet name
    Set wsSource = ActiveSheet ' Or replace with: Set wsSource = ActiveWorkbook.Sheets("SheetName")

    weekKey = ComboBox1.value
    ComboBox2.Clear

    If weekKey <> "" Then
        ' Populate processes for the selected week
        For i = 8 To wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
            If Trim(wsSource.Cells(i, 1).value) = weekKey Then
                processKey = Trim(wsSource.Cells(i, 4).value)
                If processKey <> "" And Not ComboBoxContains(ComboBox2, processKey) Then ' Uses ComboBoxContains here
                    ComboBox2.AddItem processKey
                End If
            End If
        Next i
    End If
End Sub


Private Sub ComboBox2_Change()
    Dim wsSource As Worksheet
    Dim selectedWeek As String
    Dim selectedProcess As String
    Dim lastRow As Long
    Dim i As Long
    Dim availableCount As Long

    Set wsSource = ActiveSheet
    selectedWeek = ComboBox1.value
    selectedProcess = ComboBox2.value
    availableCount = 0

    If selectedWeek <> "" And selectedProcess <> "" Then
        lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

        For i = 8 To lastRow
            If wsSource.Cells(i, 1).value = selectedWeek And wsSource.Cells(i, 4).value = selectedProcess Then
                Dim contactName As String
                Dim companyName As String
                contactName = Trim(wsSource.Cells(i, 5).value)
                companyName = Trim(wsSource.Cells(i, 6).value)
                
                ' Count if at least one of the fields is not blank
                If Not (contactName = "" And companyName = "") Then
                    availableCount = availableCount + 1
                End If
            End If
        Next i
    End If

    lblAvailable.Caption = "Available contacts: " & availableCount
End Sub


Private Sub CommandButton1_Click()
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim weekKey As String
    Dim processKey As String
    Dim selectedContacts As Long
    Dim i As Long
    Dim addedContacts As Long
    Dim weekSeparatorRow As Long
    Dim contactExists As Boolean

    Set wsSource = ActiveSheet

    ' Check if "Randomized Results" sheet exists, create if not
    On Error Resume Next
    Set wsOutput = ActiveWorkbook.Sheets("Randomized Results")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Set wsOutput = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count))
        wsOutput.Name = "Randomized Results"
        wsOutput.Range("A1:F1").value = Array("Week", "Process", "Contact Name", "Company Name", "Tickets", "Comments")
    End If

    weekKey = ComboBox1.value
    processKey = ComboBox2.value
    selectedContacts = Val(TextBox1.value)

    If selectedContacts <= 0 Then
        MsgBox "Please enter a valid number of contacts to select."
        Exit Sub
    End If

    lastRow = wsOutput.Cells(wsOutput.Rows.count, 1).End(xlUp).Row
    rowIndex = lastRow + 1

    If weekKey <> "" And processKey <> "" Then
        If rowIndex > 2 Then
            If wsOutput.Cells(rowIndex - 1, 1).value <> weekKey Then
                wsOutput.Cells(rowIndex, 1).value = ""
                rowIndex = rowIndex + 1
            End If
        End If

        ' Gather all matching rows
        Dim matchIndexes() As Long
        Dim totalMatches As Long
        ReDim matchIndexes(1 To 1)
        totalMatches = 0

        For i = 8 To wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
            If wsSource.Cells(i, 1).value = weekKey And wsSource.Cells(i, 4).value = processKey Then
                If wsSource.Cells(i, 5).value <> "" And wsSource.Cells(i, 6).value <> "" Then
                    totalMatches = totalMatches + 1
                    ReDim Preserve matchIndexes(1 To totalMatches)
                    matchIndexes(totalMatches) = i
                End If
            End If
        Next i

        If totalMatches = 0 Then
            MsgBox "No available contacts found for the selected Week and Process."
            Exit Sub
        End If

        ' Random selection with duplicate check
        Dim usedIndexes As Object
        Set usedIndexes = CreateObject("Scripting.Dictionary")
        Dim srcRow As Long
        Dim rndIndex As Long
        addedContacts = 0

        Do While addedContacts < selectedContacts And usedIndexes.count < totalMatches
            rndIndex = Int((totalMatches) * Rnd) + 1

            If Not usedIndexes.exists(rndIndex) Then
                usedIndexes.Add rndIndex, True
                srcRow = matchIndexes(rndIndex)

                Dim contactName As String
                Dim companyName As String
                contactName = wsSource.Cells(srcRow, 5).value
                companyName = wsSource.Cells(srcRow, 6).value

                contactExists = False
                Dim j As Long
                For j = 2 To wsOutput.Cells(wsOutput.Rows.count, 1).End(xlUp).Row
                    If wsOutput.Cells(j, 3).value = contactName And wsOutput.Cells(j, 4).value = companyName Then
                        contactExists = True
                        Exit For
                    End If
                Next j

                If Not contactExists Then
                    wsOutput.Cells(rowIndex, 1).value = weekKey
                    wsOutput.Cells(rowIndex, 2).value = processKey
                    wsOutput.Cells(rowIndex, 3).value = contactName
                    wsOutput.Cells(rowIndex, 4).value = companyName
                    wsOutput.Cells(rowIndex, 5).value = wsSource.Cells(srcRow, 7).value
                    wsOutput.Cells(rowIndex, 6).value = wsSource.Cells(srcRow, 8).value

                    rowIndex = rowIndex + 1
                    addedContacts = addedContacts + 1
                End If
            End If
        Loop

        If addedContacts < selectedContacts Then
            MsgBox "Only " & addedContacts & " unique contacts were available for the selected process."
        End If
    End If

    wsOutput.Columns.AutoFit
End Sub


Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    ' Your existing code here

    ' Initialize ComboBox1 (Weeks)
    Dim wsSource As Worksheet
    Dim lastRow As Long
    Dim weekKey As Variant

    Set wsSource = ActiveSheet
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row

    ' Populate ComboBox1 (Weeks)
    For i = 8 To lastRow
        weekKey = Trim(wsSource.Cells(i, 1).value)
        If Left(weekKey, 4) = "Week" And weekKey <> "" Then ' Ensure it starts with "Week" and is not blank
            If Not ComboBoxContains(ComboBox1, weekKey) Then ' Add weeks if not already in the list
                ComboBox1.AddItem weekKey
            End If
        End If
    Next i
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub


Function ComboBoxContains(cmb As Object, value As Variant) As Boolean
    Dim i As Long
    ComboBoxContains = False
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = value Then
            ComboBoxContains = True
            Exit Function
        End If
    Next i
End Function

