Option Explicit

' === Helper: counts Mon–Fri between two dates (inclusive).
'     Optionally pass a holidays range like wsBuilder.Range("AA2:AA50").
Private Function BusinessDays(ByVal d1 As Date, ByVal d2 As Date, Optional holidays As Range = Nothing) As Long
    If d1 > d2 Then
        BusinessDays = 0
        Exit Function
    End If
    If holidays Is Nothing Then
        BusinessDays = Application.WorksheetFunction.NetworkDays(d1, d2)
    Else
        BusinessDays = Application.WorksheetFunction.NetworkDays(d1, d2, holidays)
    End If
End Function

Public Sub CountRowsByMonthAndOTIFAndOccupiedAndAverageAgeAndVacantAndColor()
    Dim wsBuilder As Worksheet
    Dim wsFormatted As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Dim monthCounts As Object
    Dim monthOTIFCounts As Object
    Dim monthOccupiedCounts As Object
    Dim monthOccupiedAges As Object
    Dim monthVacantCounts As Object
    Dim monthVacantAges As Object
    
    Dim monthName As String
    Dim dateValue As Date
    Dim averageAge As Double
    Dim ageValue As Variant
    Dim statusText As String
    
    ' --- Config: change if your columns differ ---
    Const COL_START As String = "E"     ' Start date
    Const COL_COMPLETED As String = "F" ' Completed date
    Const WRITE_BACK_AGE As Boolean = True    ' write business-day age into column G
    Const EXCLUSIVE_START As Boolean = False  ' True to subtract 1 day (exclude start)
    'Dim holidayRange As Range ' (optional) set to your holidays range and pass it below
    
    ' Initialize dictionaries
    Set monthCounts = CreateObject("Scripting.Dictionary")
    Set monthOTIFCounts = CreateObject("Scripting.Dictionary")
    Set monthOccupiedCounts = CreateObject("Scripting.Dictionary")
    Set monthOccupiedAges = CreateObject("Scripting.Dictionary")
    Set monthVacantCounts = CreateObject("Scripting.Dictionary")
    Set monthVacantAges = CreateObject("Scripting.Dictionary")
    
    ' Set references to the worksheets
    On Error Resume Next
    Set wsBuilder = ThisWorkbook.Sheets("Builder Data")
    On Error GoTo 0
    
    If wsBuilder Is Nothing Then
        MsgBox "The sheet 'Builder Data' was not found.", vbCritical
        Exit Sub
    End If
    
    ' Check if "Formatted data" sheet exists, if not, create it
    On Error Resume Next
    Set wsFormatted = ThisWorkbook.Sheets("Formatted data")
    On Error GoTo 0
    If wsFormatted Is Nothing Then
        Set wsFormatted = ThisWorkbook.Sheets.Add
        wsFormatted.Name = "Formatted data"
    End If
    
    ' Clear existing data in "Formatted data" sheet
    wsFormatted.Cells.Clear
    
    ' Get the last row with data in column E of the "Builder Data" sheet
    lastRow = wsBuilder.Cells(wsBuilder.Rows.Count, COL_START).End(xlUp).Row
    
    ' Loop through each row to count and color
    For i = 2 To lastRow
        If IsDate(wsBuilder.Cells(i, COL_START).Value) Then
            dateValue = wsBuilder.Cells(i, COL_START).Value
            monthName = Format(dateValue, "mmmm")
            
            Dim startDate As Variant, endDate As Variant
            Dim ageBiz As Variant
            startDate = wsBuilder.Cells(i, COL_START).Value
            endDate = wsBuilder.Cells(i, COL_COMPLETED).Value
            
            If IsDate(startDate) And IsDate(endDate) Then
                ' If you keep holidays, pass them: BusinessDays(CDate(startDate), CDate(endDate), holidayRange)
                ageBiz = BusinessDays(CDate(startDate), CDate(endDate))
                If EXCLUSIVE_START Then ageBiz = ageBiz - 1
                If ageBiz < 0 Then ageBiz = 0
                ageValue = ageBiz
                If WRITE_BACK_AGE Then wsBuilder.Cells(i, "G").Value = ageValue
            Else
                ' If one of the dates is missing or invalid, fall back to existing G or blank
                ageValue = wsBuilder.Cells(i, "G").Value
            End If

            statusText = UCase$(Trim$(CStr(wsBuilder.Cells(i, "I").Value)))
            
            ' ---- Color by Status (I) + Age (G) thresholds ----
            If IsNumeric(ageValue) And Len(statusText) > 0 Then
                Select Case statusText
                    Case "VACANT"
                        Select Case CLng(ageValue)
                            Case 0 To 3
                                wsBuilder.Rows(i).Interior.Color = RGB(144, 238, 144) ' Light green
                            Case 4 To 5
                                wsBuilder.Rows(i).Interior.Color = RGB(255, 255, 0)   ' Yellow
                            Case Is >= 6
                                wsBuilder.Rows(i).Interior.Color = RGB(255, 0, 0)     ' Red
                            Case Else
                                wsBuilder.Rows(i).Interior.Pattern = xlNone
                        End Select
                    Case "OCCUPIED"
                        Select Case CLng(ageValue)
                            Case 0 To 5
                                wsBuilder.Rows(i).Interior.Color = RGB(144, 238, 144) ' Light green
                            Case 6 To 8
                                wsBuilder.Rows(i).Interior.Color = RGB(255, 255, 0)   ' Yellow
                            Case Is >= 9
                                wsBuilder.Rows(i).Interior.Color = RGB(255, 0, 0)     ' Red
                            Case Else
                                wsBuilder.Rows(i).Interior.Pattern = xlNone
                        End Select
                    Case Else
                        ' leave as-is
                End Select
            End If
            ' ---- END color block ----
            
            ' Count months
            If monthCounts.Exists(monthName) Then
                monthCounts(monthName) = monthCounts(monthName) + 1
            Else
                monthCounts.Add monthName, 1
            End If
            
            ' Count OTIF "Yes"
            If wsBuilder.Cells(i, "J").Value = "Yes" Then
                If monthOTIFCounts.Exists(monthName) Then
                    monthOTIFCounts(monthName) = monthOTIFCounts(monthName) + 1
                Else
                    monthOTIFCounts.Add monthName, 1
                End If
            End If
            
            ' Count Occupied and sum ages
            If statusText = "OCCUPIED" And IsNumeric(ageValue) Then
                If monthOccupiedCounts.Exists(monthName) Then
                    monthOccupiedCounts(monthName) = monthOccupiedCounts(monthName) + 1
                    monthOccupiedAges(monthName) = monthOccupiedAges(monthName) + CDbl(ageValue)
                Else
                    monthOccupiedCounts.Add monthName, 1
                    monthOccupiedAges.Add monthName, CDbl(ageValue)
                End If
            End If
            
            ' Count Vacant and sum ages
            If statusText = "VACANT" And IsNumeric(ageValue) Then
                If monthVacantCounts.Exists(monthName) Then
                    monthVacantCounts(monthName) = monthVacantCounts(monthName) + 1
                    monthVacantAges(monthName) = monthVacantAges(monthName) + CDbl(ageValue)
                Else
                    monthVacantCounts.Add monthName, 1
                    monthVacantAges.Add monthName, CDbl(ageValue)
                End If
            End If
        End If
    Next i
    
    ' Output the results to the "Formatted data" sheet
    Dim monthOrder As Variant
    monthOrder = Array("January", "February", "March", "April", "May", "June", _
                       "July", "August", "September", "October", "November", "December")
    
    ' Set up headers
    wsFormatted.Cells(1, 1).Value = "Month"
    wsFormatted.Cells(1, 2).Value = "Count"
    wsFormatted.Cells(1, 3).Value = "OTIF Count"
    wsFormatted.Cells(1, 4).Value = "Occupied Count"
    wsFormatted.Cells(1, 5).Value = "Average Age (Occupied)"
    wsFormatted.Cells(1, 6).Value = "Vacant Count"
    wsFormatted.Cells(1, 7).Value = "Average Age (Vacant)"
    
    Dim rowIndex As Long
    rowIndex = 2
    
    Dim m As Variant
    For Each m In monthOrder
        wsFormatted.Cells(rowIndex, 1).Value = m
        
        If monthCounts.Exists(m) Then
            wsFormatted.Cells(rowIndex, 2).Value = monthCounts(m)
        Else
            wsFormatted.Cells(rowIndex, 2).Value = 0
        End If
        
        If monthOTIFCounts.Exists(m) Then
            wsFormatted.Cells(rowIndex, 3).Value = monthOTIFCounts(m)
        Else
            wsFormatted.Cells(rowIndex, 3).Value = 0
        End If
        
        If monthOccupiedCounts.Exists(m) Then
            wsFormatted.Cells(rowIndex, 4).Value = monthOccupiedCounts(m)
            If monthOccupiedCounts(m) > 0 Then
                averageAge = monthOccupiedAges(m) / monthOccupiedCounts(m)
            Else
                averageAge = 0
            End If
            wsFormatted.Cells(rowIndex, 5).Value = averageAge
        Else
            wsFormatted.Cells(rowIndex, 4).Value = 0
            wsFormatted.Cells(rowIndex, 5).Value = 0
        End If
        
        If monthVacantCounts.Exists(m) Then
            wsFormatted.Cells(rowIndex, 6).Value = monthVacantCounts(m)
            If monthVacantCounts(m) > 0 Then
                averageAge = monthVacantAges(m) / monthVacantCounts(m)
            Else
                averageAge = 0
            End If
            wsFormatted.Cells(rowIndex, 7).Value = averageAge
        Else
            wsFormatted.Cells(rowIndex, 6).Value = 0
            wsFormatted.Cells(rowIndex, 7).Value = 0
        End If
        
        rowIndex = rowIndex + 1
    Next m
End Sub
