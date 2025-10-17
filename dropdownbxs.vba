Public Sub BuildNamedRanges()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long
    
    ' 1) Identify the sheet that has your newly imported data
    Set ws = ThisWorkbook.Sheets("Lists")
    
    ' 2) Find the last row in column A (Reason)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found in 'Lists' sheet!", vbExclamation
        Exit Sub
    End If
    
    ' 3) Create a dictionary to group sub reasons by reason
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare  ' case-insensitive keys
    
    ' 4) Loop from row 2 to lastRow, collecting reason/sub reason pairs
    For i = 2 To lastRow
        Dim r As String, sr As String
        r = Trim(ws.Cells(i, "A").Value)  ' Reason
        sr = Trim(ws.Cells(i, "B").Value) ' Sub Reason
        
        If r <> "" Then
            If Not dict.Exists(r) Then
                dict.Add r, sr
            Else
                ' Add more sub reasons separated by semicolons
                dict(r) = dict(r) & ";" & sr
            End If
        End If
    Next i
    
    ' 5) Now we have something like dict("VENDOR") = "BACK ORDERED;DELAY IN ARRIVAL;WEATHER"
    '    We'll place these sub reasons in columns D & E (or wherever) as a scratch area
    ws.Range("D:E").ClearContents
    
    Dim rowCounter As Long
    rowCounter = 2
    
    Dim key As Variant
    Dim arrSub() As String
    
    ' We'll also compile a unique list of reasons in col G
    ws.Range("G:G").ClearContents
    Dim reasonCounter As Long
    reasonCounter = 2
    
    ' 6) For each reason in the dictionary, split the sub reasons into lines
    For Each key In dict.Keys
        ws.Cells(rowCounter, "D").Value = key  ' store the reason name here for clarity
       
        arrSub = Split(dict(key), ";")
        
        Dim j As Long
        For j = LBound(arrSub) To UBound(arrSub)
            ws.Cells(rowCounter + j, "E").Value = Trim(arrSub(j))
        Next j
        
        ' We'll define a named range for key
        ' That named range points to E(rowCounter) through E(rowCounter + subCount - 1)
        Dim subCount As Long
        subCount = UBound(arrSub) - LBound(arrSub) + 1
        
        ' Delete any old named range that might match this reason
        On Error Resume Next
        ThisWorkbook.Names(key).Delete
        On Error GoTo 0
        
        ThisWorkbook.Names.Add _
            Name:=key, _
            RefersTo:="='" & ws.Name & "'!$E$" & rowCounter & ":$E$" & (rowCounter + subCount - 1)
        
        ' Move rowCounter so the next reason doesn't overwrite
        rowCounter = rowCounter + subCount + 1
        
        ' Meanwhile, also put the reason in col G for building "ReasonList"
        ws.Cells(reasonCounter, "G").Value = key
        reasonCounter = reasonCounter + 1
    Next key
    
    ' 7) Create or update the "ReasonList" named range to reference col G
    '    from G2 down to reasonCounter-1
    On Error Resume Next
    ThisWorkbook.Names("ReasonList").Delete
    On Error GoTo 0
    
    If reasonCounter > 2 Then
        ThisWorkbook.Names.Add _
            Name:="ReasonList", _
            RefersTo:="='" & ws.Name & "'!$G$2:$G$" & (reasonCounter - 1)
    End If
    
    MsgBox "Named ranges created or updated successfully!", vbInformation
End Sub




