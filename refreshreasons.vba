Public Sub RefreshReasons()
    Dim wbMaster As Workbook
    Dim wsMaster As Worksheet
    Dim wsLocal As Worksheet
    Dim masterPath As String

    masterPath = "S:\SA - Warranty\BUILDER SCORE CARDS\MasterList.xlsx" ' <-- Path to your master list
    Set wsLocal = ThisWorkbook.Sheets("Lists")
    
    On Error Resume Next
    Set wbMaster = Workbooks.Open(masterPath, UpdateLinks:=0, ReadOnly:=True)
    On Error GoTo 0
    
    If wbMaster Is Nothing Then
        MsgBox "Could not open: " & masterPath, vbCritical
        Exit Sub
    End If
    
    Set wsMaster = wbMaster.Sheets("MasterData") ' rename if needed
    
    ' Clear old data in local Lists sheet
    wsLocal.Cells.Clear
    
    ' Copy master data (A:B for reasons & subreasons)
    wsMaster.Range("A1:B1000").Copy wsLocal.Range("A1")
    
    wbMaster.Close SaveChanges:=False
    
    MsgBox "Master list imported!", vbInformation
End Sub
