Sub OperationConsign(ByVal time1 As Single)
    On Error GoTo ErrorHandler
    Dim Sheet As Worksheet: Set Sheet = Worksheets("1540")
    Dim SummarySht As Worksheet: Set SummarySht = Worksheets(CStr(Year(Date)))
    Dim Table As ListObject: Set Table = Sheet.ListObjects(1)
    Dim NewSheet As Worksheet: Set NewSheet = Worksheets.Add
    Dim LastColumn As Integer: LastColumn = Sheet.Range("A1").End(xlToRight).Column
    Dim LC As Integer: LC = SummarySht.Cells(13, SummarySht.Columns.Count).End(xlToLeft).Column
    Dim LastRow, i As Integer: LastRow = Sheet.Range("A1").End(xlDown).Row
    Dim NewShtLastRow As Integer
    Dim CachedRange As Range: Set CachedRange = Sheet.Range("A1").Resize(1, LastColumn)
    Dim DataRange, temp As Range: Set DataRange = Sheet.Range("A2").Resize(LastRow - 1, LastColumn)
    Dim VendorCount As Integer: VendorCount = Worksheets("Mapping Table").Range("G1").End(xlDown).Row
    Vendors = Worksheets("Mapping Table").Range("G2").Resize(VendorCount - 1, 1)
    Days = Worksheets("Mapping Table").Range("H2").Resize(VendorCount - 1, 1)
    
    
    For i = 1 To UBound(Vendors)
        
        Table.Range.AutoFilter Table.ListColumns("Category").Index, "Supplies", xlFilterValues
        Table.Range.AutoFilter Table.ListColumns("Name 1").Index, Vendors(i, 1), xlFilterValues
        Table.Range.AutoFilter Table.ListColumns("Days Since Receipt").Index, Days(i, 1), xlFilterValues
                      
        Set temp = DataRange.SpecialCells(xlCellTypeVisible)
                
        If Not temp Is Nothing Then
            
           Set CachedRange = Union(CachedRange, temp)
        
        End If
               
        
        Table.AutoFilter.ShowAllData
    Next
    
    CachedRange.Copy NewSheet.Range("A1")
    NewSheet.Name = "Supplies"
    NewSheet.Range("A:AF").EntireColumn.AutoFit
    
    NewShtLastRow = NewSheet.Cells(NewSheet.Rows.Count, 1).End(xlUp).Row
    
    If NewShtLastRow > 1 Then
        SummarySht.Cells(13, LC + 1).Value = WorksheetFunction.Sum(NewSheet.Range("Y2").Resize(NewShtLastRow - 1, 1))
    Else
        SummarySht.Cells(13, LC + 1).Value = 0
    End If
    
    SummarySht.Activate
    
    MsgBox Format((Timer - time1) / 86400, "hh:mm:ss")
    
    Exit Sub
ErrorHandler:
    If Err.Number = 1004 Then
               
        Set temp = Nothing
        Resume Next
    Else
        MsgBox ("資料格式有誤--Phase3--Reason:" & Err.Description)
    
    End If
End Sub

Sub Consign(ByVal time1 As Single)
    On Error GoTo ErrorHandler
    Dim Sheet As Worksheet: Set Sheet = Worksheets("1540")
    Dim SummarySht As Worksheet: Set SummarySht = Worksheets(CStr(Year(Date)))
    Dim LC As Integer: LC = SummarySht.Cells(4, SummarySht.Columns.Count).End(xlToLeft).Column
    Dim Table As ListObject: Set Table = Sheet.ListObjects.Add(Source:=Sheet.Range("A1").CurrentRegion)
    Dim NewSheet As Worksheet
    Dim LastColumn As Integer: LastColumn = Sheet.Range("A1").End(xlToRight).Column
    Dim LastRow, i As Integer: LastRow = Sheet.Range("A1").End(xlDown).Row
    Dim CachedRange As Range: Set CachedRange = Sheet.Range("A1").Resize(1, LastColumn)
    Dim DataRange, temp As Range: Set DataRange = Sheet.Range("A2").Resize(LastRow - 1, LastColumn)
    Dim NewShtLastRow As Integer
    arr = [{"Leadframe","Substrate","Packing Materials","Die attach","Mold Compound","Gold wire","Copper wire","Solder","Contactor"}]
    Days = Worksheets("Mapping Table").Range("E2:E10")
    
    For i = 1 To UBound(arr)
        
        Table.Range.AutoFilter Table.ListColumns("Category").Index, arr(i), xlFilterValues
        Table.Range.AutoFilter Table.ListColumns("Days Since Receipt").Index, Days(i, 1), xlFilterValues
        
        Set temp = DataRange.SpecialCells(xlCellTypeVisible)
                
        If Not temp Is Nothing Then
            
           Set CachedRange = Union(CachedRange, temp)
        
        End If
        
        Set NewSheet = Worksheets.Add
        
        NewSheet.Name = arr(i)
                
        CachedRange.Copy NewSheet.Range("A1")
        
        NewSheet.Range("A:AF").EntireColumn.AutoFit
        
        NewShtLastRow = NewSheet.Cells(NewSheet.Rows.Count, 1).End(xlUp).Row
        
        If NewShtLastRow > 1 Then
        
            SummarySht.Cells(i + 3, LC + 1).Value = WorksheetFunction.Sum(NewSheet.Range("Y2").Resize(NewShtLastRow - 1, 1))
        
        Else
            
            SummarySht.Cells(i + 3, LC + 1).Value = 0
        End If
        
        
        Table.AutoFilter.ShowAllData
        
        Set CachedRange = Sheet.Range("A1").Resize(1, LastColumn)
    
    Next
        
    Call OperationConsign(time1)
    Exit Sub
ErrorHandler:
    
    If Err.Number = 1004 Then
               
        Set temp = Nothing
        Resume Next
    Else
        MsgBox ("資料格式有誤--Phase2--Reason:" & Err.Description)
    
    End If
End Sub

Sub InsertMappingData()
    On Error GoTo ErrorHandler
    Dim Sheet As Worksheet: Set Sheet = Worksheets("1540")
    Dim LastRow As Integer: LastRow = Sheet.Range("A1").End(xlDown).Row
    Dim time1 As Single
    
    time1 = Timer
    
    Sheet.Range("B:B").EntireColumn.Insert
    Sheet.Range("B1").Value = "Category"
    Sheet.Range("B2").Resize(LastRow - 1, 1).Formula = "=IFERROR(INDEX('Mapping table'!C2,MATCH('1540'!RC[25],'Mapping table'!C1,0)),""NA"")"
    
    Call Consign(time1)
    Exit Sub
ErrorHandler:
    MsgBox ("資料格式有誤--Phase1-Mapping Data--Reason:" & Err.Description)
End Sub

Sub DeleteData()
    On Error GoTo ErrorHandler
    arr = [{"1540","Leadframe","Substrate","Packing Materials","Die attach","Mold Compound","Gold wire","Copper wire","Solder","Contactor","Supplies"}]
    Dim Sheet As Worksheet
    
    Application.DisplayAlerts = False
    
    Worksheets(arr).Delete
    
    Application.DisplayAlerts = True
    
    MsgBox ("Done")
    
    Exit Sub
ErrorHandler:
    
    MsgBox ("No available Worksheet found.")
End Sub
