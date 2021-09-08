Attribute VB_Name = "Module1"
Sub raw_to_processed()
'
'  Macro for the processing of raw data- creates a new file, sorts & spaces samples correctly, and performs statistical
'  tests for each material. For major and trace elements.
'



  ' check that data has not already been processed
    Dim response As String
    If (ActiveSheet.Name = "Processed Data") Or (Range("D10") <> 0) Then
        response = MsgBox("Check that active sheet has not already been processed from raw. If you have selected the wrong sheet, please hit cancel.", vbOKCancel)
        If response = vbCancel Then
            Exit Sub
        End If
    End If
    
    

  ' creates new sheet, copies data, renames to "Processed Data", freezes first column
    Cells.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    ActiveSheet.Name = "Processed Data"
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 0
    End With
    ActiveWindow.FreezePanes = True
    
    
    
    
  ' finds & deletes 'summary' rows, number,avg, etc before sorting
    Dim Cell As Range
    Columns("A:A").Select
    
    Set Cell = Selection.Find(What:="Number", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If Not Cell Is Nothing Then
        Cell.EntireRow.Delete
    End If
    
    Set Cell = Selection.Find(What:="Average", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If Not Cell Is Nothing Then
        Cell.EntireRow.Delete
    End If
    
    Set Cell = Selection.Find(What:="Maximum", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If Not Cell Is Nothing Then
        Cell.EntireRow.Delete
    End If
    
    Set Cell = Selection.Find(What:="Minimum", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If Not Cell Is Nothing Then
        Cell.EntireRow.Delete
    End If
    
    Set Cell = Selection.Find(What:="Range", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If Not Cell Is Nothing Then
        Cell.EntireRow.Delete
    End If
    
    Set Cell = Selection.Find(What:="Std dev.", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If Not Cell Is Nothing Then
        Cell.EntireRow.Delete
    End If
    
    Set Cell = Selection.Find(What:="RSD(%)", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If Not Cell Is Nothing Then
        Cell.EntireRow.Delete
    End If


     
  ' deletes empty columns
    Dim rng As Range, result As Long
    
    Set rng = ActiveSheet.UsedRange
    For result = rng.Columns.count To 1 Step -1
        If Application.CountA(Columns(result).EntireColumn) = 0 Then
            Columns(result).Delete
        End If
    Next result

  ' deletes columns D & E if filled with zero's
    Dim count As Integer
    For count = 1 To 2
        Range("D50").Select
        ActiveCell.FormulaR1C1 = "=SUM(R[-49]C:R[-1]C)"
        If ActiveCell = 0 Then
            ActiveCell.EntireColumn.Delete
        End If
    Next count

  ' sorts list by sample A-Z
    Range("A3:AZ100").Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add2 Key:=Range( _
        "A3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A3:AZ100")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
  ' find last row & column, stores as variables
    
    Dim LastRow As Long, LastCol As Long
    
    ActiveSheet.UsedRange
    LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.count).Row
    LastRow = LastRow + (3 * LastRow)

    ActiveSheet.UsedRange
    LastCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.count).Column
    
    
    
    
  ' adds 3 spaces between each diff. sample & adds stats for avg, std dev, & RSD
    Dim a As String, b As String, cont As Integer, e As Integer, rg As Integer, fin As Double
    Dim unaccept As FormatCondition, rmcont As Integer
    
    
    rmcont = 0
    i = 3
    cont = 0
       
    For i = 3 To LastRow
        a = Range("A" & i)
        b = Range("A" & i + 1)
        
        If Not (StrComp(a, b, vbTextCompare)) Then
            cont = cont + 1
        End If
    
        If (StrComp(a, b, vbTextCompare)) Then
            If Not Range("A" & i + 1) = 0 Then
                cont = cont + 1
            End If
            rg = i + 1 - cont
            fin = 0
            Range("A" & i + 1).Select
            Selection.EntireRow.Insert
            With Selection.EntireRow
                .ClearFormats
            End With
            
            ActiveCell.FormulaR1C1 = "Average"
            For e = 4 To LastCol
                Do While rg < i + 1
                    If (ActiveSheet.Cells(rg, e).Interior.Color) <> RGB(255, 255, 255) Then
                        cont = cont - 1
                        rmcont = rmcont + 1
                    Else
                        fin = fin + (Cells(rg, e))
                    End If
                    rg = rg + 1
                Loop
                
                Cells(i + 1, e) = fin / cont
                fin = 0
                cont = cont + rmcont
                rg = i + 1 - cont
                rmcont = 0
            Next e
            
            Range("A" & i + 2).Select
            Selection.EntireRow.Insert shift:=xlDown
            ActiveCell.FormulaR1C1 = "Std dev"
            For e = 4 To LastCol
                Do While rg < i + 2
                    If (ActiveSheet.Cells(rg, e).Interior.Color) = RGB(255, 255, 255) Then
                        fin = fin + ((Cells(rg, e)) - (Cells(i + 1, e))) * ((Cells(rg, e)) - (Cells(i + 1, e)))
                    End If
                    rg = rg + 1
                Loop
                If (cont > 1) Then
                    Cells(i + 2, e) = ((fin / (cont - 1))) ^ (1 / 2)
                Else: Cells(i + 2, e) = "0"
                End If
                fin = 0
                rg = i + 1 - cont
            Next e
            
            Range("A" & i + 3).Select
            Selection.EntireRow.Insert shift:=xlDown
            ActiveCell.FormulaR1C1 = "RSD(%)"
            For e = 4 To LastCol
                If Not Cells(i + 1, e) = 0 Then
                    Cells(i + 3, e) = 100 * Cells(i + 2, e) / Cells(i + 1, e)
                    Cells(i + 3, e).Select
                    Selection.FormatConditions.Delete
                    Set unaccept = Selection.FormatConditions.Add(xlCellValue, xlGreater, "=4.0")
                    With unaccept
                        .Font.ColorIndex = 53
                        .Interior.Color = RGB(255, 182, 193)
                    End With
                Else
                    Cells(i + 3, e) = 0
                End If
            Next e
            
            i = i + 3
            cont = 0
        End If
    Next i
    
    
  ' autofits header
  
    Cells.EntireColumn.AutoFit
    

End Sub




Sub major_elements_data_table()
'
'
'

  ' variable dec.
    
    Dim eleStr As String, i As Integer, countCol As Integer, countRow As Integer, a As String, b As String, ii As Integer, temp As Double
  
  ' find last row & column of processed sheet, stores as variables
       
    Dim LastRow As Long, LastCol As Long
    
    ActiveSheet.UsedRange
    LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.count).Row
    LastRow = LastRow

    ActiveSheet.UsedRange
    LastCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.count).Column
    
    

' failsafes

  ' check that data has already been processed by first macro
    Dim response As String
    If Range("D3") = 0 And Range("D10") = 0 Then
        response = MsgBox("Check that active sheet has already been processed from raw. If you selected the wrong sheet, please hit cancel.", vbOKCancel)
        If response = vbCancel Then
            Exit Sub
        End If
    End If
  ' check for ppm
    For i = 2 To LastCol
        If (Cells(2, i) = "ppm") Then
            response = MsgBox("Check that active sheet is processed major data. If you selected the wrong sheet, please hit cancel.", vbOKCancel)
            If response = vbCancel Then
                Exit Sub
            End If
        End If
    Next
   
   ' setting up array & adding first 2 values
   
     Dim arr()
     
     ReDim arr(1, 1)
     
     arr(0, 0) = "Sample ID"
     arr(1, 0) = "Major Elements (wt. %)"

       
  ' copying oxide elements & adding to array
    
    countCol = 2
     
    For i = 4 To LastRow
       
        eleStr = Cells(1, i)
        
        If (eleStr Like "*O*") Then
            arr = ReDimPreserve(arr, countCol, countCol)
            arr(countCol, 0) = Cells(1, i)
            
            countCol = countCol + 1
         End If
        
    Next i
    
    
  ' next iteration of rows; gathers sample ID and averages for each oxide & stores each array
  
    countRow = 1
        
    For i = 3 To LastRow
       
        a = Range("A" & i)
        b = Range("A" & i + 1)
    
        If (StrComp(a, b, vbTextCompare)) Then
    
            arr = ReDimPreserve(arr, countCol, countRow)
            arr(0, countRow) = Range("A" & i)
            For ii = 2 To countCol
            
                temp = Cells(i + 1, ii + 2)
                arr(ii, countRow) = (temp)
            Next ii
            
            countRow = countRow + 1
            i = i + 3

        End If
    
    Next i
    
     
  ' new sheet!
  
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Data Table"
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 0
    End With
      
     
  ' adding array to new sheet
    
    Dim numRow As Long, numCol As Long, e As Integer
    
    e = 1
    numRows = UBound(arr)
    numCols = UBound(arr, 2)
    
    
    With ActiveSheet.Range("A1:AZ1")
        .Range(.Cells(1, 1), .Cells(1, (numCols + 1))).Resize(numRows).Value = arr
    End With
    
    
    
  ' now we start with calculating original sum, multiplying Fe2O3, & renormalizing so new sum = 100
  
    Dim ogsum As Double, newsum As Double, finsum As Double
     
  
    Cells(numRows + 1, 1) = "Orignal Sum"

    For i = 2 To numCols + 1
        ogsum = 0#
        For e = 2 To numRows
            ogsum = ogsum + Cells(e, i)
        Next e
        Cells(numRows + 1, i) = Round(ogsum, 2)
    Next i
   
  
  ' multiplyng Fe2O3 to get FeO
    
    For i = 2 To numRows
        If (Cells(i, 1) = "Fe2O3") Then
            Cells(i, 1) = "FeO*"
            For e = 2 To numCols + 1
                Cells(i, e) = Cells(i, e).Value * 0.8998
            Next e
        End If
    Next i
    
  ' calculating newsum
    Cells(numRows + 2, 1) = "Sum after FeO Conversion"
    For i = 2 To numCols + 1
        newsum = 0#
        For e = 3 To numRows
            newsum = newsum + Cells(e, i).Value
        Next e
        Cells(numRows + 2, i) = Round(newsum, 2)
    Next i
    
    
 ' multiplying everything 100/newsum
  
    For i = 2 To numCols + 1
        For e = 3 To numRows
            Cells(e, i) = Round(Cells(e, i).Value * 100 / (Cells(numRows + 2, i)), 2)
        Next e
    Next i
  
  ' calulating finsum
  
    Cells(numRows + 3, 1) = "Final Sum"
    For i = 2 To numCols + 1
        finsum = 0#
        For e = 3 To numRows
            finsum = finsum + Cells(e, i).Value
        Next e
        Cells(numRows + 3, i) = finsum
    Next i
  
  
  ' re-sorting rows: 'SiO2, TiO2, Al2O3, FeO*, MnO, MgO, CaO, Na2O,
  ' K2O, P2O5, then anything else
    
    Dim elementsList(1 To 10) As String
    elementsList(1) = "SiO2"
    elementsList(2) = "TiO2"
    elementsList(3) = "Al2O3"
    elementsList(4) = "FeO*"
    elementsList(5) = "MnO"
    elementsList(6) = "MgO"
    elementsList(7) = "CaO"
    elementsList(8) = "Na2O"
    elementsList(9) = "K2O"
    elementsList(10) = "P2O5"
    
    Application.AddCustomList ListArray:=elementsList
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Range("A3:AZ" & numRows).Sort , Key1:=Range("A:A"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=Application.CustomListCount + 1, MatchCase:=False, _
    Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    ActiveSheet.Sort.SortFields.Clear
    Application.DeleteCustomList Application.CustomListCount
    
  ' autofits header
  
    Cells.EntireColumn.AutoFit


'
End Sub


Public Function ReDimPreserve(aArrayToPreserve, nNewFirstUBound, nNewLastUBound)
    ReDimPreserve = False
    'check if its in array first
    If IsArray(aArrayToPreserve) Then
        'create new array
        ReDim aPreservedArray(nNewFirstUBound, nNewLastUBound)
        'get old lBound/uBound
        nOldFirstUBound = UBound(aArrayToPreserve, 1)
        nOldLastUBound = UBound(aArrayToPreserve, 2)
        'loop through first
        For nFirst = LBound(aArrayToPreserve, 1) To nNewFirstUBound
            For nLast = LBound(aArrayToPreserve, 2) To nNewLastUBound
                'if its in range, then append to new array the same way
                If nOldFirstUBound >= nFirst And nOldLastUBound >= nLast Then
                    aPreservedArray(nFirst, nLast) = aArrayToPreserve(nFirst, nLast)
                End If
            Next
        Next
        'return the array redimmed
        If IsArray(aPreservedArray) Then ReDimPreserve = aPreservedArray
    End If
End Function

Sub trace_processed_to_table()
'
'
'

' find last row & column of processed sheet, stores as variables
    
     
    Dim LastRow As Long, LastCol As Long
    
    ActiveSheet.UsedRange
    LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.count).Row
    LastRow = LastRow

    ActiveSheet.UsedRange
    LastCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.count).Column
    
    
    
   ' other variable dec.
    
     Dim eleStr As String, i As Integer, countCol As Integer, countRow As Integer, a As String, b As String, ii As Integer
     Dim countPPMCol As Integer, countPPMRow As Integer
     
    

' Failsafes
  ' check that data has already been processed by first macro
    Dim response As String
    If (ActiveSheet.Name = "Raw Data") Or (Range("D10") = 0) Then
        response = MsgBox("Check that active sheet has already been processed from raw. If you selected the wrong sheet, please hit cancel.", vbOKCancel)
        If response = vbCancel Then
            Exit Sub
        End If
    End If
    
  ' check for ppm - needs work
    For i = 2 To LastCol
        If (Cells(2, i) = "ppm") Then
            'response = MsgBox("Check that active sheet is processed trace data. If you selected the wrong sheet, please hit cancel.", vbOKCancel)
            'If response = vbCancel Then
                'Exit Sub
            'End If
        End If
    Next
   

  
   ' setting up array & adding first 2 values
   
     Dim arr()
     
     ReDim arr(1, 1)
     
     arr(0, 0) = "Sample ID"
     'arr(1, 0) = "Major Elements (wt. %)"

       
  ' copying elements measured in ppm & adding to array
    
    countRow = 2
     
    For i = 4 To LastRow
       
        eleStr = Cells(2, i)
        
        If (eleStr = "ppm") Then
            arr = ReDimPreserve(arr, countRow, countRow)
            arr(countRow, 0) = Cells(1, i)
            arr(countRow, 1) = "ppm"
            
            countRow = countRow + 1
         End If
        
    Next i
    
    
  ' next iteration of rows; gathers sample ID and averages for each element in ppm & stores in array
  
    countCol = 2
    countPPMCol = 2
        
    arr = ReDimPreserve(arr, 100, 100)
    For i = 3 To LastRow
       
        a = Range("A" & i)
        b = Range("A" & i + 1)
    
        If (StrComp(a, b, vbTextCompare)) Then
            countPPMCol = 2
            
            arr(0, countCol) = Range("A" & i)
            
            For ii = 2 To LastCol
            
                If (Cells(2, ii) = "ppm") Then
                    If (Cells(1, ii) = "Rb") Or (Cells(1, ii) = "Y") Then
                        arr(countPPMCol, countCol) = Round(Cells(i + 1, ii), 1)
                    Else
                        arr(countPPMCol, countCol) = Format(Cells(i + 1, ii), "0")
                    End If
                    countPPMCol = countPPMCol + 1
                End If
            Next ii
            
            countCol = countCol + 1
            i = i + 3

        End If
    
    Next i
    
     
  ' new sheet!
  
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Data Table"
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 0
    End With
      
     
  ' adding array to new sheet
    
    Dim numRows As Long, numCols As Long, e As Integer
    
    e = 1
    numRows = UBound(arr)
    numCols = UBound(arr, 2)
    
    
    With ActiveSheet.Range("A1:AZ1")
        .Range(.Cells(1, 1), .Cells(1, (numRows + 1))).Resize(numCols).Value = arr
    End With
    
    
  ' re-sorting rows: ' V, Cr, Ni, Cu, Zn, Rb, Sr, Y, Zr, Nb, Ba, Pb, then anything else'
    
    Dim traceList(1 To 12) As String
    traceList(1) = "V"
    traceList(2) = "Cr"
    traceList(3) = "Ni"
    traceList(4) = "Cu"
    traceList(5) = "Zn"
    traceList(6) = "Rb"
    traceList(7) = "Sr"
    traceList(8) = "Y"
    traceList(9) = "Zr"
    traceList(10) = "Nb"
    traceList(11) = "Ba"
    traceList(12) = "Pb"
    
    Application.AddCustomList ListArray:=traceList
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Range("A3:AZ" & numRows).Sort , Key1:=Range("A:A"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=Application.CustomListCount + 1, MatchCase:=False, _
    Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    ActiveSheet.Sort.SortFields.Clear
    Application.DeleteCustomList Application.CustomListCount
    
    
    
  ' autofits header
  
    Cells.EntireColumn.AutoFit

'
End Sub






Sub major_processed_to_standard()
'
' major_processed_to_standard Macro
'

  ' checking sheet
  
    Dim response As String
    
    If (Not (Range("A1") = "")) Or (Range("D3") = 0) Or (Range("D10") = 0) Then
        response = MsgBox("Check that active sheet is major minerals already processed by raw. If you selected the wrong sheet, please hit cancel.", vbOKCancel, "Correct Sheet?")
        If response = vbCancel Then Exit Sub
    End If
    

  ' general var dec
  
    Dim countCol As Integer, countRow As Integer, i As Integer, a As String, b As String, ii As Integer, n As Integer, j As Integer, k As Integer

  ' find last row & column of processed sheet, stores as variables
       
    Dim LastRow As Long, LastCol As Long
    
    ActiveSheet.UsedRange
    LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.count).Row
    LastRow = LastRow

    ActiveSheet.UsedRange
    LastCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.count).Column
    
 
  ' setting up array & adding first 2 values
   
     Dim arr()
     
     ReDim arr(2, 1)
     
     arr(0, 0) = "Method"
     arr(1, 0) = "Standard/Rep"
     arr(2, 0) = "n"

 
  ' custom list, needed for sorting
  
    Dim elementsList(1 To 10) As String
    elementsList(1) = "SiO2"
    elementsList(2) = "TiO2"
    elementsList(3) = "Al2O3"
    elementsList(4) = "Fe2O3"
    elementsList(5) = "MnO"
    elementsList(6) = "MgO"
    elementsList(7) = "CaO"
    elementsList(8) = "Na2O"
    elementsList(9) = "K2O"
    elementsList(10) = "P2O5"
    
    
  ' array for all stds, elements, and values (need multiple dimensions, list not sufficient)
  
    Dim standardsList(9, 10, 1) As Variant
    
  ' loads up all standards with elements
  
    For i = 0 To 9
        For j = 1 To 10
            standardsList(i, j, 0) = elementsList(j)
        Next
    Next
    
    standardsList(0, 0, 0) = "AGV-2"
    standardsList(0, 1, 1) = "59.3"
    standardsList(0, 2, 1) = "1.05"
    standardsList(0, 3, 1) = "16.91"
    standardsList(0, 4, 1) = "6.69"
    standardsList(0, 5, 1) = "0.099"
    standardsList(0, 6, 1) = "1.79"
    standardsList(0, 7, 1) = "5.2"
    standardsList(0, 8, 1) = "4.19"
    standardsList(0, 9, 1) = "2.88"
    standardsList(0, 10, 1) = "0.48"
    
    standardsList(1, 0, 0) = "BCR-2"
    standardsList(1, 1, 1) = "54.1"
    standardsList(1, 2, 1) = "2.26"
    standardsList(1, 3, 1) = "13.5"
    standardsList(1, 4, 1) = "13.8"
    standardsList(1, 5, 1) = "0.196"
    standardsList(1, 6, 1) = "3.59"
    standardsList(1, 7, 1) = "7.12"
    standardsList(1, 8, 1) = "3.16"
    standardsList(1, 9, 1) = "1.79"
    standardsList(1, 10, 1) = "0.35"
    
    standardsList(2, 0, 0) = "BHVO-2"
    standardsList(2, 1, 1) = "49.9"
    standardsList(2, 2, 1) = "2.73"
    standardsList(2, 3, 1) = "13.5"
    standardsList(2, 4, 1) = "12.3"
    standardsList(2, 5, 1) = "0.167"
    standardsList(2, 6, 1) = "7.23"
    standardsList(2, 7, 1) = "11.4"
    standardsList(2, 8, 1) = "2.22"
    standardsList(2, 9, 1) = "0.52"
    standardsList(2, 10, 1) = "0.27"
    
    standardsList(3, 0, 0) = "BIR-1a"
    standardsList(3, 1, 1) = "47.96"
    standardsList(3, 2, 1) = "0.96"
    standardsList(3, 3, 1) = "15.5"
    standardsList(3, 4, 1) = "11.3"
    standardsList(3, 5, 1) = "0.175"
    standardsList(3, 6, 1) = "9.7"
    standardsList(3, 7, 1) = "13.3"
    standardsList(3, 8, 1) = "1.82"
    standardsList(3, 9, 1) = "0.03"
    standardsList(3, 10, 1) = "0.021"
    
    standardsList(4, 0, 0) = "GSP-2"
    standardsList(4, 1, 1) = "66.6"
    standardsList(4, 2, 1) = "0.66"
    standardsList(4, 3, 1) = "14.9"
    standardsList(4, 4, 1) = "4.9"
    standardsList(4, 5, 1) = "0.15"
    standardsList(4, 6, 1) = "0.96"
    standardsList(4, 7, 1) = "2.1"
    standardsList(4, 8, 1) = "2.78"
    standardsList(4, 9, 1) = "5.38"
    standardsList(4, 10, 1) = "0.29"
    
    standardsList(5, 0, 0) = "W-2a"
    standardsList(5, 1, 1) = "52.68"
    standardsList(5, 2, 1) = "1.06"
    standardsList(5, 3, 1) = "15.45"
    standardsList(5, 4, 1) = "10.83"
    standardsList(5, 5, 1) = "0.167"
    standardsList(5, 6, 1) = "6.37"
    standardsList(5, 7, 1) = "10.86"
    standardsList(5, 8, 1) = "2.2"
    standardsList(5, 9, 1) = "0.626"
    standardsList(5, 10, 1) = "0.14"
    
    standardsList(6, 0, 0) = "SDC-1"
    standardsList(6, 1, 1) = "65.8"
    standardsList(6, 2, 1) = "1.01"
    standardsList(6, 3, 1) = "15.8"
    standardsList(6, 4, 1) = "6.32"
    standardsList(6, 5, 1) = "0.11"
    standardsList(6, 6, 1) = "1.69"
    standardsList(6, 7, 1) = "1.4"
    standardsList(6, 8, 1) = "2.05"
    standardsList(6, 9, 1) = "3.28"
    standardsList(6, 10, 1) = "0.16"
    
    standardsList(7, 0, 0) = "DNC-1a"
    standardsList(7, 1, 1) = "47.15"
    standardsList(7, 2, 1) = "0.48"
    standardsList(7, 3, 1) = "18.34"
    standardsList(7, 4, 1) = "9.97"
    standardsList(7, 5, 1) = "0.15"
    standardsList(7, 6, 1) = "10.13"
    standardsList(7, 7, 1) = "11.49"
    standardsList(7, 8, 1) = "1.89"
    standardsList(7, 9, 1) = "0.234"
    standardsList(7, 10, 1) = "0.07"
    
    standardsList(8, 0, 0) = "688" ' NIST-688 --> 688, if reported as NIST- or -Basalt should work?
    standardsList(8, 1, 1) = "48.4"
    standardsList(8, 2, 1) = "1.17"
    standardsList(8, 3, 1) = "17.36"
    standardsList(8, 4, 1) = "10.35"
    standardsList(8, 5, 1) = "0.167"
    standardsList(8, 6, 1) = "8.4"
    standardsList(8, 7, 1) = "12.17"
    standardsList(8, 8, 1) = "2.15"
    standardsList(8, 9, 1) = "0.187"
    standardsList(8, 10, 1) = "0.134"
    
 
    'add All92, COQ-1, DTS-2, G-2, NIST_278, NIST_70A, QLO, SCO-1, SGR-1?
    'standardsList(9, 0, 0) = "UND-E-11-03"?
    'no certified?
    
    
    
  ' finding all oxides for certified
      
    countCol = 4
 
    For i = 4 To LastRow
        For ii = 1 To 10
            If (Cells(1, i) = elementsList(ii)) Then
                arr = ReDimPreserve(arr, countCol, countCol)
                arr(countCol, 0) = Cells(1, i)
                countCol = countCol + 1
            End If
        Next
    Next
    
    
  ' begin adding the sample & averages/std dev/certified/%error/%rsd to array
    
    countRow = 1
    n = 1
    
    For i = 3 To LastRow
        a = Range("A" & i)
        b = Range("A" & i + 1)
        
        If (StrComp(a, b, vbTextCompare) = 0) Then ' if the texts are the same
            n = n + 1
        End If
        
        If (StrComp(a, b, vbTextCompare)) Then
            If Not (StrComp(b, Range("A" & i + 2), vbTextCompare)) Then ' if the samples are not used in the standard, and will be ignored by next for loop, set the counter back to one
                n = 1
            End If
            
            For j = 0 To 8
            If (a Like ("*" & standardsList(j, 0, 0) & "*")) Then ' checking for the std, not element
                arr = ReDimPreserve(arr, countCol, countRow + 4)
                arr(0, countRow) = "XRF"
                If (a) = (standardsList(j, 0, 0) & "Dup") Then
                    arr(1, countRow) = standardsList(j, 0, 0) & " Dup" 'adds std name if includes Dup
                Else: arr(1, countRow) = standardsList(j, 0, 0) ' else adds normal std name
                End If
                arr(3, countRow) = "Average"
                arr(3, countRow + 1) = "Std dev"
                arr(3, countRow + 2) = "Certified"
                arr(3, countRow + 3) = "%Error"
                arr(3, countRow + 4) = "%RSD"
                For ii = 3 To 12
                    arr(ii + 1, countRow) = (Cells(i + 1, ii + 1)) 'add avg/mean
                    arr(ii + 1, countRow + 1) = (Cells(i + 2, ii + 1)) 'add std dev
                    arr(ii + 1, countRow + 4) = (Cells(i + 3, ii + 1)) ' add %RSD
                    For k = 0 To 9
                        If (arr(ii + 1, 0) = standardsList(j, k + 1, 0)) Then
                            arr(ii + 1, countRow + 2) = standardsList(j, k + 1, 1)
                            arr(ii + 1, countRow + 3) = Round(Abs(100 * ((arr(ii + 1, countRow) - arr(ii + 1, countRow + 2)) / (arr(ii + 1, countRow + 2)))), 2)
                            arr(ii + 1, countRow) = Round(arr(ii + 1, countRow), 2)
                            arr(ii + 1, countRow + 1) = Round(arr(ii + 1, countRow + 1), 2)
                            arr(ii + 1, countRow + 4) = Round(arr(ii + 1, countRow + 4), 2)
                        End If
                    Next
                Next
                arr(2, countRow) = n
                n = 1
                countRow = countRow + 6
                i = i + 3
            
            End If
            Next
        End If

    Next

  ' new sheet!
  
     Dim Name As String, cmon As Boolean
    
    cmon = True
    
    Name = "Standards"
    
    For Each Sheet In Worksheets
        If (Name = Sheet.Name) And (Name = "Standards") Then Name = "Standards1"
        If (("Standards1" = Sheet.Name) And ((Name = "Standards1") Or (Name = "Standards"))) Then Name = "Standards2"
        If ("Standards2" = Sheet.Name) Then
            response = MsgBox("C'mon, why do you need so many standards sheets?                            (Right-Clicking on a sheet's tab at the bottom will allow you to rename or delete that sheet)", 0, "What do you need all these for?")
            Name = "Standards3"
        End If
        If ("Standards3" = Sheet.Name) Then cmon = False
    Next
    
    Sheets.Add After:=ActiveSheet
    If cmon Then ActiveSheet.Name = Name
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 0
    End With
    ActiveWindow.FreezePanes = True
     
  ' adding array to new sheet
    
    Dim numRow As Long, numCol As Long, e As Integer
    
    e = 1
    numRows = UBound(arr)
    numCols = UBound(arr, 2)
    
    With ActiveSheet.Range("A1:AZ1")
        .Range(.Cells(1, 1), .Cells(1, (numCols + 1))).Resize(numRows).Value = arr
    End With
    
    
  ' sort it
    
    Application.AddCustomList ListArray:=elementsList
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Range("A5:BZ20").Sort , Key1:=Range("A:A"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=Application.CustomListCount + 1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    ActiveSheet.Sort.SortFields.Clear
    Application.DeleteCustomList Application.CustomListCount
    
    
  ' autofits header --> makes it too compressed?
    'Cells.EntireColumn.AutoFit

    

'
End Sub



Sub trace_processed_to_standards()
'
'
'


  ' checking sheet
    
    Dim response As String
    
    If (Not (Range("A1") = "")) Or (Range("D3") = 0) Or (Range("D10") = 0) Then
        response = MsgBox("Check that active sheet is trace minerals already processed by raw. If you selected the wrong sheet, please hit cancel.", vbOKCancel, "Correct Sheet?")
        If response = vbCancel Then Exit Sub
    End If
    
  ' general var dec

    Dim countCol As Integer, countRow As Integer, i As Integer, a As String, b As String, ii As Integer, n As Integer, j As Integer, k As Integer, countPPM As Integer
    

  ' find last row & column of processed sheet, stores as variables
       
    Dim LastRow As Long, LastCol As Long
    
    ActiveSheet.UsedRange
    LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.count).Row
    LastRow = LastRow

    ActiveSheet.UsedRange
    LastCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.count).Column
    
 
  ' setting up array & adding first 2 values
   
     Dim arr()
     
     ReDim arr(2, 1)
     
     arr(0, 0) = "Method"
     arr(1, 0) = "Standard/Rep"
     arr(2, 0) = "n"

 
  ' custom list, needed for sorting
  
   Dim traceList(1 To 12) As String
    traceList(1) = "V"
    traceList(2) = "Cr"
    traceList(3) = "Ni"
    traceList(4) = "Cu"
    traceList(5) = "Zn"
    traceList(6) = "Rb"
    traceList(7) = "Sr"
    traceList(8) = "Y"
    traceList(9) = "Zr"
    traceList(10) = "Nb"
    traceList(11) = "Ba"
    traceList(12) = "Pb"
    
    
  ' array for all stds, elements, and values (need multiple dimensions, list not sufficient)
  
    Dim standardsList(9, 12, 1) As Variant
    
  ' loads up all standards with elements
  
    For i = 0 To 9
        For j = 1 To 12
            standardsList(i, j, 0) = traceList(j)
        Next
    Next
    
    standardsList(0, 0, 0) = "AGV-2"
    standardsList(0, 1, 1) = "120"
    standardsList(0, 2, 1) = "17"
    standardsList(0, 3, 1) = "19"
    standardsList(0, 4, 1) = "53"
    standardsList(0, 5, 1) = "86"
    standardsList(0, 6, 1) = "69"
    standardsList(0, 7, 1) = "658"
    standardsList(0, 8, 1) = "20"
    standardsList(0, 9, 1) = "230"
    standardsList(0, 10, 1) = "15"
    standardsList(0, 11, 1) = "1140"
    standardsList(0, 12, 1) = "13"
    
    standardsList(1, 0, 0) = "BCR-2"
    standardsList(1, 1, 1) = "416"
    standardsList(1, 2, 1) = "18"
    standardsList(1, 3, 1) = ""
    standardsList(1, 4, 1) = "19"
    standardsList(1, 5, 1) = "127"
    standardsList(1, 6, 1) = "48"
    standardsList(1, 7, 1) = "346"
    standardsList(1, 8, 1) = "37"
    standardsList(1, 9, 1) = "188"
    standardsList(1, 10, 1) = ""
    standardsList(1, 11, 1) = "683"
    standardsList(1, 12, 1) = ""
    
    
    standardsList(2, 0, 0) = "BHVO-2"
    standardsList(2, 1, 1) = "317"
    standardsList(2, 2, 1) = "280"
    standardsList(2, 3, 1) = "119"
    standardsList(2, 4, 1) = "127"
    standardsList(2, 5, 1) = "103"
    standardsList(2, 6, 1) = "10"
    standardsList(2, 7, 1) = "389"
    standardsList(2, 8, 1) = "26"
    standardsList(2, 9, 1) = "172"
    standardsList(2, 10, 1) = "18"
    standardsList(2, 11, 1) = "130"
    standardsList(2, 12, 1) = ""
    
    standardsList(3, 0, 0) = "BIR-1a"
    standardsList(3, 1, 1) = "310"
    standardsList(3, 2, 1) = "370"
    standardsList(3, 3, 1) = "170"
    standardsList(3, 4, 1) = "125"
    standardsList(3, 5, 1) = "70"
    standardsList(3, 6, 1) = ""
    standardsList(3, 7, 1) = "110"
    standardsList(3, 8, 1) = "16"
    standardsList(3, 9, 1) = "18"
    standardsList(3, 10, 1) = ""
    standardsList(3, 11, 1) = "109.52"
    standardsList(3, 12, 1) = ""
    
    standardsList(4, 0, 0) = "GSP-2"
    standardsList(4, 1, 1) = "52"
    standardsList(4, 2, 1) = "20"
    standardsList(4, 3, 1) = "17"
    standardsList(4, 4, 1) = "43"
    standardsList(4, 5, 1) = "120"
    standardsList(4, 6, 1) = "245"
    standardsList(4, 7, 1) = "240"
    standardsList(4, 8, 1) = "28"
    standardsList(4, 9, 1) = "550"
    standardsList(4, 10, 1) = "27"
    standardsList(4, 11, 1) = "1340"
    standardsList(4, 12, 1) = "42"
    
    standardsList(5, 0, 0) = "W-2a"
    standardsList(5, 1, 1) = "260"
    standardsList(5, 2, 1) = "92"
    standardsList(5, 3, 1) = "70"
    standardsList(5, 4, 1) = "110"
    standardsList(5, 5, 1) = "80"
    standardsList(5, 6, 1) = "21"
    standardsList(5, 7, 1) = "190"
    standardsList(5, 8, 1) = "23"
    standardsList(5, 9, 1) = "100"
    standardsList(5, 10, 1) = "8"
    standardsList(5, 11, 1) = "170"
    standardsList(5, 12, 1) = ""
    
    standardsList(6, 0, 0) = "SDC-1"
    standardsList(6, 1, 1) = "102"
    standardsList(6, 2, 1) = "64"
    standardsList(6, 3, 1) = "38"
    standardsList(6, 4, 1) = "30"
    standardsList(6, 5, 1) = "103"
    standardsList(6, 6, 1) = "127"
    standardsList(6, 7, 1) = "180"
    standardsList(6, 8, 1) = ""
    standardsList(6, 9, 1) = "290"
    standardsList(6, 10, 1) = "21"
    standardsList(6, 11, 1) = "630"
    standardsList(6, 12, 1) = "25"
    
    standardsList(7, 0, 0) = "DNC-1a"
    standardsList(7, 1, 1) = "148"
    standardsList(7, 2, 1) = "270"
    standardsList(7, 3, 1) = "247"
    standardsList(7, 4, 1) = "100"
    standardsList(7, 5, 1) = "70"
    standardsList(7, 6, 1) = "4.5"
    standardsList(7, 7, 1) = "144"
    standardsList(7, 8, 1) = "18"
    standardsList(7, 9, 1) = "38"
    standardsList(7, 10, 1) = "3"
    standardsList(7, 11, 1) = "118"
    standardsList(7, 12, 1) = ""
    
    standardsList(8, 0, 0) = "688" ' NIST-688 --> 688, if reported as NIST- or -Basalt should work?
    standardsList(8, 1, 1) = "250"
    standardsList(8, 2, 1) = "332"
    standardsList(8, 3, 1) = ""
    standardsList(8, 4, 1) = "96"
    standardsList(8, 5, 1) = "58"
    standardsList(8, 6, 1) = "2"
    standardsList(8, 7, 1) = "169"
    standardsList(8, 8, 1) = ""
    standardsList(8, 9, 1) = ""
    standardsList(8, 10, 1) = ""
    standardsList(8, 12, 1) = "200"
    standardsList(8, 12, 1) = "3"
    
 
    'add All92, COQ-1, DTS-2, G-2, NIST_278, NIST_70A, QLO, SCO-1, SGR-1?
    'standardsList(9, 0, 0) = "UND-E-11-03"?
    'no certified?
    
    
    
  ' finding all oxides for certified
      
    countCol = 4
 
    For i = 4 To LastRow
        For ii = 1 To 10
            If (Cells(1, i) = traceList(ii)) Then
                arr = ReDimPreserve(arr, countCol, countCol)
                arr(countCol, 0) = Cells(1, i)
                countCol = countCol + 1
            End If
        Next
    Next
    
    
  ' begin adding the sample & averages/std dev/certified/%error/%rsd to array
    
    countRow = 1
    n = 1
    countPPM = 1
    
    For i = 3 To LastRow
        a = Range("A" & i)
        b = Range("A" & i + 1)
        
        If (StrComp(a, b, vbTextCompare) = 0) Then ' if the texts are the same
            n = n + 1
        End If
        
        If (StrComp(a, b, vbTextCompare)) Then
            If Not (StrComp(b, Range("A" & i + 2), vbTextCompare)) Then ' if the samples are not used in the standard, and will be ignored by next for loop, set the counter back to one
                n = 1
            End If
            
            For j = 0 To 8
            If (a Like ("*" & standardsList(j, 0, 0) & "*")) Then ' checking for the std, not element
                arr = ReDimPreserve(arr, countCol, countRow + 4)
                arr(0, countRow) = "XRF"
                If (a) = (standardsList(j, 0, 0) & "Dup") Then
                    arr(1, countRow) = standardsList(j, 0, 0) & " Dup" 'adds std name if includes Dup
                Else: arr(1, countRow) = standardsList(j, 0, 0) ' else adds normal std name
                End If
                arr(3, countRow) = "Average"
                arr(3, countRow + 1) = "Std dev"
                arr(3, countRow + 2) = "Certified"
                arr(3, countRow + 3) = "%Error"
                arr(3, countRow + 4) = "%RSD"
                For ii = 3 To LastCol
                If (Cells(2, ii + 1) = "ppm") Then
                If countPPM < 12 Then
                If (Cells(1, ii + 1) = traceList(countPPM)) Then
                'For countPPM = 3 To 12
                    countPPM = countPPM + 1
                    arr(countPPM + 2, countRow) = (Cells(i + 1, ii + 1)) 'add avg/mean
                    arr(countPPM + 2, countRow + 1) = (Cells(i + 2, ii + 1)) 'add std dev
                    arr(countPPM + 2, countRow + 4) = (Cells(i + 3, ii + 1)) ' add %RSD
                    For k = 0 To 9
                        If (arr(countPPM + 2, 0) = standardsList(j, k + 1, 0)) Then
                            arr(countPPM + 2, countRow + 2) = standardsList(j, k + 1, 1)
                            If Not (arr(countPPM + 2, countRow + 2) = "") Then
                                arr(countPPM + 2, countRow + 3) = Round(Abs(100 * ((arr(countPPM + 2, countRow) - arr(countPPM + 2, countRow + 2)) / (arr(countPPM + 2, countRow + 2)))), 2)
                            End If
                            arr(countPPM + 2, countRow) = Round(arr(countPPM + 2, countRow), 2)
                            arr(countPPM + 2, countRow + 1) = Round(arr(countPPM + 2, countRow + 1), 2)
                            arr(countPPM + 2, countRow + 4) = Round(arr(countPPM + 2, countRow + 4), 2)
                        End If
                    Next
                'Next
                End If
                End If
                End If
                Next
                
                arr(2, countRow) = n
                n = 1
                countRow = countRow + 6
                i = i + 3
                countPPM = 1
            
           ' End If
            End If
            Next
        End If

    Next

  ' new sheet!
  
    Dim Name As String, cmon As Boolean
    
    cmon = True
    
    Name = "Standards"
    
    For Each Sheet In Worksheets
        If (Name = Sheet.Name) And (Name = "Standards") Then Name = "Standards1"
        If (("Standards1" = Sheet.Name) And ((Name = "Standards1") Or (Name = "Standards"))) Then Name = "Standards2"
        If ("Standards2" = Sheet.Name) Then
            response = MsgBox("C'mon, why do you need so many standards sheets?                            (Right-Clicking on a sheet's tab at the bottom will allow you to rename or delete that sheet)", 0, "What do you need all these for?")
            Name = "Standards3"
        End If
        If ("Standards3" = Sheet.Name) Then cmon = False
    Next
    
    Sheets.Add After:=ActiveSheet
    If cmon Then ActiveSheet.Name = Name
    
    With ActiveWindow
        .SplitColumn = 1
        .SplitRow = 0
    End With
    ActiveWindow.FreezePanes = True
     
  ' adding array to new sheet
    
    Dim numRow As Long, numCol As Long, e As Integer
    
    e = 1
    numRows = UBound(arr)
    numCols = UBound(arr, 2)
    
    With ActiveSheet.Range("A1:AZ1")
        .Range(.Cells(1, 1), .Cells(1, (numCols + 1))).Resize(numRows).Value = arr
    End With
    
    
  ' sort it
    
    Application.AddCustomList ListArray:=traceList
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Range("A5:BZ20").Sort , Key1:=Range("A:A"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=Application.CustomListCount + 1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    ActiveSheet.Sort.SortFields.Clear
    Application.DeleteCustomList Application.CustomListCount
    
    
  ' autofits header --> makes it too compressed?
    'Cells.EntireColumn.AutoFit

'
End Sub



Sub mjr_raw_to_graph()

'
'macro presents userform to determine # of graphs (one or all) and what should be plotted.
'data taken from form and current sheet (raw) and transferred to new sheet as scatterplots
'

Dim xele As String, yele As String, response As String, totalCol As Integer, totalRow As Integer, i As Integer, yn As String
Dim sing As Boolean, many As Boolean, xCol As Integer, yCol As Integer, count As Integer, ii As Integer



ActiveSheet.UsedRange
totalCol = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.count).Column
totalRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.count).Row


'get info for # and elements of graph

graphForm.Show
sing = graphForm.radioSingle.Value
many = graphForm.radioMany.Value
xele = graphForm.xeleInput.Value
yele = graphForm.yeleInput.Value


'make sure they're actual elements included in data provided, saving columns as ints

xCol = 0
yCol = 0

For i = 4 To totalCol
    If (Cells(1, i) = xele) Then
        xCol = i
    ElseIf (Cells(1, i) = yele) Then
        yCol = i
    End If
Next i

If (sing = True) Then
    If (yCol = 0) Then
        MsgBox (yele & " is not a recognized element. Ending macro, please try again.")
        Exit Sub
    End If
End If

If (xCol = 0) Then
    MsgBox (xele & " is not a recognized element. Ending macro, please try again.")
    Exit Sub
End If


'copy into -array- ranges


'xele values into xrng for both types (sing & many), ymrng to hold different yele values as yrng (ranges)
Dim xrng As Range, yrng As Range, ymrng As Range

' an arr to keep track of ele names - x axis as arr(0), each
Dim arr()
ReDim arr(totalCol)


Set xrng = Cells(3, xCol).Resize(totalRow - 2)

count = 0

'single yele
If (sing = True) Then
    Set yrng = Cells(3, yCol).Resize(totalRow - 2) ' needed?
    Set ymrng = Cells(3, yCol).Resize(totalRow - 2)
    arr(1) = Cells(1, yCol)
    
'many yele
ElseIf (many = True) Then
    'a range of ranges
    
    For i = 4 To totalCol
        If ((Not (Cells(1, i) = "")) And (Not (Cells(1, i) = xele))) Then
            count = count + 1
            yele = Cells(1, i)
            arr(count) = yele
            Set yrng = Cells(3, i).Resize(totalRow - 2)
        
            If ymrng Is Nothing Then
                Set ymrng = Cells(3, i).Resize(totalRow - 2)
            Else: Set ymrng = Union(ymrng, yrng)
            End If
            
        End If
    Next i
End If



'new sheet, create graph(s)

Sheets.Add After:=ActiveSheet

If (sing = True) Then ActiveSheet.Name = xele & "-" & yele
If (many = True) Then ActiveSheet.Name = xele


Dim x1, x2, y1, y2, c2 As Integer
x1 = 50
y1 = 40
x2 = 500
y2 = 250
c2 = 1

For Each yrng In ymrng.Areas

    Dim Chart1 As ChartObject
    Set Chart1 = ActiveSheet.ChartObjects.Add(x1, y1, x2, y2)
    
    With Chart1.Chart
    
        .HasTitle = True
        .ChartTitle.Text = xele & "-" & arr(c2)
        .ChartType = xlXYScatter
        .SeriesCollection.NewSeries
        
        .SeriesCollection(1).XValues = xrng
        .SeriesCollection(1).Values = yrng
        
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Caption = xele & " Mass%"
        
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Caption = arr(c2) & " Mass%"
        
        .SetElement (msoElementLegendNone)
          
    End With
    
    c2 = c2 + 1
    x1 = x1 + 550

Next yrng


End Sub


