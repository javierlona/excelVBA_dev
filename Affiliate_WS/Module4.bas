Attribute VB_Name = "Module4"
Public MyValOCEPC As Integer
Public netWorthAffTotal As Double
Public netProfit1AffTotal As Double
Public netProfit2AffTotal As Double
Public netWorthOCEPCTotal As Double
Public netProfit1OCEPCTotal As Double
Public netProfit2OCEPCTotal As Double
Public avgNetProfitOCEPCTotal As Double
Public avgNetProfitAffTotal As Double
Sub showNS(ByVal string1 As String, ByVal int1 As Integer)
    MsgBox "Procedure ShowNS " & string1 & " " & int1
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim strSearch As String
    Dim aCell As Range
    Dim rowNumber As Integer
    
    MyValOCEPC = int1
    
    Set ws = Sheets("Sheet1")
    lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
    strSearch = string1
    
    Set aCell = ws.Range("B7:B" & lastRow).Find(What:=strSearch, LookIn:=xlValues, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)
    If Not aCell Is Nothing Then
        'MsgBox "Value Found in Cell " & aCell.Address
        aCell.EntireRow.Select
        Selection.Offset(2, 0).Select
        rowNumber = Selection.Row
    End If
    
    'MyIB1 = InputBox("How Many EPCs and/or OCs?")
    'For x = 1 To MyIB1
    
    For x = 1 To MyValOCEPC
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove 'additional row for dates
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("B" & CStr(rowNumber), "K" & CStr(rowNumber) + 7).Select 'changed from "J" to "K" and 6 to 7
        'Selection.ClearFormats
        Selection.Interior.ColorIndex = 0
        Selection.Font.Bold = False
        Selection.EntireRow.AutoFit
        
        'this line can be used to describe date line
        Range("G" & CStr(rowNumber), "K" & CStr(rowNumber)).Select 'changed from'changed from "E" to "F"
        Selection.HorizontalAlignment = xlCenter
        Selection.NumberFormat = "@" 'format cells to text
        Range("G" & CStr(rowNumber)) = "Financial"
        Range("H" & CStr(rowNumber)) = "20??"
        Range("I" & CStr(rowNumber)) = "20??"
        Range("J" & CStr(rowNumber)) = "20??"
        'This portion creates a new OC or EPC
        Range("B" & CStr(rowNumber) + 1, "G" & CStr(rowNumber) + 1).Select 'changed from'changed from "E" to "G"
        Selection.Interior.Color = 8388608
        Range("B" & CStr(rowNumber) + 1) = "(EPC NAME or OC if applicable)"
        Selection.Font.Size = 14
        Selection.Font.Bold = True
        Selection.Font.Name = "Tahoma"
        Selection.Font.ThemeColor = xlThemeColorDark1
        'This portion merges single cell for entity name
        Range("B" & CStr(rowNumber) + 1, "C" & CStr(rowNumber) + 1).Select
        Selection.Merge
        'portion formats the owner cell formating above the total
        Range("B" & CStr(rowNumber) + 2, "K" & CStr(rowNumber) + 5).Select 'changed from "J" to "K"; and 4 to 5?
        With Selection.Font
            .Name = "Tahoma"
            .Size = 14
            .Color = vbBlack
        End With
        'centers EIN no., Title, Guarantor, % Ownership
        Range("D" & CStr(rowNumber) + 1, "G" & CStr(rowNumber) + 6).Select 'changed "F" to "G"
        Selection.HorizontalAlignment = xlCenter
        
        'Wordwrap the owner section
        Range("C" & CStr(rowNumber) + 2, "C" & CStr(rowNumber) + 4).Select 'these #'s would have to increase by one
        Selection.wraptext = True
        'write content
        Range("B" & CStr(rowNumber) + 2) = "Owner #1"
        Range("B" & CStr(rowNumber) + 3) = "Owner #2"
        Range("B" & CStr(rowNumber) + 4) = "Owner #3"
        Range("B" & CStr(rowNumber) + 5) = ""
        Range("B" & CStr(rowNumber) + 6) = "Total"
        Range("B" & CStr(rowNumber) + 6).Font.Bold = True
        Range("G" & CStr(rowNumber) + 2, "G" & CStr(rowNumber) + 4).Select 'change "F" to "G"
        Selection.NumberFormat = "0.00%"
        Range("G" & CStr(rowNumber) + 6) = "=SUM(" & Range("G" & CStr(rowNumber) + 2, "G" & CStr(rowNumber) + 4).Address(False, False) & ")"
        Range("G" & CStr(rowNumber) + 6).Font.Bold = True
        Range("B" & CStr(rowNumber) + 2, "G" & CStr(rowNumber) + 6).Select
        With Selection.Borders
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        'creates the total slash for each individual section
        Range("G" & CStr(rowNumber) + 6).Select
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        'creates comment section
        Range("H" & CStr(rowNumber) + 2, "K" & CStr(rowNumber) + 7).Select 'changed "G" to H"
        For Each border In Selection.Borders
            border.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        Next
        Range("H" & CStr(rowNumber) + 2) = "Comments:"
        Range("H" & CStr(rowNumber) + 2, "K" & CStr(rowNumber) + 7).Select
        Selection.Merge
        'italisize here
        Selection.Font.FontStyle = "Italic"
        Selection.VerticalAlignment = xlTop
        Selection.HorizontalAlignment = xlLeft
        With Selection
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End With
        
        'this creates the number section entry for epc fin info
        'Net Worth, TR Net Profit, TR Net Profit
        Range("H" & CStr(rowNumber) + 1, "K" & CStr(rowNumber) + 1).Select
        Selection.HorizontalAlignment = xlCenter
        Selection.NumberFormat = "$#,##0.0"
        With Selection.Borders
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Range("K" & CStr(rowNumber) + 1) = "=SUM(RC[-2]:RC[-1])/2"
        Range("K" & CStr(rowNumber) + 1).Font.Bold = True
        'ready for the next entry
        aCell.EntireRow.Select
        Selection.Offset(2, 0).Select
        rowNumber = Selection.Row
    Next x
End Sub
Sub calculateSums()

Dim ws As Worksheet
Dim lastRow As Long, i As Long
Dim strSearch As String
Dim strSearch2 As String
Dim strSearch3 As String
Dim strSearch4 As String
Dim strSearch5 As String
Dim strSearch6 As String
Dim strSearch7 As String
Dim aCell As Range
Dim rowNumber As Integer
Dim epcocrowa As Integer
Dim epcocrowb As Integer
Dim affiliatea As Integer
Dim affiliateb As Integer

Set ws = Sheets("Sheet1")
lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
strSearch = "NAME OF OC and/or EPC ENTITIES"
strSearch2 = "NAME OF GUARANTOR AFFILIATES"
strSearch3 = "TOTAL EPC AND OC"
strSearch4 = "NAME OF GUARANTOR AFFILIATES"
strSearch5 = "TOTAL EPC AND OC"
strSearch6 = "TOTAL AFFILIATES"
strSearch7 = "GRAND TOTAL"

'calculates epc and oc entities
Set aCell = ws.Range("B7:B" & lastRow).Find(What:=strSearch, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
If Not aCell Is Nothing Then
    'MsgBox "Value Found in Cell " & aCell.Address
    aCell.EntireRow.Select
    epcocrowa = Selection.Row
End If

Set aCell = ws.Range("B7:B" & lastRow).Find(What:=strSearch2, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
If Not aCell Is Nothing Then
    'MsgBox "Value Found in Cell " & aCell.Address
    aCell.EntireRow.Select
    epcocrowb = Selection.Row
End If

Set aCell = ws.Range("B7:B" & lastRow).Find(What:=strSearch3, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
If Not aCell Is Nothing Then
    'MsgBox "Value Found in Cell " & aCell.Address
    aCell.EntireRow.Select
    rowNumber = Selection.Row
    'set cell value equal to variable
    Range("H" & CStr(rowNumber)).Value = netWorthOCEPCTotal
    Range("I" & CStr(rowNumber)).Value = netProfit1OCEPCTotal
    Range("J" & CStr(rowNumber)).Value = netProfit2OCEPCTotal
    Range("K" & CStr(rowNumber)).Value = avgNetProfitOCEPCTotal
End If

'calculates affiliates
Set aCell = ws.Range("B7:B" & lastRow).Find(What:=strSearch4, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
If Not aCell Is Nothing Then
    'MsgBox "Value Found in Cell " & aCell.Address
    aCell.EntireRow.Select
    affiliatea = Selection.Row
End If

Set aCell = ws.Range("B7:B" & lastRow).Find(What:=strSearch5, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
If Not aCell Is Nothing Then
    'MsgBox "Value Found in Cell " & aCell.Address
    aCell.EntireRow.Select
    affiliateb = Selection.Row
End If

Set aCell = ws.Range("B7:B" & lastRow).Find(What:=strSearch6, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
If Not aCell Is Nothing Then
    'MsgBox "Value Found in Cell " & aCell.Address
    aCell.EntireRow.Select
    rowNumber = Selection.Row
    'set cell value equal to variable
    Range("H" & CStr(rowNumber)).Value = netWorthAffTotal
    Range("I" & CStr(rowNumber)).Value = netProfit1AffTotal
    Range("J" & CStr(rowNumber)).Value = netProfit2AffTotal
    Range("K" & CStr(rowNumber)).Value = avgNetProfitAffTotal
End If

Set aCell = ws.Range("B7:B" & lastRow).Find(What:=strSearch7, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
If Not aCell Is Nothing Then
    'MsgBox "Value Found in Cell " & aCell.Address
    aCell.EntireRow.Select
    rowNumber = Selection.Row
    'grand total row
    Range("H" & CStr(rowNumber)) = "=SUM(" & Range("H" & CStr(rowNumber) - 2, "H" & CStr(rowNumber) - 1).Address(False, False) & ")"
    Range("I" & CStr(rowNumber)) = "=SUM(" & Range("I" & CStr(rowNumber) - 2, "I" & CStr(rowNumber) - 1).Address(False, False) & ")"
    Range("J" & CStr(rowNumber)) = "=SUM(" & Range("J" & CStr(rowNumber) - 2, "J" & CStr(rowNumber) - 1).Address(False, False) & ")"
    Range("K" & CStr(rowNumber)) = "=SUM(" & Range("K" & CStr(rowNumber) - 2, "K" & CStr(rowNumber) - 1).Address(False, False) & ")"
    
End If
End Sub
Sub calculateNetWorthNetProfit()

Dim lColor As Long
Dim ws As Worksheet
Dim lastRow As Long, i As Long
Dim strSearch As String
Dim aCell As Range
Dim bCell As Range
Dim rowNumber As Integer
Dim netWorthAffEntry As Double
Dim netProfit1AffEntry As Double
Dim netProfit2AffEntry As Double
Dim netWorthOCEPCEntry As Double
Dim netProfit1OCEPCEntry As Double
Dim netProfit2OCEPCEntry As Double
Dim avgNetProfitOCEPCEntry As Double
Dim avgNetProfitAffEntry As Double
Dim lastRowforOCEPC As Integer

'error handling
On Error GoTo ErrHandler:
ErrHandler:
    If Err.Number = 13 Then 'Type Mismatch
         mistake = MsgBox("You typed text in a cell field(s) which only accepts numbers." & vbCr & vbCr & "Check the cell fields to the right of the blue rows.", 48, "Type Mismatch Error")
        Exit Sub
    End If

Set ws = Sheets("Sheet1")
lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
strSearch = "NAME OF GUARANTOR AFFILIATES"

Set aCell = ws.Range("B7:B" & lastRow).Find(What:=strSearch, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
If Not aCell Is Nothing Then
    aCell.EntireRow.Select
    Selection.Offset(2, 0).Select
    rowNumber = Selection.Row
    lastRowforOCEPC = rowNumber
End If

lColor = RGB(0, 0, 128)
Set MR = Range("B" & rowNumber, "B" & lastRow)

'reset global values
netWorthAffEntry = 0
netWorthAffTotal = 0
netProfit1AffEntry = 0
netProfit1AffTotal = 0
netProfit2AffEntry = 0
netProfit2AffTotal = 0
avgNetProfitAffEntry = 0
avgNetProfitAffTotal = 0

'Calculate Affiliates section
For Each aCell In MR
If aCell.Interior.Color = lColor Then
    aCell.Select
    rowNumber = Selection.Row
    aCell.Select
    Selection.Offset(0, 6).Select 'for net worth cells
    netWorthAffEntry = Cells(rowNumber, "H").Value
    netWorthAffTotal = netWorthAffEntry + netWorthAffTotal
    
    Selection.Offset(0, 1).Select 'for TR Net Profit 1st column cells
    netProfit1AffEntry = Cells(rowNumber, "I").Value
    netProfit1AffTotal = netProfit1AffEntry + netProfit1AffTotal
    
    Selection.Offset(0, 1).Select 'for TR Net Profit 2nd column cells
    netProfit2AffEntry = Cells(rowNumber, "J").Value
    netProfit2AffTotal = netProfit2AffEntry + netProfit2AffTotal
    
    Selection.Offset(0, 1).Select 'for TR AVG Net Profit 3rd column cells
    avgNetProfitAffEntry = Cells(rowNumber, "K").Value
    avgNetProfitAffTotal = avgNetProfitAffEntry + avgNetProfitAffTotal

End If
Next aCell

strSearch = "NAME OF OC and/or EPC ENTITIES"

Set bCell = ws.Range("B7:B" & lastRowforOCEPC).Find(What:=strSearch, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
If Not bCell Is Nothing Then
    bCell.EntireRow.Select
    Selection.Offset(2, 0).Select
    rowNumber = Selection.Row
End If

lColor = RGB(0, 0, 128)
Set MR = Range("B" & rowNumber, "B" & lastRowforOCEPC)

'reset global values
netWorthOCEPCEntry = 0
netWorthOCEPCTotal = 0
netProfit1OCEPCEntry = 0
netProfit1OCEPCTotal = 0
netProfit2OCEPCEntry = 0
netProfit2OCEPCTotal = 0
avgNetProfitOCEPCEntry = 0
avgNetProfitOCEPCTotal = 0

'Calculate OC/EPC section
For Each bCell In MR
If bCell.Interior.Color = lColor Then
    bCell.Select
    rowNumber = Selection.Row
    bCell.Select
    Selection.Offset(0, 6).Select 'for net worth cells
    netWorthOCEPCEntry = Cells(rowNumber, "H").Value
    netWorthOCEPCTotal = netWorthOCEPCEntry + netWorthOCEPCTotal
    
    Selection.Offset(0, 1).Select 'for TR Net Profit 1st colum cells
    netProfit1OCEPCEntry = Cells(rowNumber, "I").Value
    netProfit1OCEPCTotal = netProfit1OCEPCEntry + netProfit1OCEPCTotal
    
    Selection.Offset(0, 1).Select 'for TR Net Profit 2nd column cells
    netProfit2OCEPCEntry = Cells(rowNumber, "J").Value
    netProfit2OCEPCTotal = netProfit2OCEPCEntry + netProfit2OCEPCTotal
    
    Selection.Offset(0, 1).Select 'for TR AVG Net Profit 3rd column cells
    avgNetProfitOCEPCEntry = Cells(rowNumber, "K").Value
    avgNetProfitOCEPCTotal = avgNetProfitOCEPCEntry + avgNetProfitOCEPCTotal

End If
Next bCell

End Sub
Sub countentities()

Dim lColor As Long
Dim ws As Worksheet
Dim lastRow As Long, i As Long
Dim strSearch As String
Dim aCell As Range
Dim rowNumber As Integer
Dim counter As Integer

counter = 0

Set ws = Sheets("Sheet1")
lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
strSearch = "NAME OF GUARANTOR AFFILIATES"

Set aCell = ws.Range("B7:B" & lastRow).Find(What:=strSearch, LookIn:=xlValues, _
LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
MatchCase:=False, SearchFormat:=False)
If Not aCell Is Nothing Then
    'MsgBox "Value Found in Cell " & aCell.Address
    aCell.EntireRow.Select
    Selection.Offset(2, 0).Select
    rowNumber = Selection.Row
End If

lColor = RGB(0, 0, 128)
Set MR = Range("B" & rowNumber, "B" & lastRow)


'Set MR = Range("B7:B" & lastRow)

For Each aCell In MR
If aCell.Interior.Color = lColor Then
    'MsgBox "Value Found in Cell " & aCell.Address
    aCell.Select
    rowNumber = Selection.Row
    aCell.Select
    counter = counter + 1
    Selection.Offset(0, -1).Select
    With Selection.Font
        .Name = "Tahoma"
        .Size = 14
        .Bold = True
    End With
    Range("A" & CStr(rowNumber)) = counter

    'ActiveCell = counter.Value
End If
Next aCell
End Sub
