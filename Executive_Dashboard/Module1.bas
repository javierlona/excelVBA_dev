Attribute VB_Name = "Module1"
Public dsheet As Worksheet
Public rptsheet As Worksheet
Public Dashboard As Worksheet
Public StartDate As Date
Public EndDate As Date
Public BegJan As Date
Public EndJan As Date
Public BegFeb As Date
Public EndFeb As Date
Public BegMar As Date
Public EndMar As Date
Public BegApr As Date
Public EndApr As Date
Public BegMay As Date
Public EndMay As Date
Public BegJune As Date
Public EndJune As Date
Public BegJuly As Date
Public EndJuly As Date
Public BegAug As Date
Public EndAug As Date
Public BegSept As Date
Public EndSept As Date
Public BegOct As Date
Public EndOct As Date
Public BegNov As Date
Public EndNov As Date
Public BegDec As Date
Public EndDec As Date
'This is the main function which gives instructions
Sub myDashboardSummary()

'where all the info is going to be pulled from
Sheets(1).Name = "data"

'these functions create the layout of the worksheets
Call LiquidatedSheet
Call PaidoffSheet
Call AuthorizedSheet
Call FundedSheet
Call DashSummary

'these functions populates our worksheets
Call AuthorizedSheetData
Call FundedSheetData
Call CountPaidOff
Call CountLiquidating
Call DynamicSorting

'the last step which makes the dashboard summary visible once the macro is finished
Dashboard.Activate

End Sub
'This function prompts the user for the start and end dates.
'It then produces our dates to provide a breakdown by month.
Sub Datesdeclared()
Dim dashsum As Worksheet
Set dashsum = Sheets("Dashboard")
dashsum.Activate

StartDate = Application.InputBox("Enter A Start Date: ")
EndDate = Application.InputBox("Enter A End Date: ")

dashsum.Activate
Range("a9") = StartDate
Range("A9").Select
Range("A10") = "=EDATE(A9,1)"
Range("A10").Select
Selection.AutoFill Destination:=Range("A10:A20"), Type:=xlFillDefault
Range("A9:A20").Select
Selection.NumberFormat = "m/d/yyyy"
Range("b9") = "=EOMONTH(a9,0)"
Range("b10") = "=EOMONTH(b9,1)"
Range("b10").Select
Selection.AutoFill Destination:=Range("b10:b20"), Type:=xlFillDefault
End Sub
'this function declares our variables to provide a breakdown by month
Sub Datesdeclared2()
Dim dashsum As Worksheet
Set dashsum = Sheets("Dashboard")
dashsum.Activate
StartDate = Range("A9")
EndDate = Range("B20")
BegOct = Range("A9")
EndOct = Range("B9")
BegNov = Range("A10")
EndNov = Range("B10")
BegDec = Range("A11")
EndDec = Range("B11")
BegJan = Range("A12")
EndJan = Range("B12")
BegFeb = Range("A13")
EndFeb = Range("B13")
BegMar = Range("A14")
EndMar = Range("B14")
BegApr = Range("A15")
EndApr = Range("B15")
BegMay = Range("A16")
EndMay = Range("B16")
BegJune = Range("A17")
EndJune = Range("B17")
BegJuly = Range("A18")
EndJuly = Range("B18")
BegAug = Range("A19")
EndAug = Range("B19")
BegSept = Range("A20")
EndSept = Range("B20")

End Sub
'The function populates the Authorized sheet using the information from LMS
Sub AuthorizedSheetData()
Dim dsheet As Worksheet

Set dsheet = Sheets("data")
Set rptsheet = Sheets("Authorized")
Set dashsum = Sheets("Dashboard")

dsheet.Activate

Dim totaljanauthmoney As Long
Dim totaljanauthnum As Integer
Dim totaljancomm As Integer
Dim totalfebauthmoney As Long
Dim totalfebauthnum As Integer
Dim totalmarauthmoney As Long
Dim totalmarauthnum As Integer
Dim totalaprauthmoney As Long
Dim totalaprauthnum As Integer
Dim totalmayauthmoney As Long
Dim totalmayauthnum As Integer
Dim totaljuneauthmoney As Long
Dim totaljuneauthnum As Integer
Dim totaljulyauthmoney As Long
Dim totaljulyauthnum As Integer
Dim totalaugauthmoney As Long
Dim totalaugauthnum As Integer
Dim totalseptauthmoney As Long
Dim totalseptauthnum As Integer
Dim totaloctauthmoney As Long
Dim totaloctauthnum As Integer
Dim totalnovauthmoney As Long
Dim totalnoveauthnum As Integer
Dim totaldecauthmoney As Long
Dim totaldecauthnum As Integer

Dim totalbncommission As Long
Dim totalarcommission As Long
Dim totalbobdauthcommission As Long
Dim totaledmauthcommission As Long
Dim totalarjrauthcommission As Long


Dim totalbndebenture As Long
Dim totalardebenture As Long
Dim totalbobdauthdebenture As Long
Dim totaledmauthdebenture As Long
Dim totalarjrauthdebenture As Long
Dim totalcadebenture As Long
Dim totalinhousedebenture As Long
Dim totalbnauthorizednum As Integer
Dim totalarauthorizednum As Integer
Dim totaledauthorizednum As Integer
Dim totalarjrauthorizednum As Integer
Dim totalcaauthorizednum As Integer
Dim totalbobdauthorizednum As Integer
Dim totalinhouseauthorizednum As Integer

Dim totaljancommission As Long
Dim totalfebcommission As Long
Dim totalmarcommission As Long
Dim totalaprcommission As Long
Dim totalmaycommission As Long
Dim totaljunecommission As Long
Dim totaljulycommission As Long
Dim totalaugcommission As Long
Dim totalseptcommission As Long
Dim totaloctcommission As Long
Dim totalnovcommission As Long
Dim totaldeccommission As Long

'Gets the dates to be used in the For loop below
Call Datesdeclared2

'these rate are multiplied by the Net Debenture amount
ARRate = 0.006
BNRate = 0.0066
BDRate = 0.005
TierOneRate = (0.0045) * (2 / 3)
TierTwoRate = (0.0065) * (2 / 3)
TierThreeRate = (0.0075) * (2 / 3)
InHouseRate = 0#

rptLR = rptsheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
lastrow = dsheet.Cells(Rows.Count, 1).End(xlUp).Row

On Error Resume Next

y = 4 'starting row

'Checks if dateapproved falls within the start and end date the user inputted and populates the authorized sheet
For x = 2 To lastrow
     If dsheet.Cells(x, 37) >= StartDate And dsheet.Cells(x, 37) <= EndDate Then
        rptsheet.Cells(y, 2) = CDate(dsheet.Cells(x, 37)) 'Approved by SBA Date/dateapproved
        rptsheet.Cells(y, 3) = CDate(dsheet.Cells(x, 34)) 'Board of Directors Resolution Date/boarddate
        rptsheet.Cells(y, 4) = CDate(dsheet.Cells(x, 43)) 'Sent to SBA Date/datesubmitted
        rptsheet.Cells(y, 5) = dsheet.Cells(x, 63) 'BDO
        rptsheet.Cells(y, 6) = dsheet.Cells(x, 66) 'Underwriter/cdcunderwriter
        rptsheet.Cells(y, 7) = dsheet.Cells(x, 69) 'Bank/lendername
        rptsheet.Cells(y, 8) = dsheet.Cells(x, 10) 'knownas
        rptsheet.Cells(y, 9) = dsheet.Cells(x, 211) 'Project #/custom_project
        rptsheet.Cells(y, 10) = CCur(dsheet.Cells(x, 59)) 'Net Debenture/sbanet
        rptsheet.Cells(y, 11) = CCur(dsheet.Cells(x, 60)) 'Gross Debenture/bondamount
        rptsheet.Cells(y, 14) = CCur(dsheet.Cells(x, 63)) 'Commission Paid to/cdcunderwriter
    y = y + 1
    End If
Next x

'dynamically sorts authorized sheet
rptsheet.Select
rptLR2 = rptsheet.Cells(Rows.Count, 3).End(xlUp).Row

    ActiveWorkbook.Worksheets("Authorized").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Authorized").Sort.SortFields.Add Key:=Range("B4:B" & rptLR2), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Authorized").Sort
        .SetRange Range("B3:S" & rptLR2)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
rptsheet.Visible = True
        
'calculates authorization commission for Bob Nance, Armando Ruiz, Ed McGuire, Bob Davis, and A.R. Ruiz Jr.
For y = 4 To rptLR2
    If rptsheet.Cells(y, 5) = "Bob Nance" Then
        rptsheet.Cells(y, 13) = rptsheet.Cells(y, 10) * BNRate
        totalbndebenture = totalbndebenture + rptsheet.Cells(y, 10).Value
        totalbncommission = totalbncommission + rptsheet.Cells(y, 13).Value
        rptsheet.Cells(y, 13) = CCur(rptsheet.Cells(y, 13))
        totalbnauthorizednum = totalbnauthorizednum + 1
    ElseIf rptsheet.Cells(y, 5) = "Armando Ruiz" Then
        rptsheet.Cells(y, 13) = rptsheet.Cells(y, 10) * ARRate
        totalardebenture = totalardebenture + rptsheet.Cells(y, 10).Value
        totalarcommission = totalarcommission + rptsheet.Cells(y, 13).Value
        rptsheet.Cells(y, 13) = CCur(rptsheet.Cells(y, 13))
        totalarauthorizednum = totalarauthorizednum + 1
    ElseIf rptsheet.Cells(y, 5) = "Bob Davis" Then
        rptsheet.Cells(y, 13) = rptsheet.Cells(y, 10) * BDRate
        totalbobdauthdebenture = totalbobdauthdebenture + rptsheet.Cells(y, 10).Value
        totalbobdauthcommission = totalbobdauthcommission + rptsheet.Cells(y, 13).Value
        rptsheet.Cells(y, 13) = CCur(rptsheet.Cells(y, 13))
        totalbobdauthorizednum = totalbobdauthorizednum + 1
    ElseIf rptsheet.Cells(y, 5) = "Chris Allen" Then
        totalcadebenture = totalcadebenture + rptsheet.Cells(y, 10).Value
        rptsheet.Cells(y, 13) = CCur(rptsheet.Cells(y, 13))
        totalcaauthorizednum = totalcaauthorizednum + 1
    ElseIf rptsheet.Cells(y, 5) = "Ed McGuire" Then
        totaledmauthdebenture = totaledmauthdebenture + rptsheet.Cells(y, 10).Value
        If totalbobdauthdebenture <= 4500000 Then
            rptsheet.Cells(y, 13) = rptsheet.Cells(y, 10) * TierOneRate
            totaledmauthcommission = totaledmauthcommission + rptsheet.Cells(y, 13).Value
            rptsheet.Cells(y, 13) = CCur(rptsheet.Cells(y, 13))
            totaledauthorizednum = totaledauthorizednum + 1
        ElseIf totaledmauthdebenture >= 4500001 And totaledmauthdebenture <= 9000000 Then
            rptsheet.Cells(y, 13) = rptsheet.Cells(y, 10) * TierTwoRate
            totaledmauthcommission = totaledmauthcommission + rptsheet.Cells(y, 13).Value
            rptsheet.Cells(y, 13) = CCur(rptsheet.Cells(y, 13))
            totaledauthorizednum = totaledauthorizednum + 1
        ElseIf totaledmauthdebenture >= 9000001 Then
            rptsheet.Cells(y, 13) = rptsheet.Cells(y, 10) * TierThreeRate
            totaledmauthcommission = totaledmauthcommission + rptsheet.Cells(y, 13).Value
            rptsheet.Cells(y, 13) = CCur(rptsheet.Cells(y, 13))
            totaledauthorizednum = totaledauthorizednum + 1
        End If
    ElseIf rptsheet.Cells(y, 5) = "A.R. Ruiz Jr." Then
        totalarjrauthdebenture = totalarjrauthdebenture + rptsheet.Cells(y, 10).Value
        If totalarjrauthdebenture <= 4500000 Then
            rptsheet.Cells(y, 13) = rptsheet.Cells(y, 10) * TierOneRate
            totalarjrauthcommission = totalarjrauthcommission + rptsheet.Cells(y, 13).Value
            rptsheet.Cells(y, 13) = CCur(rptsheet.Cells(y, 13))
            totalarjrauthorizednum = totalarjrauthorizednum + 1
        ElseIf totalarjrauthdebenture >= 4500001 And totalarjrauthdebenture <= 9000000 Then
            rptsheet.Cells(y, 13) = rptsheet.Cells(y, 10) * TierTwoRate
            totalarjrauthcommission = totalarjrauthcommission + rptsheet.Cells(y, 13).Value
            rptsheet.Cells(y, 13) = CCur(rptsheet.Cells(y, 13))
            totalarjrauthorizednum = totalarjrauthorizednum + 1
        ElseIf totalarjrauthdebenture >= 9000001 Then
            rptsheet.Cells(y, 13) = rptsheet.Cells(y, 10) * TierThreeRate
            totalarjrauthcommission = totalarjrauthcommission + rptsheet.Cells(y, 13).Value
            rptsheet.Cells(y, 13) = CCur(rptsheet.Cells(y, 13))
            totalarjrauthorizednum = totalarjrauthorizednum + 1
        End If
    Else 'rptsheet.Cells(y, 2) = "Corey Gaskill" Or rptsheet.Cells(x, 3) = "Oscar Martinez" Or rptsheet.Cells(x, 3) = "Suzanna Caballero" Or rptsheet.Cells(x, 3) = "Blyth Rehberg" Then
        rptsheet.Cells(y, 13) = 0
        totalinhousedebenture = totalinhousedebenture + rptsheet.Cells(y, 10).Value
        totalinhouseauthorizednum = totalinhouseauthorizednum + 1
    End If
Next y


'calculates the running total for commissions and net debenture and provides a breakdown by month
For y = 4 To rptLR2
   If CDate(rptsheet.Cells(y, 2)) >= BegOct And CDate(rptsheet.Cells(y, 2)) <= EndOct Then
        totaloctauthmoney = totaloctauthmoney + rptsheet.Cells(y, 10).Value
        totaloctcommission = totaloctcommission + rptsheet.Cells(y, 13).Value
        totaloctauthnum = totaloctauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegNov And CDate(rptsheet.Cells(y, 2)) <= EndNov Then
        totalnovauthmoney = totalnovauthmoney + rptsheet.Cells(y, 10).Value
        totalnovcommission = totalnovcommission + rptsheet.Cells(y, 13).Value
        totalnovauthnum = totalnovauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegDec And CDate(rptsheet.Cells(y, 2)) <= EndDec Then
        totaldecauthmoney = totaldecauthmoney + rptsheet.Cells(y, 10).Value
        totaldeccommission = totaldeccommission + rptsheet.Cells(y, 13).Value
        totaldecauthnum = totaldecauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegJan And CDate(rptsheet.Cells(y, 2)) <= EndJan Then
        totaljanauthmoney = totaljanauthmoney + rptsheet.Cells(y, 10).Value
        totaljancommission = totaljancommission + rptsheet.Cells(y, 13).Value
        totaljanauthnum = totaljanauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegFeb And CDate(rptsheet.Cells(y, 2)) <= EndFeb Then
        totalfebauthmoney = totalfebauthmoney + rptsheet.Cells(y, 10).Value
        totalfebcommission = totalfebcommission + rptsheet.Cells(y, 13).Value
        totalfebauthnum = totalfebauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegMar And CDate(rptsheet.Cells(y, 2)) <= EndMar Then
        totalmarauthmoney = totalmarauthmoney + rptsheet.Cells(y, 10).Value
        totalmarcommission = totalmarcommission + rptsheet.Cells(y, 13).Value
        totalmarauthnum = totalmarauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegApr And CDate(rptsheet.Cells(y, 2)) <= EndApr Then
        totalaprauthmoney = totalaprauthmoney + rptsheet.Cells(y, 10).Value
        totalaprcommission = totalaprcommission + rptsheet.Cells(y, 13).Value
        totalaprauthnum = totalaprauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegMay And CDate(rptsheet.Cells(y, 2)) <= EndMay Then
        totalmayauthmoney = totalmayauthmoney + rptsheet.Cells(y, 10).Value
        totalmaycommission = totalmaycommission + rptsheet.Cells(y, 13).Value
        totalmayauthnum = totalmayauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegJune And CDate(rptsheet.Cells(y, 2)) <= EndJune Then
        totaljuneauthmoney = totaljuneauthmoney + rptsheet.Cells(y, 10).Value
        totaljunecommission = totaljunecommission + rptsheet.Cells(y, 13).Value
        totaljuneauthnum = totaljuneauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegJuly And CDate(rptsheet.Cells(y, 2)) <= EndJuly Then
        totaljulyauthmoney = totaljulyauthmoney + rptsheet.Cells(y, 10).Value
        totaljulycommission = totaljulycommission + rptsheet.Cells(y, 13).Value
        totaljulyauthnum = totaljulyauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegAug And CDate(rptsheet.Cells(y, 2)) <= EndAug Then
        totalaugauthmoney = totalaugauthmoney + rptsheet.Cells(y, 10).Value
        totalaugcommission = totalaugcommission + rptsheet.Cells(y, 13).Value
        totalaugauthnum = totalaugauthnum + 1
    ElseIf CDate(rptsheet.Cells(y, 2)) >= BegSept And CDate(rptsheet.Cells(y, 2)) <= EndSept Then
        totalseptauthmoney = totalseptauthmoney + rptsheet.Cells(y, 10).Value
        totalseptcommission = totalseptcommission + rptsheet.Cells(y, 13).Value
        totalseptauthnum = totalseptauthnum + 1
    End If
Next y


'Pastes results to Dashboard Screen
dashsum.Range("c28") = totaloctauthmoney
dashsum.Range("c29") = totalnovauthmoney
dashsum.Range("c30") = totaldecauthmoney
dashsum.Range("c31") = totaljanauthmoney
dashsum.Range("c32") = totalfebauthmoney
dashsum.Range("c33") = totalmarauthmoney
dashsum.Range("c34") = totalaprauthmoney
dashsum.Range("c35") = totalmayauthmoney
dashsum.Range("c36") = totaljuneauthmoney
dashsum.Range("c37") = totaljulyauthmoney
dashsum.Range("c38") = totalaugauthmoney
dashsum.Range("c39") = totalseptauthmoney


dashsum.Range("d28") = totaloctauthnum
dashsum.Range("d29") = totalnovauthnum
dashsum.Range("d30") = totaldecauthnum
dashsum.Range("d31") = totaljanauthnum
dashsum.Range("d32") = totalfebauthnum
dashsum.Range("d33") = totalmarauthnum
dashsum.Range("d34") = totalaprauthnum
dashsum.Range("d35") = totalmayauthnum
dashsum.Range("d36") = totaljuneauthnum
dashsum.Range("d37") = totaljulyauthnum
dashsum.Range("d38") = totalaugauthnum
dashsum.Range("d39") = totalseptauthnum

dashsum.Range("F28") = totaloctcommission
dashsum.Range("F29") = totalnovcommission
dashsum.Range("F30") = totaldeccommission
dashsum.Range("F31") = totaljancommission
dashsum.Range("F32") = totalfebcommission
dashsum.Range("F33") = totalmarcommission
dashsum.Range("F34") = totalaprcommission
dashsum.Range("F35") = totalmaycommission
dashsum.Range("F36") = totaljunecommission
dashsum.Range("F37") = totaljulycommission
dashsum.Range("F38") = totalaugcommission
dashsum.Range("F39") = totalseptcommission

'pastes results to bottom of Authorized sheet
rptsheet.Range("g91") = totalbndebenture
rptsheet.Range("g92") = totalardebenture
rptsheet.Range("g93") = totalbobdauthdebenture
rptsheet.Range("g94") = totaledmauthdebenture
rptsheet.Range("g95") = totalarjrauthdebenture
rptsheet.Range("g96") = totalcadebenture
rptsheet.Range("g97") = totalinhousedebenture

rptsheet.Range("h91") = totalbncommission
rptsheet.Range("h92") = totalarcommission
rptsheet.Range("h93") = totalbobdauthcommission
rptsheet.Range("h94") = totaledmauthcommission
rptsheet.Range("h95") = totalarjrauthcommission

rptsheet.Range("i91") = totalbnauthorizednum
rptsheet.Range("i92") = totalarauthorizednum
rptsheet.Range("i93") = totalbobdauthorizednum
rptsheet.Range("i94") = totaledauthorizednum
rptsheet.Range("i95") = totalarjrauthorizednum
rptsheet.Range("i96") = totalcaauthorizednum
rptsheet.Range("i97") = totalinhouseauthorizednum


'configuration for printing
    rptsheet.Activate
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "=OFFSET(Authorized!$B$1,2,0,COUNTA(Authorized!$B:$B),13),$E$90:$J$98"
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0)
        .CenterHorizontally = False
        .CenterVertically = False
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintNoComments
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .PrintTitleRows = "$1:$1"
    End With
    
'adjust sheet via AutoFit
rptsheet.Select
Cells.Select
Selection.Columns.AutoFit
Columns("M:M").ColumnWidth = 11
End Sub
Sub FundedSheetData()
Dim dsheet As Worksheet

Set dsheet = Sheets("data")
Set fndsheet = Sheets("Funded")
Set dashsum = Sheets("Dashboard")

Dim totalgrossoctfndmoney As Long
Dim totalgrossnovfndmoney As Long
Dim totalgrossdecfndmoney As Long
Dim totalgrossjanfndmoney As Long
Dim totalgrossfebfndmoney As Long
Dim totalgrossmarfndmoney As Long
Dim totalgrossaprfndmoney As Long
Dim totalgrossmayfndmoney As Long
Dim totalgrossjunefndmoney As Long
Dim totalgrossjulyfndmoney As Long
Dim totalgrossaugfndmoney As Long
Dim totalgrossseptfndmoney As Long

Dim totalnetoctfndmoney As Long
Dim totalnetnovfndmoney As Long
Dim totalnetdecfndmoney As Long
Dim totalnetjanfndmoney As Long
Dim totalnetfebfndmoney As Long
Dim totalnetmarfndmoney As Long
Dim totalnetaprfndmoney As Long
Dim totalnetmayfndmoney As Long
Dim totalnetjunefndmoney As Long
Dim totalnetjulyfndmoney As Long
Dim totalnetaugfndmoney As Long
Dim totalnetseptfndmoney As Long

Dim totalfndfeeoct As Long
Dim totalfndfeenov As Long
Dim totalfndfeedec As Long
Dim totalfndfeejan As Long
Dim totalfndfeefeb As Long
Dim totalfndfeemar As Long
Dim totalfndfeeapr As Long
Dim totalfndfeemay As Long
Dim totalfndfeejune As Long
Dim totalfndfeejuly As Long
Dim totalfndfeeaug As Long
Dim totalfndfeesept As Long

Dim totalfndsalescommoct As Long
Dim totalfndsalescommnov As Long
Dim totalfndsalescommdec As Long
Dim totalfndsalescommjan As Long
Dim totalfndsalescommfeb As Long
Dim totalfndsalescommmar As Long
Dim totalfndsalescommapr As Long
Dim totalfndsalescommmay As Long
Dim totalfndsalescommjune As Long
Dim totalfndsalescommjuly As Long
Dim totalfndsalescommaug As Long
Dim totalfndsalescommsept As Long

Dim totalfndnumoct As Integer
Dim totalfndnumnov As Integer
Dim totalfndnumdec As Integer
Dim totalfndnumjan As Integer
Dim totalfndnumfeb As Integer
Dim totalfndnummar As Integer
Dim totalfndnumapr As Integer
Dim totalfndnummay As Integer
Dim totalfndnumjune As Integer
Dim totalfndnumjuly As Integer
Dim totalfndnumaug As Integer
Dim totalfndnumsept As Integer

Dim totalbnfncommission As Long
Dim totalarfncommission As Long
Dim totalbdfncommission As Long
Dim totaledmcommission As Long
Dim totalarjrcommission As Long
Dim totalbnfndebenture As Long
Dim totalarfndebenture As Long
Dim totalbdfndebenture As Long
Dim totaledmfndebenture As Long
Dim totalarjrfndebenture As Long
Dim totalcafndebenture As Long
Dim totalinhousefndebenture As Long
Dim totalbnfundednum As Integer
Dim totalarfundednum As Integer
Dim totalbdfundednum As Integer
Dim totaledmfundednum As Integer
Dim totalarjrfundednum As Integer
Dim totalcafundednum As Integer
Dim totalinhousefundednum As Integer

'the rate is multiplied to the Net Debenture
fundfee = 0.005

'the rate is multiplied to the Funding Fee Earned
bncsalescomm = 0.66
arsalescomm = 0.6
bdavissalescomm = 0.1

'rates for Ed McGuire, A.R. Ruiz Jr.
TierOneRate = 0.0045 * (1 / 3)
TierTwoRate = 0.0065 * (1 / 3)
TierThreeRate = 0.0075 * (1 / 3)
InHouseRate = 0#

Call Datesdeclared2

fndsheetLR = fndsheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
dsheet.Activate
lastrow = dsheet.Cells(Rows.Count, 1).End(xlUp).Row
'On Error Resume Next


y = 3 'starting row

'Checks if dateapproved falls within the start and end date the user inputted and populates the funded sheet
For x = 2 To lastrow
     If dsheet.Cells(x, 39) >= StartDate And dsheet.Cells(x, 39) <= EndDate Then
        fndsheet.Cells(y, 1) = CDate(dsheet.Cells(x, 39)) 'Date Funded/datefunded
        fndsheet.Cells(y, 2) = dsheet.Cells(x, 10) 'SBA Loan Name/knownas
        fndsheet.Cells(y, 3) = dsheet.Cells(x, 6) 'SBA Loan #/loannumber
        fndsheet.Cells(y, 4) = dsheet.Cells(x, 70) 'SBA District Office/sbaoffice
        fndsheet.Cells(y, 5) = dsheet.Cells(x, 67) 'Closing Attorney/cdcattorney
        fndsheet.Cells(y, 6) = dsheet.Cells(x, 69) 'Bank/lendername
        fndsheet.Cells(y, 7) = CCur(dsheet.Cells(x, 60)) 'Gross Debenture/bondamount
        fndsheet.Cells(y, 9) = CCur(dsheet.Cells(x, 59)) 'Net Debenture/SBAnet
        fndsheet.Cells(y, 10) = fndsheet.Cells(y, 9).Value * fundfee 'calculates funding fee
        fndsheet.Cells(y, 10) = CCur(fndsheet.Cells(y, 10)) 'converts funding fee to currency
        fndsheet.Cells(y, 12) = dsheet.Cells(x, 63) 'Officer/cdcofficername
    y = y + 1
    End If
Next x

'Dynamically sorts the data by Date Funded
fndsheet.Select
fndLR = fndsheet.Cells(Rows.Count, 2).End(xlUp).Row

    ActiveWorkbook.Worksheets("Funded").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Funded").Sort.SortFields.Add Key:=Range("A3:A" & fndLR), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Funded").Sort
        .SetRange Range("A2:L" & fndLR)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
fndsheet.Visible = True
        
'calculates sales comission based on funded fee earned for BDOs and tallies up funded commission,
'number of funded loans funded net debenture amount
For y = 3 To fndLR
    If fndsheet.Cells(y, 12) = "Bob Nance" Then
        fndsheet.Cells(y, 11) = fndsheet.Cells(y, 10).Value * bncsalescomm
        fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
        totalbnfncommission = totalbnfncommission + fndsheet.Cells(y, 11).Value
        totalbnfndebenture = totalbnfndebenture + fndsheet.Cells(y, 9).Value
        totalbnfundednum = totalbnfundednum + 1
    ElseIf fndsheet.Cells(y, 12) = "Armando Ruiz" Then
        fndsheet.Cells(y, 11) = fndsheet.Cells(y, 10).Value * arsalescomm
        fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
        totalarfncommission = totalarfncommission + fndsheet.Cells(y, 11).Value
        totalarfndebenture = totalarfndebenture + fndsheet.Cells(y, 9).Value
        totalarfundednum = totalarfundednum + 1
    ElseIf fndsheet.Cells(y, 12) = "Bob Davis" Then
        fndsheet.Cells(y, 11) = fndsheet.Cells(y, 10).Value * bdavissalescomm
        fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
        totalbdfncommission = totalbdfncommission + fndsheet.Cells(y, 11).Value
        totalbdfndebenture = totalbdfndebenture + fndsheet.Cells(y, 9).Value
        totalbdfundednum = totalbdfundednum + 1
    ElseIf fndsheet.Cells(y, 12) = "Chris Allen" Then
        fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
        totalcafndebenture = totalcafndebenture + fndsheet.Cells(y, 9).Value
        totalcafundednum = totalcafundednum + 1
    ElseIf fndsheet.Cells(y, 12) = "Ed McGuire" Then
        totaledmfndebenture = totaledmfndebenture + fndsheet.Cells(y, 9).Value
        If totaledmfndebenture <= 4500000 Then
            fndsheet.Cells(y, 11) = fndsheet.Cells(y, 9).Value * TierOneRate
            fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
            totaledmcommission = totaledmcommission + fndsheet.Cells(y, 11).Value
            totaledmfundednum = totaledmfundednum + 1
        ElseIf totaledmfndebenture >= 4500001 And totaledmfndebenture <= 9000000 Then
            fndsheet.Cells(y, 11) = fndsheet.Cells(y, 9).Value * TierTwoRate
            fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
            totaledmcommission = totaledmcommission + fndsheet.Cells(y, 11).Value
            totaledmfundednum = totaledmfundednum + 1
        ElseIf totaledmfndebenture >= 9000001 Then
            fndsheet.Cells(y, 11) = fndsheet.Cells(y, 9).Value * TierThreeRate
            fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
            totaledmcommission = totaledmcommission + fndsheet.Cells(y, 11).Value
            totaledmfundednum = totaledmfundednum + 1
        End If
    ElseIf fndsheet.Cells(y, 12) = "A.R. Ruiz Jr." Then
        totalarjrfndebenture = totalarjrfndebenture + fndsheet.Cells(y, 9).Value
        If totalarjrfndebenture <= 4500000 Then
            fndsheet.Cells(y, 11) = fndsheet.Cells(y, 9).Value * TierOneRate
            fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
            totalarjrcommission = totalarjrcommission + fndsheet.Cells(y, 11).Value
            totalarjrfundednum = totalarjrfundednum + 1
        ElseIf totalarjrfndebenture >= 4500001 And totalarjrfndebenture <= 9000000 Then
            fndsheet.Cells(y, 11) = fndsheet.Cells(y, 9).Value * TierTwoRate
            fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
            totalarjrcommission = totalarjrcommission + fndsheet.Cells(y, 11).Value
            totalarjrfundednum = totalarjrfundednum + 1
        ElseIf totalarjrfndebenture >= 9000001 Then
            fndsheet.Cells(y, 11) = fndsheet.Cells(y, 9).Value * TierThreeRate
            fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
            totalarjrcommission = totalarjrcommission + fndsheet.Cells(y, 11).Value
            totalarjrfundednum = totalarjrfundednum + 1
        End If
    ElseIf fndsheet.Cells(y, 12) = "Corey Gaskill" Or fndsheet.Cells(y, 12) = "Oscar Martinez" Or fndsheet.Cells(y, 12) = "Suzanna Caballero" Or fndsheet.Cells(y, 12) = "Blyth Rehberg" Then
        fndsheet.Cells(y, 11) = 0
        fndsheet.Cells(y, 11) = fndsheet.Cells(y, 10).Value * InHouseRate
        fndsheet.Cells(y, 11) = CCur(fndsheet.Cells(y, 11))
        totalinhousefndebenture = totalinhousefndebenture + fndsheet.Cells(y, 9).Value
        totalinhousefundednum = totalinhousefundednum + 1
    End If
Next y
        
'calculates the running total for commissions, net debenture, funding fee and provides a breakdown by month
For y = 3 To fndLR
    If CDate(fndsheet.Cells(y, 1)) >= BegOct And CDate(fndsheet.Cells(y, 1)) <= EndOct Then
        totalgrossoctfndmoney = totalgrossoctfndmoney + fndsheet.Cells(y, 7).Value
        totalnetoctfndmoney = totalnetoctfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeeoct = totalfndfeeoct + fndsheet.Cells(y, 10).Value
        totalfndsalescommoct = totalfndsalescommoct + fndsheet.Cells(y, 11).Value
        totalfndnumoct = totalfndnumoct + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegNov And CDate(fndsheet.Cells(y, 1)) <= EndNov Then
        totalgrossnovfndmoney = totalgrossnovfndmoney + fndsheet.Cells(y, 7).Value
        totalnetnovfndmoney = totalnetnovfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeenov = totalfndfeenov + fndsheet.Cells(y, 10).Value
        totalfndsalescommnov = totalfndsalescommnov + fndsheet.Cells(y, 11).Value
        totalfndnumnov = totalfndnumnov + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegDec And CDate(fndsheet.Cells(y, 1)) <= EndDec Then
        totalgrossdecfndmoney = totalgrossdecfndmoney + fndsheet.Cells(y, 7).Value
        totalnetdecfndmoney = totalnetdecfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeedec = totalfndfeedec + fndsheet.Cells(y, 10).Value
        totalfndsalescommdec = totalfndsalescommdec + fndsheet.Cells(y, 11).Value
        totalfndnumdec = totalfndnumdec + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegJan And CDate(fndsheet.Cells(y, 1)) <= EndJan Then
        totalgrossjanfndmoney = totalgrossjanfndmoney + fndsheet.Cells(y, 7).Value
        totalnetjanfndmoney = totalnetjanfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeejan = totalfndfeejan + fndsheet.Cells(y, 10).Value
        totalfndsalescommjan = totalfndsalescommjan + fndsheet.Cells(y, 11).Value
        totalfndnumjan = totalfndnumjan + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegFeb And CDate(fndsheet.Cells(y, 1)) <= EndFeb Then
        totalgrossfebfndmoney = totalgrossfebfndmoney + fndsheet.Cells(y, 7).Value
        totalnetfebfndmoney = totalnetfebfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeefeb = totalfndfeefeb + fndsheet.Cells(y, 10).Value
        totalfndsalescommfeb = totalfndsalescommfeb + fndsheet.Cells(y, 11).Value
        totalfndnumfeb = totalfndnumfeb + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegMar And CDate(fndsheet.Cells(y, 1)) <= EndMar Then
        totalgrossmarfndmoney = totalgrossmarfndmoney + fndsheet.Cells(y, 7).Value
        totalnetmarfndmoney = totalnetmarfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeemar = totalfndfeemar + fndsheet.Cells(y, 10).Value
        totalfndsalescommmar = totalfndsalescommmar + fndsheet.Cells(y, 11).Value
        totalfndnummar = totalfndnummar + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegApr And CDate(fndsheet.Cells(y, 1)) <= EndApr Then
        totalgrossaprfndmoney = totalgrossaprfndmoney + fndsheet.Cells(y, 7).Value
        totalnetaprfndmoney = totalnetaprfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeeapr = totalfndfeeapr + fndsheet.Cells(y, 10).Value
        totalfndsalescommapr = totalfndsalescommapr + fndsheet.Cells(y, 11).Value
        totalfndnumapr = totalfndnumapr + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegMay And CDate(fndsheet.Cells(y, 1)) <= EndMay Then
        totalgrossmayfndmoney = totalgrossmayfndmoney + fndsheet.Cells(y, 7).Value
        totalnetmayfndmoney = totalnetmayfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeemay = totalfndfeemay + fndsheet.Cells(y, 10).Value
        totalfndsalescommmay = totalfndsalescommmay + fndsheet.Cells(y, 11).Value
        totalfndnummay = totalfndnummay + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegJune And CDate(fndsheet.Cells(y, 1)) <= EndJune Then
        totalgrossjunefndmoney = totalgrossjunefndmoney + fndsheet.Cells(y, 7).Value
        totalnetjunefndmoney = totalnetjunefndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeejune = totalfndfeejune + fndsheet.Cells(y, 10).Value
        totalfndsalescommjune = totalfndsalescommjune + fndsheet.Cells(y, 11).Value
        totalfndnumjune = totalfndnumjune + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegJuly And CDate(fndsheet.Cells(y, 1)) <= EndJuly Then
        totalgrossjulyfndmoney = totalgrossjulyfndmoney + fndsheet.Cells(y, 7).Value
        totalnetjulyfndmoney = totalnetjulyfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeejuly = totalfndfeejuly + fndsheet.Cells(y, 10).Value
        totalfndsalescommjuly = totalfndsalescommjuly + fndsheet.Cells(y, 11).Value
        totalfndnumjuly = totalfndnumjuly + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegAug And CDate(fndsheet.Cells(y, 1)) <= EndAug Then
        totalgrossaugfndmoney = totalgrossaugfndmoney + fndsheet.Cells(y, 7).Value
        totalnetaugfndmoney = totalnetaugfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeeaug = totalfndfeeaug + fndsheet.Cells(y, 10).Value
        totalfndsalescommaug = totalfndsalescommaug + fndsheet.Cells(y, 11).Value
        totalfndnumaug = totalfndnumaug + 1
    ElseIf CDate(fndsheet.Cells(y, 1)) >= BegSept And CDate(fndsheet.Cells(y, 1)) <= EndSept Then
        totalgrossseptfndmoney = totalgrossseptfndmoney + fndsheet.Cells(y, 7).Value
        totalnetseptfndmoney = totalnetseptfndmoney + fndsheet.Cells(y, 9).Value
        totalfndfeesept = totalfndfeesept + fndsheet.Cells(y, 10).Value
        totalfndsalescommsept = totalfndsalescommsept + fndsheet.Cells(y, 11).Value
        totalfndnumsept = totalfndnumsept + 1
    End If
Next y

'Pastes results to Dashboard Sheet
dashsum.Range("c9") = totalgrossoctfndmoney
dashsum.Range("c10") = totalgrossnovfndmoney
dashsum.Range("c11") = totalgrossdecfndmoney
dashsum.Range("c12") = totalgrossjanfndmoney
dashsum.Range("c13") = totalgrossfebfndmoney
dashsum.Range("c14") = totalgrossmarfndmoney
dashsum.Range("c15") = totalgrossaprfndmoney
dashsum.Range("c16") = totalgrossmayfndmoney
dashsum.Range("c17") = totalgrossjunefndmoney
dashsum.Range("c18") = totalgrossjulyfndmoney
dashsum.Range("c19") = totalgrossaugfndmoney
dashsum.Range("c20") = totalgrossseptfndmoney

dashsum.Range("d9") = totalfndnumoct
dashsum.Range("d10") = totalfndnumnov
dashsum.Range("d11") = totalfndnumdec
dashsum.Range("d12") = totalfndnumjan
dashsum.Range("d13") = totalfndnumfeb
dashsum.Range("d14") = totalfndnummar
dashsum.Range("d15") = totalfndnumapr
dashsum.Range("d16") = totalfndnummay
dashsum.Range("d17") = totalfndnumjune
dashsum.Range("d18") = totalfndnumjuly
dashsum.Range("d19") = totalfndnumaug
dashsum.Range("d20") = totalfndnumsept

dashsum.Range("n9") = totalfndfeeoct
dashsum.Range("n10") = totalfndfeenov
dashsum.Range("n11") = totalfndfeedec
dashsum.Range("n12") = totalfndfeejan
dashsum.Range("n13") = totalfndfeefeb
dashsum.Range("n14") = totalfndfeemar
dashsum.Range("n15") = totalfndfeeapr
dashsum.Range("n16") = totalfndfeemay
dashsum.Range("n17") = totalfndfeejune
dashsum.Range("n18") = totalfndfeejuly
dashsum.Range("n19") = totalfndfeeaug
dashsum.Range("n20") = totalfndfeesept

dashsum.Range("o9") = totalfndsalescommoct
dashsum.Range("o10") = totalfndsalescommnov
dashsum.Range("o11") = totalfndsalescommdec
dashsum.Range("o12") = totalfndsalescommjan
dashsum.Range("o13") = totalfndsalescommfeb
dashsum.Range("o14") = totalfndsalescommmar
dashsum.Range("o15") = totalfndsalescommapr
dashsum.Range("o16") = totalfndsalescommmay
dashsum.Range("o17") = totalfndsalescommjune
dashsum.Range("o18") = totalfndsalescommjuly
dashsum.Range("o19") = totalfndsalescommaug
dashsum.Range("o20") = totalfndsalescommsept

'Pastes results to bottom of Funded Sheet
fndsheet.Range("g91") = totalbnfndebenture
fndsheet.Range("g92") = totalarfndebenture
fndsheet.Range("g93") = totalbdfndebenture
fndsheet.Range("g94") = totaledmfndebenture
fndsheet.Range("g95") = totalarjrfndebenture
fndsheet.Range("g96") = totalcafndebenture
fndsheet.Range("g97") = totalinhousefndebenture
fndsheet.Range("h91") = totalbnfncommission
fndsheet.Range("h92") = totalarfncommission
fndsheet.Range("h93") = totalbdfncommission
fndsheet.Range("h94") = totaledmcommission
fndsheet.Range("h95") = totalarjrcommission
fndsheet.Range("i91") = totalbnfundednum
fndsheet.Range("i92") = totalarfundednum
fndsheet.Range("i93") = totalbdfundednum
fndsheet.Range("i94") = totaledmfundednum
fndsheet.Range("i95") = totalarjrfundednum
fndsheet.Range("i96") = totalcafundednum
fndsheet.Range("i97") = totalinhousefundednum

'configuration for printing
    fndsheet.Activate
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "=OFFSET(Funded!$A$1,1,0,COUNTA(Funded!$A:$A),12),$E$90:$J$98"
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0)
        .CenterHorizontally = False
        .CenterVertically = False
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintNoComments
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .PrintTitleRows = "$1:$1"
    End With

fndsheet.Select
Cells.Select
Selection.Columns.AutoFit

End Sub
Sub CountPaidOff()

Dim dsheet As Worksheet

Set dashsum = Sheets("Dashboard")
Set dsheet = Sheets("data")
Set pdsheet = Sheets("Paid Off")

Dim totaljanpd As Integer
Dim totalfebpd As Integer
Dim totalmarpd As Integer
Dim totalaprpd As Integer
Dim totalmaypd As Integer
Dim totaljunepd As Integer
Dim totaljulypd As Integer
Dim totalaugpd As Integer
Dim totalseptpd As Integer
Dim totaloctpd As Integer
Dim totalnovpd As Integer
Dim totaldecpd As Integer

Call Datesdeclared2

y = 3 'starting row

'Locates where the last row falls
pdsheetLR = pdsheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
lastrow = dsheet.Cells(Rows.Count, 1).End(xlUp).Row

'Checks if paidoff date falls within the start and end date the user inputted and populates the Paid Off sheet
For x = 2 To lastrow
    If dsheet.Cells(x, 41) >= StartDate And dsheet.Cells(x, 41) <= EndDate And dsheet.Cells(x, 31) = "Paid Off" Then
        pdsheet.Cells(y, 1) = CDate(dsheet.Cells(x, 41)) 'Date Paid Off/datepaid
        pdsheet.Cells(y, 2) = dsheet.Cells(x, 10) 'Known As/Knownas
        pdsheet.Cells(y, 3) = dsheet.Cells(x, 6) 'SBA Loan #/loannumber
        pdsheet.Cells(y, 4) = dsheet.Cells(x, 69) 'Bank/lendername
        pdsheet.Cells(y, 5) = CCur(dsheet.Cells(x, 60)) 'Gross Debenture/bondamount
        pdsheet.Cells(y, 6) = CCur(dsheet.Cells(x, 59)) 'Net Debenture/SBAnet
        pdsheet.Cells(y, 7) = dsheet.Cells(x, 63) 'Underwriter/cdcunderwriter
          
'Provides a breakdown by month
        If CDate(dsheet.Cells(x, 41)) >= BegOct And CDate(dsheet.Cells(x, 41)) <= EndOct Then
            totaloctpd = totaloctpd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegNov And CDate(dsheet.Cells(x, 41)) <= EndNov Then
            totalnovpd = totalnovpd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegDec And CDate(dsheet.Cells(x, 41)) <= EndDec Then
            totaldecpd = totaldecpd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegJan And CDate(dsheet.Cells(x, 41)) <= EndJan Then
            totaljanpd = totaljanpd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegFeb And CDate(dsheet.Cells(x, 41)) <= EndFeb Then
            totalfebpd = totalfebpd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegMar And CDate(dsheet.Cells(x, 41)) <= EndMar Then
            totalmarpd = totalmarpd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegApr And CDate(dsheet.Cells(x, 41)) <= EndApr Then
            totalaprpd = totalaprpd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegMay And CDate(dsheet.Cells(x, 41)) <= EndMay Then
            totalmaypd = totalmaypd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegJune And CDate(dsheet.Cells(x, 41)) <= EndJune Then
            totaljunepd = totaljunepd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegJuly And CDate(dsheet.Cells(x, 41)) <= EndJuly Then
            totaljulypd = totaljulypd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegAug And CDate(dsheet.Cells(x, 41)) <= EndAug Then
            totalaugpd = totalaugpd + 1
        ElseIf CDate(dsheet.Cells(x, 41)) >= BegSept And CDate(dsheet.Cells(x, 41)) <= EndSept Then
            totalseptpd = totalseptpd + 1
        End If
     y = y + 1
     End If
     
'Paste the Paid Off results to Dashboard screen
Next x
dashsum.Range("f9") = totaloctpd
dashsum.Range("f10") = totalnovpd
dashsum.Range("f11") = totaldecpd
dashsum.Range("f12") = totaljanpd
dashsum.Range("f13") = totalfebpd
dashsum.Range("f14") = totalmarpd
dashsum.Range("f15") = totalaprpd
dashsum.Range("f16") = totalmaypd
dashsum.Range("f17") = totaljunepd
dashsum.Range("f18") = totaljulypd
dashsum.Range("f19") = totalaugpd
dashsum.Range("f20") = totalseptpd

'configuration for printing
    pdsheet.Activate
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "=OFFSET('Paid Off'!$A$1,1,0,COUNTA('Paid Off'!$A:$A)-1,7)"
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0)
        .CenterHorizontally = False
        .CenterVertically = False
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintNoComments
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .PrintTitleRows = "$1:$1"
    End With

Call DynamicSorting

End Sub
Sub CountLiquidating()

Dim dsheet As Worksheet
Set dashsum = Sheets("Dashboard")
Set ldsheet = Sheets("Liquidated")

Dim totaljanlqnum As Integer
Dim totalfeblqnum As Integer
Dim totalmarlqnum As Integer
Dim totalaprlqnum As Integer
Dim totalmaylqnum As Integer
Dim totaljunelqnum As Integer
Dim totaljulylqnum As Integer
Dim totalauglqnum As Integer
Dim totalseptlqnum As Integer
Dim totaloctlqnum As Integer
Dim totalnovlqnum As Integer
Dim totaldeclqnum As Integer


Call Datesdeclared2

y = 3

Set dsheet = Sheets("data")
lastrow = dsheet.Cells(Rows.Count, 1).End(xlUp).Row

'Checks if Liquidated date falls within the start and end date the user inputted and populates the Liquidated sheet
For x = 2 To lastrow
    If dsheet.Cells(x, 42) >= StartDate And dsheet.Cells(x, 42) <= EndDate And dsheet.Cells(x, 31) = "Liquidated" Then
        ldsheet.Cells(y, 1) = CDate(dsheet.Cells(x, 42)) 'Date Paid Off/datepaid
        ldsheet.Cells(y, 2) = dsheet.Cells(x, 10) 'Known As/Knownas
        ldsheet.Cells(y, 3) = dsheet.Cells(x, 6) 'SBA Loan #/loannumber
        ldsheet.Cells(y, 4) = dsheet.Cells(x, 69) 'Bank/lendername
        ldsheet.Cells(y, 5) = CCur(dsheet.Cells(x, 60)) 'Gross Debenture/bondamount
        ldsheet.Cells(y, 6) = CCur(dsheet.Cells(x, 59)) 'Net Debenture/SBAnet
        ldsheet.Cells(y, 7) = dsheet.Cells(x, 63) 'Underwriter/cdcunderwriter
        
'Provides a breakdown by month
        If CDate(dsheet.Cells(x, 42)) >= BegOct And CDate(dsheet.Cells(x, 42)) <= EndOct Then
            totaloctlqnum = totaloctlqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegNov And CDate(dsheet.Cells(x, 42)) <= EndNov Then
            totalnovlqnum = totalnovlqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegDec And CDate(dsheet.Cells(x, 42)) <= EndDec Then
            totaldeclqnum = totaldeclqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegJan And CDate(dsheet.Cells(x, 42)) <= EndJan Then
            totaljanlqnum = totaljanlqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegFeb And CDate(dsheet.Cells(x, 42)) <= EndFeb Then
            totalfeblqnum = totalfeblqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegMar And CDate(dsheet.Cells(x, 42)) <= EndMar Then
            totalmarlqnum = totalmarlqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegApr And CDate(dsheet.Cells(x, 42)) <= EndApr Then
            totalaprlqnum = totalaprlqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegMay And CDate(dsheet.Cells(x, 42)) <= EndMay Then
            totalmaylqnum = totalmaylqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegJune And CDate(dsheet.Cells(x, 42)) <= EndJune Then
            totaljunelqnum = totaljunelqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegJuly And CDate(dsheet.Cells(x, 42)) <= EndJuly Then
            totaljulylqnum = totaljulylqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegAug And CDate(dsheet.Cells(x, 42)) <= EndAug Then
            totalauglqnum = totalauglqnum + 1
        ElseIf CDate(dsheet.Cells(x, 42)) >= BegSept And CDate(dsheet.Cells(x, 42)) <= EndSept Then
            totalseptlqnum = totalseptlqnum + 1
        End If
     y = y + 1
     End If
Next x

'Paste the Liquidated results to Dashboard screen
dashsum.Range("h9") = totaloctlqnum
dashsum.Range("h10") = totalnovlqnum
dashsum.Range("h11") = totaldeclqnum
dashsum.Range("h12") = totaljanlqnum
dashsum.Range("h13") = totalfeblqnum
dashsum.Range("h14") = totalmarlqnum
dashsum.Range("h15") = totalaprlqnum
dashsum.Range("h16") = totalmaylqnum
dashsum.Range("h17") = totaljunelqnum
dashsum.Range("h18") = totaljulylqnum
dashsum.Range("h19") = totalauglqnum
dashsum.Range("h20") = totalseptlqnum

'configuration for printing
    ldsheet.Activate
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "=OFFSET(Liquidated!$A$1,1,0,COUNTA(Liquidated!$A:$A)-1,7)"
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0)
        .CenterHorizontally = False
        .CenterVertically = False
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintNoComments
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .PrintTitleRows = "$1:$1"
    End With

End Sub
Sub DynamicSorting()
'Sort by date the Paid Off and Liquidated Sheets

Set pdsheet = Sheets("Paid Off")
pdsheet.Select
pdLR = pdsheet.Cells(Rows.Count, 2).End(xlUp).Row

    ActiveWorkbook.Worksheets("Paid Off").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Paid Off").Sort.SortFields.Add Key:=Range("A3:A" & pdLR), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Paid Off").Sort
        .SetRange Range("A2:G" & pdLR)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
pdsheet.Visible = True
pdsheet.Select
Cells.Select
Selection.Columns.AutoFit

Set lqsheet = Sheets("Liquidated")
lqsheet.Select
lqLR = pdsheet.Cells(Rows.Count, 2).End(xlUp).Row

    ActiveWorkbook.Worksheets("Liquidated").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Liquidated").Sort.SortFields.Add Key:=Range("A3:A" & lqLR), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Liquidated").Sort
        .SetRange Range("A2:G" & lqLR)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
lqsheet.Visible = True
lqsheet.Select
Cells.Select
Selection.Columns.AutoFit

End Sub
Sub DashSummary()
'creates Dashboard summary sheet
Set Dashboard = Sheets.Add
Dashboard.Name = "Dashboard"

Call Datesdeclared
Set dashsum = Sheets("Dashboard")
dashsum.Activate
    Range("B1:O1").Select
    With Selection
        .MergeCells = False
        .HorizontalAlignment = xlCenterAcrossSelection
        .Font.Bold = True
    End With
    ActiveCell.FormulaR1C1 = "TXCDC Dashboard"
    
    Range("B5:O21").Select
    With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C5") = "$ New Loans Funded"
    Range("D5") = "# New Loans Funded"
    Range("E5") = "$ Pay-offs"
    Range("F5") = "# Pay-offs"
    Range("G5") = "$ Liquidations"
    Range("H5") = "# Liquidations"
    Range("I5") = "Borrowers' Payments"
    Range("J5") = "Portfolio Balance"
    Range("K5") = "# of Loans"
    Range("L5") = "Servicing Income"
    Range("M5") = "Servicing Commission"
    Range("N5") = "Funding Fees Earned"
    Range("O5") = "Funding Sales Commissions"
    Range("B6") = "Budget"
    Range("B5:O7").Select
    Selection.Font.Bold = True
    Range("B7") = "% of Budget Achieved"
    Range("C7") = "=C21/C6"
    Range("C7").Select
    Selection.Copy
    Range("E7,G7,I7,O7,N7,M7,L7").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("J7") = "=(J6-J8)/J8-(J9-J8)/J8"
    
    Range("B21") = "Total"
    Range("B21:O21").Select
    Selection.Font.Bold = True
    Range("Q5") = "Used for calculation"
    Range("Q9") = "=J8+C9-E9-G9"
    Range("Q10") = "=J9+C10-E10-G10"
    Range("Q9:Q10").Select
    Selection.AutoFill Destination:=Range("Q9:Q20"), Type:=xlFillDefault
    Range("Q9:Q20").Select
    With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("Q5").Select
    Selection.Font.Bold = True
    Range("I20") = "=Q20-J20"
    Range("I19") = "=Q19-J19"
    Range("I20").Activate
    Selection.AutoFill Destination:=Range("I9:I20"), Type:=xlFillDefault
    Range("I9:I20").Select
    Range("C21") = "=SUM(C8:C20)"
    Range("C21").Select
    Selection.Copy
    Range("D21:I21,L21:O21").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'Creates the Authorized summary section
    Range("B25:F42").Select
    With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B24:F24").Select
    ActiveCell.FormulaR1C1 = "AUTHORIZED"
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Calibri"
        .Size = 16
        .ThemeColor = xlThemeColorLight1
    End With
    Range("B24:F24").Select
    Selection.HorizontalAlignment = xlCenterAcrossSelection
    Range("B25:F27,B25:B42").Select
    Selection.Font.Bold = True
    Range("C25") = "$ New Loans Authorized"
    Range("D25") = "# New Loans Authorized"
    Range("E25") = "Fees Earned"
    Range("F25") = "Commission"
    Range("B26") = "Budget"
    Range("B27") = "% of Budget Achieved"
    Range("C27") = "=C40/C26"
    Range("E27") = "=E40/E26"
    Range("F27") = "=F40/F26"
    Range("B42") = "Budgeted Average"
    Range("B41") = "Average Debenture"
    Range("B40") = "Year to Date"
    Range("C40") = "=SUM(C28:C39)"
    Range("C40").Select
    Selection.AutoFill Destination:=Range("C40:F40"), Type:=xlFillDefault
    Range("C40:F40").Select
    Range("C41") = "=C40/D40"
    Range("E28") = "=C28*0.01"
    Range("E28").Select
    Selection.AutoFill Destination:=Range("E28:E39"), Type:=xlFillDefault
    Range("B9:B20").Select
    Selection.Copy
    Range("B28:B39").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


'Creates the Income summary section
    Range("J25:O29").Select
    With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("J24") = "INCOME SUMMARY"
    Range("J24:O24").Select
    Selection.HorizontalAlignment = xlCenterAcrossSelection
    Selection.Font.Bold = True
    Selection.Font.Size = 16
    Range("K25") = "Budget"
    Range("L25") = "YTD"
    Range("M25") = "% of Budget Achieved"
    Range("N25") = "% of Yr."
    Range("O25") = "$ Varaiance"
    Range("J26") = "Net Funding Fees"
    Range("J27") = "Net Authorization Fees"
    Range("J28") = "Net Servicing Fees"
    Range("J29") = "Total"
    Range("K25:O25,J26:J29").Select
    Selection.Font.Bold = True
    Range("K29") = "=SUM(K26:K28)"
    Range("L29") = "=SUM(L26:L28)"
    Range("M29") = "=L29/K29"
    Range("K26") = "=N6-O6"
    Range("K27") = "=E26-F26"
    Range("K28") = "=L6-M6"
    Range("L26") = "=N21-O21"
    Range("L27") = "=E40-F40"
    Range("L28") = "=L21-M21"
    Range("M26") = "=L26/K26"
    Range("M27") = "=L27/K27"
    Range("M28") = "=L28/K28"
    Range("O26") = "=L26-K26"
    Range("O26").Select
    Selection.AutoFill Destination:=Range("O26:O28"), Type:=xlFillDefault
    Range("O29") = "=SUM(O26:O28)"
    Range("M22") = "=M21/L21"
    Range("N22") = "commissioned"
    Range("O22") = "=O21/N21"
    Range("P22") = "commissioned"
    Range("F43") = "=F40/E40"
    Range("G43") = "% commissioned"
    
    Range("J25:O29,B5:O21,B25:F42").Select
    With Selection.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B7:O7,B27:F27").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

'creates the todays date header
    Range("B2:O2").Select
    With Selection
        .MergeCells = False
        .HorizontalAlignment = xlCenterAcrossSelection
        .Font.Bold = True
    End With
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("B2").Select
    Selection.NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Formatting of page
    Range("C7,E7,G7,I7,J7,L7,L7,M7,m22,N7,O7,o22,C27,E27,F27,F43,M26:N29").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("C6,C26,E6,E26,F26,G6,I6,J6,L6,M6,N6,O6,L8:O21,G8:G21,E8:E21,C8:C21,I8:J21,C28:C42,E28:F40,k26:L29,O26:O29").Select
    Selection.Style = "Currency"
    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    Range("B8:B20,B28:B39").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("D8:D21,D6,D26,D28:D40,H6,H8:H21,F6,F8:F21,K6,K8:K21,C5:O5,C25:F25,K25:O25").Select
    Selection.HorizontalAlignment = xlCenter
    
'Adjust Dashboard page for printing and viewing
    Columns("A:A").ColumnWidth = 2.43
    Columns("B:B").ColumnWidth = 19
    Columns("C:C").ColumnWidth = 16.71
    Columns("D:D").ColumnWidth = 11.43
    Columns("E:E").ColumnWidth = 12.71
    Columns("F:F").ColumnWidth = 11
    Columns("G:G").ColumnWidth = 16.71
    Columns("H:H").ColumnWidth = 11.57
    Columns("I:I").ColumnWidth = 13.71
    Columns("J:J").ColumnWidth = 13.71
    Columns("K:K").ColumnWidth = 11.43
    Columns("L:L").ColumnWidth = 12.71
    Columns("M:M").ColumnWidth = 11
    Columns("N:N").ColumnWidth = 13.14
    Columns("O:O").ColumnWidth = 12.57
    Columns("Q:Q").ColumnWidth = 18.57
    Selection.WrapText = True
    
'print configuration
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = "$B$1:$O$43"
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    Application.PrintCommunication = True

End Sub

Sub FundedSheet()
'Creates Funded worksheet

Set Funded = Sheets.Add
Funded.Name = "Funded"
    Range("A1:L1").Select
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    ActiveCell.FormulaR1C1 = "Funded"
    Range("A2") = "Date Funded"
    Range("B2") = "SBA Loan Name"
    Range("C2") = "SBA Loan #"
    Range("D2") = "SBA District Office"
    Range("E2") = "Closing Attorney"
    Range("F2") = "Bank"
    Range("G2") = "Gross Debenture"
    Range("H2") = "Total Gross"
    Range("I2") = "Net Debenture"
    Range("J2") = "Funding Fee Earned"
    Range("K2") = "Commission Due"
    Range("L2") = "Officer"
    Range("A2:L2").Select
    Selection.Font.Underline = xlUnderlineStyleSingle
    Selection.Font.Bold = True
    
    Range("E90") = "OFFICER"
    Range("G90") = "DEBENTURE"
    Range("H90") = "COMMISSION"
    Range("E91") = "Bod Nance's"
    Range("E92") = "Armando Ruiz"
    Range("E93") = "Bob Davis"
    Range("E94") = "Ed McGuire"
    Range("E95") = "A.R. Ruiz Jr."
    Range("E96") = "Chris Allen"
    Range("E97") = "In House (Blyth, Oscar, Corey, Suzanna)"
    Range("E98") = "Total"
    Range("I90") = "NUMBER"
    Range("J90") = "PERCENTAGE"
    Range("G98") = "=SUM(G91:G97)"
    Range("G98").Select
    Selection.AutoFill Destination:=Range("G98:J98"), Type:=xlFillDefault
    Range("J91") = "=G91/$G$98"
    Range("J91").Select
    Selection.AutoFill Destination:=Range("J91:J98"), Type:=xlFillDefault
    Range("J91:J98").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("G91:H98").Select
    Selection.Style = "Currency"
    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    Range("I91:I98").Select
    Selection.HorizontalAlignment = xlCenter
    Range("E90:J90").Select
    Selection.Font.Bold = True

End Sub

Sub AuthorizedSheet()
'Creates Authorized Sheet
Set Authorized = Sheets.Add
Authorized.Name = "Authorized"

    Range("A1:S1").Select
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .Font.Bold = True
    End With
    ActiveCell.FormulaR1C1 = "Authorized"
    Range("B3:S3").Select
    Selection.Font.Underline = xlUnderlineStyleSingle
    Selection.Font.Bold = True
    Sheets("Authorized").Select
    Range("B3") = "Approved by SBA"
    Range("C3") = "Directors Resolution Date"
    Range("D3") = "Sent to SBA"
    Range("E3") = "BDO"
    Range("F3") = "Underwriter"
    Range("G3") = "Bank"
    Range("H3") = "Known as"
    Range("I3") = "Project #"
    Range("J3") = "Net Debenture"
    Range("K3") = "Gross Debenture"
    Range("L3") = "Authorization Fee Earned"
    Range("M3") = "Commission"
    Range("N3") = "Commission Paid To"
    Range("O3") = "Original"
    Range("P3") = "Action"
    Range("Q3") = "Collected From Borrower"
    Range("R3") = "WIP"
    Range("S3") = "Commission Payable"
    Range("B3:T3,E90:J90").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("E90") = "OFFICER"
    Range("G90") = "DEBENTURE"
    Range("H90") = "COMMISSION"
    Range("E91") = "Bod Nance's"
    Range("E92") = "Armando Ruiz"
    Range("E93") = "Bob Davis"
    Range("E94") = "Ed McGuire"
    Range("E95") = "A.R. Ruiz Jr."
    Range("E96") = "Chris Allen"
    Range("E97") = "In House (Blyth, Oscar, Corey, Suzanna)"
    Range("E98") = "Total"
    Range("I90") = "NUMBER"
    Range("J90") = "PERCENTAGE"
    Range("G98") = "=SUM(G91:G97)"
    Range("G98").Select
    Selection.AutoFill Destination:=Range("G98:J98"), Type:=xlFillDefault
    Range("J91") = "=G91/$G$98"
    Range("J91").Select
    Selection.AutoFill Destination:=Range("J91:J97"), Type:=xlFillDefault
    Range("J91:J98").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("G91:H98").Select
    Selection.Style = "Currency"
    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    Range("I91:I98").Select
    Selection.HorizontalAlignment = xlCenter

End Sub
 
Sub PaidoffSheet()
'Creates Paid Off Sheet
Set Paidoff = Sheets.Add
Paidoff.Name = "Paid Off"

Range("A1:G1").Select
    With Selection
        .MergeCells = False
        .HorizontalAlignment = xlCenterAcrossSelection
        .Font.Bold = True
    End With
    ActiveCell.FormulaR1C1 = "Paid Off"

Range("A2") = "Date Paid Off"
Range("B2") = "Known As"
Range("C2") = "SBA Loan #"
Range("D2") = "Bank"
Range("E2") = "Gross Debenture"
Range("F2") = "Net Debenture"
Range("G2") = "Underwriter"
Range("A2:H2").Select
Selection.Font.Underline = xlUnderlineStyleSingle
Selection.Font.Bold = True

End Sub

Sub LiquidatedSheet()
'Creates Liquidated Sheet
Set Liquidated = Sheets.Add
Liquidated.Name = "Liquidated"

Range("A1:G1").Select
    With Selection
        .MergeCells = False
        .HorizontalAlignment = xlCenterAcrossSelection
        .Font.Bold = True
    End With
    ActiveCell.FormulaR1C1 = "Liquidated"

Range("A2") = "Date Liquidated"
Range("B2") = "Known As"
Range("C2") = "SBA Loan #"
Range("D2") = "Bank"
Range("E2") = "Gross Debenture"
Range("F2") = "Net Debenture"
Range("G2") = "Underwriter"
Range("A2:H2").Select
Selection.Font.Underline = xlUnderlineStyleSingle
Selection.Font.Bold = True

End Sub
Sub deletesheets()
'Deletes Dashboard, Authorized, Funded, Paid Off, Liquidated sheets in order for macro to run again
Dim dashsum As Worksheet
Dim dsheet As Worksheet
Set dashsum = Sheets("Dashboard")
Set rptsheet = Sheets("Authorized")
Set fndsheet = Sheets("Funded")
Set pdsheet = Sheets("Paid Off")
Set ldsheet = Sheets("Liquidated")
Set dsheet = Sheets("data")

    Application.DisplayAlerts = False
    dashsum.Delete
    rptsheet.Delete
    fndsheet.Delete
    pdsheet.Delete
    ldsheet.Delete
    Application.DisplayAlerts = True
End Sub






