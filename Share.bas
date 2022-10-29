Attribute VB_Name = "Share"
Global Const Ver = 1

Sub Exchange_Rate()

Dim lr As Integer, Data As Worksheet, Gdata As Worksheet
Dim Day As Double, ans As Long


'    Application.EnableEvents = False
'    Application.Calculation = False
'    Application.ScreenUpdating = False
'    Application.EnableAnimations = False
'    Application.PrintCommunication = False
     
ans = MsgBox("Confirmation for getting exchange rate data" & _
    vbNewLine & _
    vbNewLine & _
    "The next inputbox is the date to start getting Exchange rate" _
    & vbNewLine & "The precess will get data in 1 month period from start date inputted", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

If ans = vbYes Then

Dim StartTime As Double
Dim MinutesElapsed As String
  StartTime = Timer

Day = Application.InputBox("Input data format dd/mm/yyyy", "Get exchange rate from date?", Format(DateAdd("m", -1, Now()), "16/mm/yyyy"), Type:=1)

    
Dim ws As Worksheet
Dim check As Boolean
    
    'Check sheet "Data" exist?
    For Each ws In Worksheets
        If ws.Name Like "Data" Then check = True: Exit For
    Next
        If check = False Then
    Worksheets.Add.Name = "Data"
    Else
    End If
    
    'Check sheet1 exist?
    For Each ws In Worksheets
        If ws.Name Like "Sheet1" Then check = True: Exit For
    Next
        If check = False Then
    Worksheets.Add.Name = "Sheet1"
    Else
    End If
    
    
    Set Gdata = Worksheets("Sheet1")
    Set Data = Worksheets("Data")
    a = 2
            
            Data.Cells.Clear
            Gdata.Activate
            Cells.Select
            
        If ActiveSheet.QueryTables.Count > 0 Then
            ActiveSheet.QueryTables(1).Delete
        End If
            Selection.ClearContents
    
    For i = Day To DateAdd("m", 1, Day) - 1
    
        With Gdata.QueryTables.Add(Connection:= _
            "URL;https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx?txttungay=" _
            & Format(i, "dd") & "/" & Format(i, "mm") & "/" & Format(i, "yyyy") & "&BacrhID=1&isEn=False" _
            , Destination:=Range("$A$1"))
    '        .CommandType = 0
            .Name = "2021&BacrhID=1&isEn=False"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .WebSelectionType = xlAllTables
            .WebFormatting = xlWebFormattingNone
            .WebPreFormattedTextToColumns = True
            .WebConsecutiveDelimitersAsOne = True
            .WebSingleBlockTextImport = False
            .WebDisableDateRecognition = False
            .WebDisableRedirections = False
            .Refresh BackgroundQuery:=False
        
        End With
        With Data
            .Range("A" & a) = "'" & Format(i, "dd") & "/" & Format(i, "mm") & "/" & Format(i, "yyyy")
            .Range("B" & a) = WorksheetFunction.VLookup("US DOLLAR", Gdata.Range("A:D"), 4, 0)
            .Range("C" & a) = Format(.Range("A" & a), "DDD")
            a = a + 1
        End With
            Gdata.Activate
            Cells.Select
            Selection.QueryTable.Delete
            Selection.ClearContents

    Next
    With Data
    
    lr = .Range("A" & Rows.Count).End(xlUp).Row
    
    .Range("B1") = "Exchange rate(mua chuyen khoan)"
    .Range("A" & lr + 1) = "Average"
    .Range("B" & lr + 1).Formula = "=ROUND(AVERAGE(B2:B" & lr & "),0)"
    .Range("A" & lr + 1 & ":B" & lr + 1).Font.Bold = True
    .Range("A" & lr + 1 & ":B" & lr + 1).Font.Color = vbRed
    .Range("B2:B" & lr + 1).NumberFormat = "#,##0"
    .Columns("B:B").AutoFit
    
    End With
    Data.Activate
    'Determine how many seconds code took to run
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    
'    Application.EnableEvents = True
'    Application.Calculation = True
'    Application.ScreenUpdating = True
'    Application.EnableAnimations = True
'    Application.PrintCommunication = True

    MsgBox "Finished getting Exchange rate from VCB website" & _
            vbNewLine & vbNewLine & "Estimate Processing Time: " & MinutesElapsed, vbInformation
Else
End If
End Sub


Sub Payroll_data()

Dim lr As Integer, ans As Long
Dim NH As Worksheet, RS As Worksheet, Pro As Worksheet, Leave As Worksheet, OT As Worksheet
Dim DataNH As Worksheet, DataRS As Worksheet, DataPro As Worksheet, DataLeave As Worksheet, DataOT As Worksheet

ans = MsgBox("Confirmation process Payroll Data", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

If ans = yes Then

    Application.Calculation = False
    Application.ScreenUpdating = False
    
    Set NH = Worksheets("New hire")
    Set RS = Worksheets("Resigner")
    Set Pro = Worksheets("Probation")
    Set Leave = wordsheets("Leave")
    Set OT = Worksheets("OT")
    Set DataNH = Worksheets("Data NH")
    Set DataRS = Worksheets("Data RS")
    Set DataPro = Worksheets("Data Probation")
    Set DataLeave = wordsheets("Data Leave")
    Set DataOT = Worksheets("Data OT")
    
    
    
    
    
    Application.Calculation = True
    Application.ScreenUpdating = True

Else
End If
End Sub

Sub Open_CNB()

Dim fd As Office.FileDialog
Dim strFile As String, ans As Long

ans = Application.InputBox("Input year", "Which year of file", Format(Now(), "yyyy"), Type:=1)
If ans = cancel Then
Exit Sub

End If
Set fd = Application.FileDialog(msoFileDialogFilePicker)


With fd

    .Filters.Clear
    .Filters.Add "Excel Files", "*.xlsx, *.xls", 1
    .Title = "Choose an Excel file"
    .AllowMultiSelect = False
 
    .InitialFileName = "\\172.16.6.6\Human Resources\2. C&B\01. Payroll-PMH\2021"
 
    If .Show = True Then
 
        strFile = .SelectedItems(1)
 
    End If
    If .SelectedItems.Count = 0 Then
    Exit Sub
    End If

End With
Workbooks.Open filename:=strFile, _
Password:="CNB@" & ans - 1 & "$", ReadOnly:=False
Application.DisplayAlerts = True
End Sub

Sub SMD_CTV_ATT()

Dim lr As Long, ans As Long, ans1 As String
Dim Data As Worksheet, HS As Worksheet, RP As Worksheet, SMD As Worksheet


ans = MsgBox("Confirmation process SMD CTV Attendance", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

If ans = vbYes Then

For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Data" Then
        exists = True
    End If
Next i

If Not exists Then
    Exit Sub
End If

ans1 = Application.InputBox("Nhap ten file", "Ten file?", "SMD Attendance " & Format(DateAdd("m", -1, Date), "mm.yyyy"), Type:=2)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
        
Set Data = Worksheets("Data")
Set HS = Worksheets("History Recognition")
Set RP = Worksheets("Report")
    
    'On Error Resume Next
    HS.Cells.Delete
    With RP
        lr = .Range("B" & Rows.Count).End(xlUp).Row
        .Range("A2:E" & lr + 1).Delete
    End With
    
    With Data
        .Cells.ClearOutline
        
        lr = .Range("C" & Rows.Count).End(xlUp).Row
        
        .Range("5:5").AutoFilter field:=8, Criteria1:="PMH - CTV"
        .Range("A5:E" & lr).copy HS.Range("A1")
    End With

    With HS
         .Cells.ClearOutline

        lr = .Range("C" & Rows.Count).End(xlUp).Row

            HS.Sort.SortFields.Clear
            HS.Sort.SortFields.Add Key:= _
                Range("C2:C" & lr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
                :=xlSortNormal
            HS.Sort.SortFields.Add Key:= _
                Range("E2:E" & lr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
                :=xlSortNormal
            With HS.Sort
                .SetRange Range("A1:G" & lr)
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        .Range("F1") = "Day"
        .Range("G1") = "Time"
        .Range("B2").Formula = "=C2&F2"
        .Range("B2:B" & lr).FillDown
        .Range("F2").Formula = "=VALUE(LEFT(E2,10))"
        .Range("F2:F" & lr).FillDown
        .Range("F2:F" & lr).NumberFormat = "dd/mm/yyyy"
        .Range("G2:G" & lr).Formula = "=VALUE(RIGHT(E2,8))"
        .Range("G2:G" & lr).FillDown
        .Range("G2:G" & lr).NumberFormat = "hh:mm:ss"
        .Range("A:F").EntireColumn.AutoFit
        .Columns("E:E").Columns.Group
        .Outline.ShowLevels RowLevels:=0, ColumnLevels:=1

        HS.Select
        .Range("C2:F" & lr).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Application.CutCopyMode = False
        Selection.copy
        Sheets("Report").Select
        Range("A2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    End With


    With RP
        lr = .Range("A" & Rows.Count).End(xlUp).Row
        .Range("F2").Formula = "=A2&C2"
        .Range("F2:F" & lr).FillDown
        .Range("$A$1:$F$" & lr).RemoveDuplicates Columns:=6, Header:=xlYes
        .Range("A1:F" & lr).Select
        .Range("B7").Activate
        .Range("D2").Formula = "=INDEX('History Recognition'!G:G,MATCH(Report!F2,'History Recognition'!B:B,0))"
        .Range("E2").Formula = "=INDEX('History Recognition'!G:G,MATCH(Report!F2,'History Recognition'!B:B,0)-1+COUNTIF('History Recognition'!B:B,F2))"

        lr = .Range("A" & Rows.Count).End(xlUp).Row
        .Range("D2:E" & lr).FillDown
        .Range("C2:C" & lr).NumberFormat = "dd/mm/yyyy"
        .Range("D2:E" & lr).NumberFormat = "hh:mm:ss"

        Columns("D:E").Select
        Selection.copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Columns("F:F").Select
        Application.CutCopyMode = False
        Selection.ClearContents

    End With

            Data.Select
            Range("G7").Select
            With ActiveWindow
                If .FreezePanes Then .FreezePanes = False
                .SplitColumn = 0
                .SplitRow = 0
                .FreezePanes = True
            End With

            HS.Select
            Range("D3").Select
            With ActiveWindow
                If .FreezePanes Then .FreezePanes = False
                .SplitColumn = 0
                .SplitRow = 0
                .FreezePanes = True
            End With

            RP.Select
            Range("D3").Select
            With ActiveWindow
                If .FreezePanes Then .FreezePanes = False
                .SplitColumn = 0
                .SplitRow = 0
                .FreezePanes = True
            End With
    ActiveWorkbook.Save
'    On Error GoTo 0
    Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs filename:=ActiveWorkbook.Path _
        & "\" & ans1 & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Sheets("Data").Select
    ActiveWindow.SelectedSheets.Delete
    ActiveWorkbook.Save
ElseIf ans = vbNo Then

Exit Sub
End If


End Sub

Sub Unmerge_T2C()
    Columns("B:B").UnMerge
    Range("A1").Select
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
End Sub

Sub ATT_Check()

Application.EnableEvents = False
Application.Calculation = False
Application.ScreenUpdating = False
Application.EnableAnimations = False
     
        For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "Attendance Detail" Then
            exists = True
        End If
        Next i
    
        If Not exists Then
            MsgBox "Sheet Attendance Detail khong ton tai, kiem tra lai file"
        Exit Sub
        End If
    On Error Resume Next
    AutoFilterMode = False
    Rows("6:6").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Range("$A$6:$BU$21025").AutoFilter field:=72, Criteria1:= _
        "=Không*", Operator:=xlOr, Criteria2:="=*1*"
    ActiveSheet.Range("$A$6:$BU$21025").AutoFilter field:=64, Criteria1:="<>"
    
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.EnableAnimations = True
End Sub

Sub PIT_Monthly()

Dim SalLocal As Worksheet, SalExpat As Worksheet, Other As Worksheet, Monthly As Worksheet, REF As Worksheet
Dim lr As Long, lr1 As Long, ans As Integer

ans = MsgBox("Confirmation process Monthly PIT", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

If ans = vbYes Then

    Application.Calculation = xlManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False

For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Monthly PIT" Then
        exists = True
    End If
Next i

If Not exists Then
    Exit Sub
End If

Set SalLocal = Worksheets("Sal-Local")
Set SalExpat = Worksheets("Sal-Expat")
Set Other = Worksheets("Other")
Set Monthly = Worksheets("Monthly PIT")
Set REF = Worksheets("REF")

Monthly.Select

With Monthly
    lr = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    .Range("5:" & lr + 5).Clear
End With

If SalLocal.AutoFilterMode Then
     SalLocal.AutoFilterMode = False
  End If
If SalExpat.AutoFilterMode Then
     SalExpat.AutoFilterMode = False
  End If
If Other.AutoFilterMode Then
     Other.AutoFilterMode = False
  End If
If Monthly.AutoFilterMode Then
     Monthly.AutoFilterMode = False
  End If

'SalLocal.Select
    With SalLocal

    lr = .Range("B" & Rows.Count).End(xlUp).Row
        .Range("6:" & lr).AutoFilter field:=107, Criteria1:=Right(Monthly.Range("C1"), 4) & Left(Monthly.Range("C1"), 2)
    .Range("B7:B" & lr).copy Monthly.Range("B5")
    .Range("G7:G" & lr).copy Monthly.Range("C5")
    End With


lr1 = Monthly.Range("B" & Rows.Count).End(xlUp).Row

'SalExpat.Select
With SalExpat
lr = .Range("C" & Rows.Count).End(xlUp).Row
    .Range("6:" & lr).AutoFilter field:=102, Criteria1:=Right(Monthly.Range("C1"), 4) & Left(Monthly.Range("C1"), 2)
    .Range("C7:C" & lr + 1).copy Monthly.Range("B" & lr1 + 1)
    .Range("F7:F" & lr + 1).copy Monthly.Range("C" & lr1 + 1)
End With

lr1 = Monthly.Range("B" & Rows.Count).End(xlUp).Row

'Other.Select
With Other
lr = .Range("B" & Rows.Count).End(xlUp).Row
    .Range("6:" & lr).AutoFilter field:=16, Criteria1:=Right(Monthly.Range("C1"), 4) & Left(Monthly.Range("C1"), 2)
    .Range("B7:C" & lr + 1).copy Monthly.Range("B" & lr1 + 1)
End With

lr1 = Monthly.Range("B" & Rows.Count).End(xlUp).Row

With Monthly
lr1 = Monthly.Range("B" & Rows.Count).End(xlUp).Row
    .Range("$B$5:$C$" & lr1 + 5).RemoveDuplicates Columns:=1, Header:=xlNo
    .Range("4:" & lr1).AutoFilter
    
lr1 = Monthly.Range("B" & Rows.Count).End(xlUp).Row
    .Range("C2").Formula = "=TEXT(C1,""yyyymm"")"
    .Range("E5").Formula = "=SUMIFS('Sal-Local'!$BN:$BN,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUMIFS('Sal-Expat'!$BF:$BF,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUM(SUMIFS(Other!$G:$G,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm""),Other!$U:$U,{""Sal-Phu cap CTV Sale"",""Sal-Allowance*""}))"
    .Range("F5").Formula = "=SUMIFS('Sal-Local'!$BO:$BO,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm"")) +SUMIFS('Sal-Local'!$BP:$BP,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm"")) +SUMIFS('Sal-Local'!$BQ:$BQ,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm""))+ SUMIFS('Sal-Expat'!$BH:$BH,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm"")) + SUMIFS('Sal-Expat'!$BG:$BG,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))"
    .Range("G5").Formula = "=SUMIFS('Sal-Local'!$BR:$BR,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm"")) + SUMIFS('Sal-Expat'!$BI:$BI,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))"
    .Range("H5").Formula = "=SUM(SUMIFS(Other!$G:$G,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm""),Other!$U:$U,{""Bonus-Extra"",""Bonus-ThuPhi"",""Hoa hong moi tai tro"",""Bonus-Retention"",""Special Bonus"",""Home Leave""}))"
    .Range("I5").Formula = "=SUM(SUMIFS(Other!$G:$G,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm""),Other!$U:$U,{""Bonus-YE"",""Indirect Bonus""}))"
    .Range("J5").Formula = "=SUM(SUMIFS(Other!$G:$G,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm""),Other!$U:$U,{""Bonus-TT2"",""Leasing Bonus"",""Acc-Phi moi gioi"",""PSD Petro"",""Mobile""}))"
    .Range("K5").Formula = "=SUMIFS(Other!$G:$G,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm""),Other!$U:$U,""<>Sal-Phu cap CTV Sale"",Other!$Y:$Y,""<>Sal-Allowance*"")-$H5-$I5-$J5"
    .Range("L5").Formula = "=E5 + F5 - G5 + H5 + J5 + K5 + I5"
    .Range("M5").Formula = "=SUMIFS('Sal-Local'!$BT:$BT,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUMIFS('Sal-Expat'!$BK:$BK,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))"
    .Range("N5").Formula = "=SUMIFS('Sal-Local'!$BU:$BU,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUMIFS('Sal-Expat'!$BL:$BL,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))"
    .Range("O5").Formula = "=SUMIFS('Sal-Local'!$BV:$BV,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUMIFS('Sal-Expat'!$BM:$BM,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))"
    .Range("P5").Formula = "=SUMIFS('Sal-Local'!$BW:$BW,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUMIFS('Sal-Expat'!$BN:$BN,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))"
    .Range("Q5").Formula = "=SUMIFS('Sal-Local'!$BX:$BX,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUMIFS('Sal-Expat'!$BO:$BO,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))"
    .Range("R5").Formula = "=SUMIFS('Sal-Local'!$BY:$BY,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUMIFS('Sal-Expat'!$BP:$BP,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))"
    .Range("T5").Formula = "=IF(AC5=0,L5-M5-O5-P5-Q5-R5-S5-K5-H5-J5-I5,ROUND(IF(L5-M5-O5-P5-Q5-R5-S5-K5-H5-J5-I5<0,0,L5-M5-O5-P5-Q5-R5-S5-K5-H5-J5-I5+AD5),0))"
    .Range("U5").Formula = "=ROUND(IF(AC5=""CK"",0,IF(AC5=""Flat 10%"",IF(T5<2000000,0,T5*0.1),IF(AC5=""Flat rate 10%"",T5*0.1,IF(or(AC5=""FLAT 20%"",AC5=""FLAT rate 20%""),max(0,L5-H5-I5-J5-K5)*0.2,IF(T5<=5000000,T5*5%,IF(T5<=10000000,(T5*10%)-250000,IF(T5<=18000000,(T5*15%)-750000,IF(T5<=32000000,(T5*20%)-1650000,IF(T5<=52000000,(T5*25%)-3250000,IF(T5<=80000000,(T5*30%)-5850000,IF(T5>80000000,(T5*35%)-9850000))))))))))),0)+AF5"
    .Range("V5").Formula = "=SUM(SUMIFS(Other!$I:$I,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm""),Other!$U:$U,{""Bonus-Extra"",""Bonus-ThuPhi"",""Hoa hong moi tai tro""}))"
    .Range("W5").Formula = "=SUM(SUMIFS(Other!$I:$I,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm""),Other!$U:$U,{""Bonus-YE"",""Indirect Bonus""}))"
    .Range("X5").Formula = "=SUM(SUMIFS(Other!$I:$I,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm""),Other!$U:$U,{""Bonus-TT2"",""Leasing Bonus"",""Acc-Phi moi gioi"",""PSD Petro"",""Mobile""}))"
    .Range("Y5").Formula = "=SUMIFS(Other!$I:$I,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm""),Other!$U:$U,""<>Sal-Phu cap CTV Sale"",Other!$U:$U,""<>Sal-Allowance*"")-$V5-$W5-$X5"
    .Range("Z5").Formula = "=SUM(U5:Y5)"
    .Range("AA5").Formula = "=SUMIFS('Sal-Local'!$CC:$CC,'Sal-Local'!$B:$B,'Monthly PIT'!$B5,'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUMIFS('Sal-Expat'!$BU:$BU,'Sal-Expat'!$C:$C,'Monthly PIT'!$B5,'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUM(SUMIFS(Other!$I:$I,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm"")))-Z5"
    .Range("AB5").Formula = "=IF(ISERROR(INDEX('Sal-Local'!D:D,MATCH(B5&$C$2,'Sal-Local'!DI:DI,0))),IF(ISERROR(INDEX('Sal-Expat'!E:E,MATCH(B5&$C$2,'Sal-Expat'!DD:DD,0))),INDEX(Other!E:E,MATCH(B5&$C$2,Other!W:W,0)),INDEX('Sal-Expat'!E:E,MATCH(B5&$C$2,'Sal-Expat'!DD:DD,0))),INDEX('Sal-Local'!D:D,MATCH(B5&$C$2,'Sal-Local'!DI:DI,0)))"
    .Range("AC5").Formula = "=IF(ISERROR(INDEX('Sal-Local'!DH:DH,MATCH(B5&$C$2,'Sal-Local'!DI:DI,0))),IF(ISERROR(INDEX('Sal-Expat'!DC:DC,MATCH(B5&$C$2,'Sal-Expat'!DD:DD,0))),INDEX(Other!V:V,MATCH(B5&$C$2,Other!W:W,0)),INDEX('Sal-Expat'!DC:DC,MATCH(B5&$C$2,'Sal-Expat'!DD:DD,0))),INDEX('Sal-Local'!DH:DH,MATCH(B5&$C$2,'Sal-Local'!DI:DI,0)))"
    .Range("AM5").Formula = "=IF(AND(L5=0,M5=0),1,0)"
    .Range("AN5").Formula = "=IFERROR(IF(ISERROR(INDEX('Sal-Local'!DB:DB,MATCH(B5&$C$2,'Sal-Local'!DI:DI,0))),INDEX('Sal-Expat'!CW:CW,MATCH(B5&$C$2,'Sal-Expat'!DD:DD,0)),INDEX('Sal-Local'!DB:DB,MATCH(B5&$C$2,'Sal-Local'!DI:DI,0))),$C$2)"
    
    .Range("E5:AC" & lr1).FillDown
    .Range("AM5:AN" & lr1).FillDown
    .Range("E5:AA" & lr1).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    .Range("E" & lr1 + 1).Formula = "=AGGREGATE(9,3,E5:E" & lr1 & ")"
    .Range("E" & lr1 + 1 & ":Z" & lr1 + 1).FillRight
    .Range("E" & lr1 + 1 & ":Z" & lr1 + 1).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
    .Range("A5").Formula = "=IF(B5="""","""",AGGREGATE(3,3,$B$5:B5))"
    .Range("A5:A" & lr1).FillDown
    
    .Range("A" & lr1 + 1 & ":Z" & lr1 + 1).Select
    Range("Z" & lr1 + 1).Activate
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
End With

Else: Exit Sub
End If

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableAnimations = True

With Monthly

    lr1 = Monthly.Range("B" & Rows.Count).End(xlUp).Row
    
    For i = 5 To lr1
        If .Range("AC" & i) = "Flat 20%" Then
        .Range("AF" & i).Formula = "=SUMIFS('Sal-Local'!$DK:$DK,'Sal-Local'!$C:$C,'Monthly PIT'!$B" & i & ",'Sal-Local'!$DC:$DC,TEXT('Monthly PIT'!$C$1,""yyyymm""))+SUMIFS('Sal-Expat'!$DE:$DE,'Sal-Expat'!$C:$C,'Monthly PIT'!$B" & i & ",'Sal-Expat'!$CX:$CX,TEXT('Monthly PIT'!$C$1,""yyyymm""))"
        .Range("AG" & i).Formula = "=SUMIFS(Other!Y:Y,Other!B:B,'Monthly PIT'!$B" & i & ",Other!P:P,TEXT('Monthly PIT'!$C$1,""yyyymm""))"
        .Range("V" & i).Formula = "=H" & i & "*0.2"
        .Range("W" & i).Formula = "=I" & i & "*0.2"
        .Range("X" & i).Formula = "=J" & i & "*0.2"
        .Range("Y" & i).Formula = "=K" & i & "*0.2"
        ElseIf .Range("AC" & i) = "Flat rate 10%" Then
        .Range("AH" & i).Formula = "=SUM(AI" & i & ":AL" & i & ")-V" & i & "-W" & i & "-X" & i & "-Y" & i
        .Range("AI" & i).Formula = "=IF($AC" & i & "=""Flat rate 10%"",H" & i & "*0.1,0)"
        .Range("AJ" & i).Formula = "=IF($AC" & i & "=""Flat rate 10%"",I" & i & "*0.1,0)"
        .Range("AK" & i).Formula = "=IF($AC" & i & "=""Flat rate 10%"",J" & i & "*0.1,0)"
        .Range("AL" & i).Formula = "=IF($AC" & i & "=""Flat rate 10%"",K" & i & "*0.1,0)"
        Else
        'Nothing
        End If
    Next i
    
    For i = 5 To lr1
        If .Range("AN" & i) <> .Range("C2") Then
            .Range("U" & i).Value = WorksheetFunction.SumIfs(SalLocal.Range("CC:CC"), SalLocal.Range("B:B"), .Range("B" & i), SalLocal.Range("DC:DC"), .Range("C2")) + WorksheetFunction.SumIfs(SalExpat.Range("BU:BU"), SalExpat.Range("C:C"), .Range("B" & i), SalExpat.Range("CX:CX"), .Range("C2"))
            .Range("AE" & i) = "Salary Revised"
            Rows(i & ":" & i).Select
            With Selection.Interior
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    
    For i = 5 To lr1
        If .Range("L" & i) = 0 And .Range("AE" & i) = "Salary Revised" Then
        Rows(i & ":" & i).Select
        Selection.Delete Shift:=xlUp
        i = i - 1
        ElseIf .Range("AM" & i) = 1 Then
        Rows(i & ":" & i).Delete Shift:=xlUp
        i = i - 1
        Else
        End If
        
    Next i
    
    
End With
MsgBox "Done - Check | LWD | Tax period | PIT revised | Column AA |"
End Sub

Sub BCLD()
Dim DSNV As Worksheet, BCLD As Worksheet, HBCLD As Worksheet, REF As Worksheet
Dim lr As Long, ans As Long, dDate As Long

ans = MsgBox("Confirmation process Labour Report", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

If ans = vbYes Then

    Application.Calculation = xlManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    
Set DSNV = Worksheets("DSNV")
Set BCLD = Worksheets("BCLD")
Set HBCLD = Worksheets("History BCLD")
Set REF = Worksheets("REF")

If DSNV.AutoFilterMode Then
     DSNV.AutoFilterMode = False
  End If

dDate = DateSerial(2021, 11, 26)
With DSNV
        .Range("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

    lr = .Range("A" & Rows.Count).End(xlUp).Row
    .Range("1:1").AutoFilter field:=88, Criteria1:=">=" & dDate, Operator:=xlOr, Criteria2:="="
    
End With
    
    
Else: Exit Sub
End If

End Sub


Sub Newhired()
Dim NH As Worksheet, SalLocal As Worksheet, SalExpat As Worksheet, REF As Worksheet
Dim lr As Long, ans As Long, dDate As Long

ans = MsgBox("Confirmation process Newhire", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

If ans = vbYes Then

    Set NH = Worksheets("New")
    Set SalLocal = Worksheets("Sheet1")
    
    With SalLocal
        lr = NH.Range("B" & Rows.Count).End(xlUp).Row
        
    End With

Else: Exit Sub
End If

MsgBox lr
End Sub

Sub Income()
Dim SalLocal As Worksheet, SalExpat As Worksheet, Other As Worksheet, Monthly As Worksheet, REF As Worksheet, Income As Worksheet
Dim lr As Long, lr1 As Long, ans As Integer

ans = MsgBox("Confirmation process Income report", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")

If ans = vbYes Then

    Application.Calculation = xlManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False

For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Income Report" Then
        exists = True
    End If
Next i

If Not exists Then
    Exit Sub
End If
Set SalLocal = Worksheets("Sal-Local")
Set SalExpat = Worksheets("Sal-Expat")
Set Other = Worksheets("Other")
Set Monthly = Worksheets("Monthly PIT")
Set REF = Worksheets("REF")
Set Income = Worksheets("Income Report")

With Income
    lr = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    .Range("4:" & lr + 5).Clear
End With

SalLocal.Select
Dim IncomeData() As Variant
Dim MNV() As Variant, FName() As Variant, Dept() As Variant, ExRate() As Variant, Month() As Variant
Dim SalVND() As Variant, SalUSD As Variant                          'Base Salary
Dim PossVND() As Variant, PossUSD As Variant                        'Position Allowance
Dim ProVND() As Variant, ProUSD As Variant                          'Pro. Allowance
Dim ToxicVND() As Variant, ToxicUSD As Variant                      'Toxic Allowance
Dim OTVND() As Variant, OTUSD As Variant                            'OT Allowance
Dim ConsultVND() As Variant, ConsultUSD As Variant                  'Consultancy Allowance
Dim ExpatVND() As Variant, ExpatUSD As Variant                      'Expat Allowance
Dim DepenVND() As Variant, DepenUSD As Variant                      'Dependant Allowance
Dim DutyVND() As Variant, DutyUSD As Variant                        'Duty Allowance
Dim ReloVND() As Variant, ReloUSD As Variant                        'Relocation allowance
Dim MealVND() As Variant, MealUSD As Variant                        'Lunch Allowance
Dim MobiVND() As Variant, MobiUSD As Variant                        'Mobile Allowance
Dim TransportVND() As Variant, TransportUSD As Variant              'Transport Allowance
Dim HardshipVND() As Variant, HardshipUSD As Variant                'Hardship
Dim HandoverVND() As Variant, HandoverUSD As Variant                'Inspection & Handover Allowance
Dim SECVND() As Variant, SECUSD As Variant                          'Security Allowance
Dim AllowVND() As Variant, AllowUSD As Variant                      'Other Allowance
Dim SalAdjustVND() As Variant, SalAdjustUSD As Variant              'Salary Adjustment
Dim UleaveVND() As Variant, UleaveUSD As Variant                    'Unused Annual leave
Dim HolidayVND() As Variant, HolidayUSD As Variant                  'Holiday Bonus
Dim BirthdayAdjustVND() As Variant, BirthdayAdjustUSD As Variant    'Birthday Gift
Dim HLVND() As Variant, HLUSD As Variant                            'Encashment unused HL & airticket 2021
Dim IndVND() As Variant, IndUSD As Variant                          'Indirect Bonus
Dim DirectVND() As Variant, DirectUSD As Variant                    'Direct Bonus
Dim EfficiencyVND() As Variant, EfficiencyUSD As Variant            'Efficiency Bonus
Dim PAVND() As Variant, PAUSD As Variant                            'PA bonus
Dim YEVND() As Variant, YEUSD As Variant                            'Year-end bonus
Dim BonusVND() As Variant, BonusUSD As Variant                      'Other bonus
Dim TVVND() As Variant, TVUSD As Variant                            'Severance Allowance
Dim BikVND() As Variant, BikUSD As Variant                          'Benefit in kind
Dim HouseVND() As Variant, HouseUSD As Variant                      'Housing Allowance
Dim SSISVND() As Variant, SSISUSD As Variant                        'Tuition fee SSIS
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Dim PITVND() As Variant, PITUSD As Variant                          'PIT
Dim QTTVND() As Variant, QTTUSD As Variant                          'PIT Finalization
Dim BHXHVND() As Variant, BHXHUSD As Variant                        'Social Insurance
Dim BHYTVND() As Variant, BHYTUSD As Variant                        'Health Insurance
Dim BHTNVND() As Variant, BHTNUSD As Variant                        'Unemployment Insurance
Dim UnionVND() As Variant, UnionUSD As Variant                      'Union fee
Dim DaTVND() As Variant, DaTUSD As Variant                          'Deduct after tax
Dim DeductVND() As Variant, DeductUSD As Variant                    'Other deduction



    With SalLocal
    lr1 = .Range("C" & Rows.Count).End(xlUp).Row

        MNV = .Range("B7:B" & lr1).Value
        FName = .Range("G7:G" & lr1).Value
        Dept = .Range("D7:D" & lr1).Value
        Description = .Range("DD7:DD" & lr1).Value
        ExRate = .Range("DE7:DE" & lr1).Value
        Month = .Range("DL7:DL" & lr1).Value
        SalVND = .Range("AL7:AL" & lr1).Value
        PossVND = .Range("AM7:AM" & lr1).Value
        ToxicVND = .Range("AO7:AO" & lr1).Value
        OTVND = .Range("BD7:BD" & lr1).Value
        DutyVND = .Range("AN7:AN" & lr1).Value
    
    End With
        ReDim IncomeData(1 To UBound(MNV, 1), 70) As Variant
    For i = 1 To UBound(MNV, 1)
        IncomeData(i, 0) = MNV(i, 1)
        IncomeData(i, 1) = FName(i, 1)
        IncomeData(i, 2) = Dept(i, 1)
        IncomeData(i, 3) = Description(i, 1)
        IncomeData(i, 4) = ExRate(i, 1)
        IncomeData(i, 5) = Month(i, 1)
    Next i

Income.Select

    With Income
       .Range("A4").Resize(UBound(IncomeData, 1), UBound(IncomeData, 2)) = IncomeData
    End With
    
    Erase IncomeData

    With SalExpat
        lr1 = .Range("C" & Rows.Count).End(xlUp).Row

        MNV = .Range("C7:C" & lr1).Value
        FName = .Range("F7:F" & lr1).Value
        Dept = .Range("E7:E" & lr1).Value
        Description = .Range("CY7:CY" & lr1).Value
        ExRate = .Range("CZ7:CZ" & lr1).Value
        Month = .Range("DG7:DG" & lr1).Value
    End With
        ReDim IncomeData(1 To UBound(MNV, 1), 70) As Variant
    For i = 1 To UBound(MNV, 1)
        IncomeData(i, 0) = MNV(i, 1)
        IncomeData(i, 1) = FName(i, 1)
        IncomeData(i, 2) = Dept(i, 1)
        IncomeData(i, 3) = Description(i, 1)
        IncomeData(i, 4) = ExRate(i, 1)
        IncomeData(i, 5) = Month(i, 1)
    Next i

Income.Select
    lr = Income.Range("A" & Rows.Count).End(xlUp).Row
    
    With Income
       .Range("A" & lr + 1).Resize(UBound(IncomeData, 1), UBound(IncomeData, 2)) = IncomeData
    End With


If SalLocal.AutoFilterMode Then
     SalLocal.AutoFilterMode = False
  End If
If SalExpat.AutoFilterMode Then
     SalExpat.AutoFilterMode = False
  End If
If Other.AutoFilterMode Then
     Other.AutoFilterMode = False
  End If
If Monthly.AutoFilterMode Then
     Monthly.AutoFilterMode = False
  End If
  

End If
End Sub


Private VBE As VBIDE.VBE

Public Sub FindProject()
  
  Dim project As VBIDE.VBProject
  Set project = Workbooks("Share.xlam").VBProject
   MsgBox project.VBE.VBProjects
  
End Sub

Sub GetModules()
Dim modName As String
Dim wb As Workbook
Dim l As Long

Set wb = Workbooks("Share.xlam")

For l = 1 To wb.VBProject.VBComponents.Count
With wb.VBProject.VBComponents(l)
modName = modName & vbCr & .Name
End With
Next

MsgBox "Module Names:" & vbCr & modName

Set wb = Nothing

End Sub


Sub DeleteVBComponent(ByVal wb As Workbook, ByVal CompName As String)
'Disabling the alert message
Application.DisplayAlerts = False
'Ignore errors
On Error Resume Next
'Delete the component
wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(CompName)
On Error GoTo 0
'Enabling the alert message
Application.DisplayAlerts = True
End Sub
Sub calling_procedure()
    'Calling DeleteVBComponent macro
    DeleteVBComponent Workbooks("Share.xlam"), "Test"
End Sub

