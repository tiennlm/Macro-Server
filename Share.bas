Attribute VB_Name = "Share"
Global Const Ver = 1

Sub Exchange_Rate_Share()

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


Sub PIT_Monthly_Share()

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
    .Range("K5").Formula = "=SUMIFS(Other!$G:$G,Other!$B:$B,'Monthly PIT'!$B5,Other!$P:$P,TEXT('Monthly PIT'!$C$1,""yyyymm""),Other!$U:$U,""<>Sal-Phu cap CTV Sale"",Other!$U:$U,""<>Sal-Allowance*"")-$H5-$I5-$J5"
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

