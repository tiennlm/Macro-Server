Attribute VB_Name = "CapCode"
Sub IP_CADENA()
Dim ar As String, lr, lr1, lr2, er, Ans As Integer, CADENA, CODE, TC, EPC, RH As Worksheet
Dim StartTime As Double
Dim MinutesElapsed As String
Application.ScreenUpdating = False
Application.DisplayAlerts = False
'Remember time when macro starts
  StartTime = Timer
Ans = MsgBox("Xac nhan chuyen thong tin sang mau CADENA?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
If Ans = vbYes Then
    Set CADENA = Worksheets("CADENA")
    Set CODE = Worksheets("1.TT co ban")
    Set EPC = Worksheets("EPC")
    Set RH = Worksheets("Rehire")
    Set TC = Worksheets("Tham chieu")

lr1 = TC.Range("A" & Rows.Count).End(xlUp).Row
lr2 = TC.Range("B" & Rows.Count).End(xlUp).Row

'Gan dong bat dau va dong cuoi cua du lieu
If ActiveSheet.Name <> "1.TT co ban" Then
    MsgBox "Vui long chon o bat dau o Sheet 1.TT co ban"
    Exit Sub
Else
    ar = ActiveCell.Row
'Xoa du lieu hien co tai sheet CADENA
    lr = CADENA.Cells.Find(What:="*", _
                    After:=CADENA.Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
        CADENA.Range("$A$4:$BC$" & 4 + lr).Clear
 'Xoa du lieu hien co tai sheet EPC
    lr = EPC.Cells.Find(What:="*", _
                    After:=EPC.Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
        EPC.Range("$A$5:$AS$" & 5 + lr).Clear
       
 'Xoa du lieu hien co tai sheet Rehired
    lr = RH.Cells.Find(What:="*", _
                    After:=RH.Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
        RH.Range("$A$3:$AZ$" & 3 + lr).Clear
       
    lr = CODE.Range("C" & ar).End(xlDown).Row
    er = TC.Range("A1").End(xlDown).Row

'MsgBox "Dong bat dau: " & ar & " - Den dong: " & lr
J = 4
K = 3
'Thuc hien lenh cho moi dong du lieu
On Error Resume Next
For i = ar To lr
'Input data vao sheet Add new
    CADENA.Range("A" & J).Value = CODE.Range("A" & i)
    CADENA.Range("B" & J).Value = Application.WorksheetFunction.Index(TC.Range("$Q$1:$Q$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("M" & i), TC.Range("$P$1:$P$" & lr2), 0))
    CADENA.Range("E" & J).Value = Right(CODE.Range("C" & i), Len(CODE.Range("C" & i)) - (InStrRev(Trim(CODE.Range("C" & i)), " ")))
    CADENA.Range("G" & J).Value = Left(CODE.Range("C" & i), (InStrRev(Trim(CODE.Range("C" & i)), " ") - 1))
    CADENA.Range("J" & J).Value = "National ID Card"
    CADENA.Range("K" & J).Value = CODE.Range("AB" & i)
    CADENA.Range("L" & J).Value = CODE.Range("AC" & i).Value
    CADENA.Range("M" & J).Value = Format(CODE.Range("AD" & i), "dd-MMM-yyyy")
    CADENA.Range("N" & J).Value = Application.WorksheetFunction.Index(TC.Range("$Z$1:$Z$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("O" & J).Value = CODE.Range("G" & i)
    CADENA.Range("Q" & J).Value = Application.WorksheetFunction.Index(TC.Range("$O$1:$O$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
        
        If CADENA.Range("N" & J).Value = "Full Time" Then
    CADENA.Range("S" & J).Value = Application.WorksheetFunction.Index(TC.Range("$AI$1:$AI$" & er), _
    Application.WorksheetFunction.Match(CODE.Range("J" & i), TC.Range("$A$1:$A$" & er), 0))
    CADENA.Range("T" & J).Value = Application.WorksheetFunction.Index(TC.Range("$AJ$1:$AJ$" & er), _
    Application.WorksheetFunction.Match(CODE.Range("J" & i), TC.Range("$A$1:$A$" & er), 0))
        Else
        End If
        
    CADENA.Range("X" & J).Value = Application.WorksheetFunction.Index(TC.Range("$AA$1:$AA$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("Y" & J).Value = Application.WorksheetFunction.Index(TC.Range("$AB$1:$AB$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("Z" & J).Value = Left(CODE.Range("J" & i), (InStr(CODE.Range("J" & i), " ") - 1)) & CODE.Range("A" & i).Value & _
    "/" & Right(CODE.Range("G" & i), 4) & "/" & Application.WorksheetFunction.Index(TC.Range("$AD$1:$AD$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
        
        If CADENA.Range("X" & J) = "Probation" Then
    CADENA.Range("AA" & J).Value = CODE.Range("G" & i)
    CADENA.Range("AB" & J).Formula = "=DATE(YEAR(AC" & J & "),MONTH(AC" & J & ")+INDEX('Tham chieu'!AC:AC,MATCH(CADENA!Q:Q,'Tham chieu'!O:O,0)),DAY(AC" & J & ")-1)"
        Else
        End If
    
    CADENA.Range("AC" & J).Value = CODE.Range("G" & i)
    CADENA.Range("AD" & J).Formula = "=DATE(YEAR(AC" & J & "),MONTH(AC" & J & ")+INDEX('Tham chieu'!AC:AC,MATCH(CADENA!Q:Q,'Tham chieu'!O:O,0)),DAY(AC" & J & ")-1)"
    CADENA.Range("AE" & J).Value = Application.WorksheetFunction.Index(TC.Range("$AE$1:$AE$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("AF" & J).Value = Application.WorksheetFunction.Index(TC.Range("$R$1:$R$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("M" & i), TC.Range("$P$1:$P$" & lr2), 0))
    CADENA.Range("AG" & J).Value = Application.WorksheetFunction.Index(TC.Range("$AL$1:$AL$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("AH" & J).Value = "Month"
    CADENA.Range("AI" & J).Value = Application.WorksheetFunction.Index(TC.Range("$S$1:$S$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("AJ" & J).Value = Application.WorksheetFunction.Index(TC.Range("$T$1:$T$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("AK" & J).Value = Application.WorksheetFunction.Index(TC.Range("$W$1:$W$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("AL" & J).Value = "VND"
    CADENA.Range("AM" & J).Value = Application.WorksheetFunction.Index(TC.Range("$U$1:$U$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("AN" & J).Value = "By Bank"
    CADENA.Range("AO" & J).Value = "TRUE"
    CADENA.Range("AP" & J).Value = Application.WorksheetFunction.Index(TC.Range("$AG$1:$AG$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("AQ" & J).Value = Application.WorksheetFunction.Index(TC.Range("$AF$1:$AF$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    CADENA.Range("AR" & J).Value = Application.WorksheetFunction.Index(TC.Range("$AH$1:$AH$" & er), _
    Application.WorksheetFunction.Match(CODE.Range("J" & i), TC.Range("$A$1:$A$" & er), 0))
    CADENA.Range("AS" & J).Value = "N/A"
    CADENA.Range("AT" & J).Value = "Stores"
    CADENA.Range("AU" & J).Value = Application.WorksheetFunction.Index(TC.Range("$AH$1:$AH$" & er), _
    Application.WorksheetFunction.Match(CODE.Range("J" & i), TC.Range("$A$1:$A$" & er), 0))
    CADENA.Range("AV" & J).Value = "STORES"
    CADENA.Range("AW" & J).Value = Application.WorksheetFunction.Index(TC.Range("$V$1:$V$" & er), _
    Application.WorksheetFunction.Match(CODE.Range("J" & i), TC.Range("$A$1:$A$" & er), 0))
    CADENA.Range("AX" & J).Value = Application.WorksheetFunction.Index(TC.Range("$N$1:$N$" & er), _
    Application.WorksheetFunction.Match(CODE.Range("J" & i), TC.Range("$A$1:$A$" & er), 0))
    CADENA.Range("AY" & J).Value = CADENA.Range("Q" & J).Value
    
'Input data vao sheet Employees Personal Contact
    EPC.Range("A" & J + 1).Value = CODE.Range("A" & i)
    EPC.Range("C" & J + 1).Value = CADENA.Range("E" & J)
    EPC.Range("E" & J + 1).Value = CADENA.Range("G" & J)
    EPC.Range("H" & J + 1).Value = CODE.Range("Q" & i).Value
    EPC.Range("I" & J + 1).Value = CODE.Range("R" & i).Value
    EPC.Range("J" & J + 1).Value = "NONE"
    EPC.Range("K" & J + 1).Value = "VIETNAMESE"
    EPC.Range("L" & J + 1).Value = "Single"
    EPC.Range("M" & J + 1).Value = Application.WorksheetFunction.Index(TC.Range("$Y$1:$Y$5"), _
    Application.WorksheetFunction.Match(CODE.Range("F" & i), TC.Range("$X$1:$X$5"), 0))
    EPC.Range("O" & J + 1).Value = "Kinh"
    EPC.Range("P" & J + 1).Value = CODE.Range("AA" & i).Value
    EPC.Range("U" & J + 1).Value = CODE.Range("AP" & i)
    EPC.Range("AA" & J + 1).Value = CODE.Range("AN" & i).Value
    EPC.Range("AB" & J + 1).Value = CODE.Range("AN" & i).Value
    EPC.Range("AE" & J + 1).Value = "VIETNAM"
    EPC.Range("AJ" & J + 1).Value = CODE.Range("AL" & i).Value
    EPC.Range("AL" & J + 1).Value = "VIETNAM"
    EPC.Range("AQ" & J + 1).Value = CODE.Range("AM" & i).Value

'Input data vao sheet Rehire
    If CODE.Range("D" & i) = "Rehired" Then
    RH.Range("A" & K).Value = CODE.Range("A" & i)
    RH.Range("B" & K).Value = "Rehired"
    RH.Range("D" & K).Value = CODE.Range("G" & i)
    RH.Range("J" & K).Value = Application.WorksheetFunction.Index(TC.Range("$Q$1:$Q$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("M" & i), TC.Range("$P$1:$P$" & lr2), 0))
    RH.Range("K" & K).Value = Application.WorksheetFunction.Index(TC.Range("$Z$1:$Z$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("L" & K).Value = CODE.Range("G" & i)
    RH.Range("M" & K).Value = Application.WorksheetFunction.Index(TC.Range("$AG$1:$AG$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("N" & K).Value = "PZN"
    RH.Range("O" & K).Value = "Stores"
    RH.Range("P" & K).Value = Application.WorksheetFunction.Index(TC.Range("$AF$1:$AF$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("Q" & K).Value = "N/A"
    RH.Range("R" & K).Value = Application.WorksheetFunction.Index(TC.Range("$O$1:$O$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
        If RH.Range("K" & K).Value = "Full Time" Then
    RH.Range("T" & K).Value = Application.WorksheetFunction.Index(TC.Range("$AI$1:$AI$" & er), _
    Application.WorksheetFunction.Match(CODE.Range("J" & i), TC.Range("$A$1:$A$" & er), 0))
    RH.Range("U" & K).Value = Application.WorksheetFunction.Index(TC.Range("$AJ$1:$AJ$" & er), _
    Application.WorksheetFunction.Match(CODE.Range("J" & i), TC.Range("$A$1:$A$" & er), 0))
        Else
        End If
    RH.Range("Y" & K).Value = Application.WorksheetFunction.Index(TC.Range("$AA$1:$AA$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("Z" & K).Value = Application.WorksheetFunction.Index(TC.Range("$AB$1:$AB$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("AA" & K).Value = Left(CODE.Range("J" & i), (InStr(CODE.Range("J" & i), " ") - 1)) & CODE.Range("A" & i).Value & _
    "/" & Right(CODE.Range("G" & i), 4) & "/" & Application.WorksheetFunction.Index(TC.Range("$AD$1:$AD$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
        
        If RH.Range("Y" & J) = "Probation" Then
    RH.Range("AB" & K).Value = CODE.Range("G" & i)
    RH.Range("AC" & K).Formula = "=DATE(YEAR(AE" & K & "),MONTH(AE" & K & ")+INDEX('Tham chieu'!AC:AC,MATCH(R:R,'Tham chieu'!O:O,0)),DAY(AE" & K & ")-1)"
        Else
        End If
    RH.Range("AE" & K).Value = CODE.Range("G" & i)
    RH.Range("AF" & K).Formula = "=DATE(YEAR(AE" & K & "),MONTH(AE" & K & ")+INDEX('Tham chieu'!AC:AC,MATCH(R:R,'Tham chieu'!O:O,0)),DAY(AE" & K & ")-1)"
    RH.Range("AG" & K).Value = Application.WorksheetFunction.Index(TC.Range("$AE$1:$AE$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("AH" & K).Value = Application.WorksheetFunction.Index(TC.Range("$R$1:$R$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("M" & i), TC.Range("$P$1:$P$" & lr2), 0))
    RH.Range("AI" & K).Value = Application.WorksheetFunction.Index(TC.Range("$AL$1:$AL$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("AJ" & K).Value = "Month"
    RH.Range("AK" & K).Value = Application.WorksheetFunction.Index(TC.Range("$S$1:$S$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("AL" & K).Value = Application.WorksheetFunction.Index(TC.Range("$T$1:$T$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("AM" & K).Value = Application.WorksheetFunction.Index(TC.Range("$W$1:$W$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("AN" & K).Value = "VND"
    RH.Range("AO" & K).Value = Application.WorksheetFunction.Index(TC.Range("$U$1:$U$" & lr2), _
    Application.WorksheetFunction.Match(CODE.Range("K" & i), TC.Range("$B$1:$B$" & lr2), 0))
    RH.Range("AP" & K).Value = "By Bank"
    RH.Range("AQ" & K).Value = "TRUE"
    RH.Range("AR" & K).Value = "PZN"
    RH.Range("AS" & K).Value = "STORES"
    RH.Range("AT" & K).Value = Application.WorksheetFunction.Index(TC.Range("$V$1:$V$" & er), _
    Application.WorksheetFunction.Match(CODE.Range("J" & i), TC.Range("$A$1:$A$" & er), 0))
    RH.Range("AU" & K).Value = Application.WorksheetFunction.Index(TC.Range("$N$1:$N$" & er), _
    Application.WorksheetFunction.Match(CODE.Range("J" & i), TC.Range("$A$1:$A$" & er), 0))
    RH.Range("AV" & K).Value = CADENA.Range("Q" & J).Value
    
    K = K + 1
    Else
    End If
    
J = J + 1
Next i
End If
'Determine how many seconds code took to run
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

'Notify user in seconds
  MsgBox "Thoi gian hoan thanh chuyen thong tin sang mau CADENA: " & MinutesElapsed & " giay", vbInformation
Else
Exit Sub
End If
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub


