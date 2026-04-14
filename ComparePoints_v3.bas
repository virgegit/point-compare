Attribute VB_Name = "ComparePoints"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ComparePoints – Excel VBA macro
' Compares two sheets of measurement points.
' Settings are read from the config sheet.
' Version: 3.1  |  April 2026
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' ── Read settings from the config sheet ─────────────────────
Function CfgStr(key As String) As String
    Dim ws As Worksheet
    Set ws = GetConfigSheet(ThisWorkbook)
    If ws Is Nothing Then
        CfgStr = ""
        Exit Function
    End If
    Dim i As Long
    For i = 1 To 200
        If CStr(ws.Cells(i, 2).Value) = key Then
            CfgStr = CStr(ws.Cells(i, 3).Value)
            Exit Function
        End If
    Next i
    CfgStr = ""
End Function

Function CfgDbl(key As String, def As Double) As Double
    Dim s As String: s = CfgStr(key)
    If Len(s) = 0 Then CfgDbl = def: Exit Function
    On Error Resume Next
    CfgDbl = CDbl(s)
    If Err.Number <> 0 Then CfgDbl = def: Err.Clear
    On Error GoTo 0
End Function

Function CfgBool(key As String, def As Boolean) As Boolean
    Dim s As String: s = LCase(Trim(CfgStr(key)))
    If s = "true" Or s = "1" Or s = "yes" Then
        CfgBool = True
    ElseIf s = "false" Or s = "0" Or s = "no" Then
        CfgBool = False
    Else
        CfgBool = def
    End If
End Function

Function PaletteColor(key As String, def As Long) As Long
    Dim s As String: s = UCase(Trim(CfgStr(key)))
    If Len(s) = 6 Then
        On Error Resume Next
        PaletteColor = RGB(CLng("&H" & Left(s,2)), CLng("&H" & Mid(s,3,2)), CLng("&H" & Right(s,2)))
        If Err.Number <> 0 Then PaletteColor = def: Err.Clear
        On Error GoTo 0
    Else
        PaletteColor = def
    End If
End Function

Function ColIdx(ws As Worksheet, hdrRow As Long, colName As String) As Long
    If Len(colName) = 0 Then ColIdx = 0: Exit Function
    Dim c As Long
    For c = 1 To 50
        If CStr(ws.Cells(hdrRow, c).Value) = colName Then ColIdx = c: Exit Function
    Next c
    ColIdx = 0
End Function

' Result array column constants
Const R_STATUS    As Long = 1
Const R_F1_NAME   As Long = 2:  Const R_F1_X As Long = 3
Const R_F1_Y      As Long = 4:  Const R_F1_Z As Long = 5
Const R_F1_I      As Long = 6:  Const R_F1_J As Long = 7:  Const R_F1_K As Long = 8
Const R_F2_NAME   As Long = 9:  Const R_F2_X As Long = 10
Const R_F2_Y      As Long = 11: Const R_F2_Z As Long = 12
Const R_F2_I      As Long = 13: Const R_F2_J As Long = 14: Const R_F2_K As Long = 15
Const R_NAME_DIFF As Long = 16
Const R_DX        As Long = 17: Const R_DY As Long = 18: Const R_DZ As Long = 19
Const R_DI        As Long = 20: Const R_DJ As Long = 21: Const R_DK As Long = 22
Const R_DIFF_FLD  As Long = 23
Const R_COLS      As Long = 23

'===========================================================
' MAIN
'===========================================================
Sub ComparePoints()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim t0 As Double: t0 = Timer

    Dim sh1     As String:  sh1    = CfgStr("SRC_SHEET1"):  If Len(sh1)=0    Then sh1 = "☕ Sheet 1"
    Dim sh2     As String:  sh2    = CfgStr("SRC_SHEET2"):  If Len(sh2)=0    Then sh2 = "☕ Sheet 2"
    Dim hdrRow  As Long:    hdrRow = CLng(CfgStr("HDR_ROW")):If hdrRow=0 Then hdrRow = 5
    Dim tolXYZ  As Double:  tolXYZ  = CfgDbl("TOL_XYZ",  0.05)
    Dim tolIJK  As Double:  tolIJK  = CfgDbl("TOL_IJK",  0.001)
    Dim useIJK  As Boolean: useIJK  = CfgBool("USE_IJK",  False)
    Dim outPfx  As String:  outPfx  = CfgStr("OUT_PREFIX"): If Len(outPfx)=0 Then outPfx="CMP_"

    Dim cMatch As Long: cMatch  = PaletteColor("MATCH",         RGB(198,239,206))
    Dim cName  As Long: cName   = PaletteColor("NAME_CHANGED",  RGB(255,235,156))
    Dim cCoord As Long: cCoord  = PaletteColor("COORD_CHANGED", RGB(255,215,215))
    Dim cDel   As Long: cDel    = PaletteColor("DELETED",       RGB(244,204,204))
    Dim cAdd   As Long: cAdd    = PaletteColor("ADDED",         RGB(217,234,211))
    Dim cDiff  As Long: cDiff   = PaletteColor("DIFF_CELL",     RGB(255,102,0))
    Dim cHdrBg As Long: cHdrBg  = PaletteColor("HEADER_BG",     RGB(68,114,196))
    Dim cHdrFg As Long: cHdrFg  = PaletteColor("HEADER_FG",     RGB(255,255,255))

    Dim wb As Workbook: Set wb = ThisWorkbook
    If GetConfigSheet(wb) Is Nothing Then MsgBox "Config sheet not found. Expected '⚙ Settings' or another sheet starting with '⚙'.", vbCritical: GoTo Cleanup
    If Not SheetExists(wb, sh1) Then MsgBox "Sheet '" & sh1 & "' was not found.", vbCritical: GoTo Cleanup
    If Not SheetExists(wb, sh2) Then MsgBox "Sheet '" & sh2 & "' was not found.", vbCritical: GoTo Cleanup

    Dim ws1 As Worksheet: Set ws1 = wb.Sheets(sh1)
    Dim ws2 As Worksheet: Set ws2 = wb.Sheets(sh2)

    ' Column indices: exact headers only
    Dim ci1_name As Long: ci1_name = ColIdx(ws1, hdrRow, "Name")
    Dim ci1_x As Long:    ci1_x    = ColIdx(ws1, hdrRow, "X")
    Dim ci1_y As Long:    ci1_y    = ColIdx(ws1, hdrRow, "Y")
    Dim ci1_z As Long:    ci1_z    = ColIdx(ws1, hdrRow, "Z")
    Dim ci1_i As Long:    ci1_i    = ColIdx(ws1, hdrRow, "I")
    Dim ci1_j As Long:    ci1_j    = ColIdx(ws1, hdrRow, "J")
    Dim ci1_k As Long:    ci1_k    = ColIdx(ws1, hdrRow, "K")

    Dim ci2_name As Long: ci2_name = ColIdx(ws2, hdrRow, "Name")
    Dim ci2_x As Long:    ci2_x    = ColIdx(ws2, hdrRow, "X")
    Dim ci2_y As Long:    ci2_y    = ColIdx(ws2, hdrRow, "Y")
    Dim ci2_z As Long:    ci2_z    = ColIdx(ws2, hdrRow, "Z")
    Dim ci2_i As Long:    ci2_i    = ColIdx(ws2, hdrRow, "I")
    Dim ci2_j As Long:    ci2_j    = ColIdx(ws2, hdrRow, "J")
    Dim ci2_k As Long:    ci2_k    = ColIdx(ws2, hdrRow, "K")

    Dim missing1 As String
    missing1 = MissingColsMsg(ci1_name, ci1_x, ci1_y, ci1_z, ci1_i, ci1_j, ci1_k)
    If Len(missing1) > 0 Then
        MsgBox "Required columns are missing on sheet '" & sh1 & "': " & missing1 & vbCrLf & _
               "Expected exact names: Name, X, Y, Z, I, J, K. Fix the headers and run again.", _
               vbExclamation, "Missing Required Columns"
        GoTo Cleanup
    End If

    Dim missing2 As String
    missing2 = MissingColsMsg(ci2_name, ci2_x, ci2_y, ci2_z, ci2_i, ci2_j, ci2_k)
    If Len(missing2) > 0 Then
        MsgBox "Required columns are missing on sheet '" & sh2 & "': " & missing2 & vbCrLf & _
               "Expected exact names: Name, X, Y, Z, I, J, K. Fix the headers and run again.", _
               vbExclamation, "Missing Required Columns"
        GoTo Cleanup
    End If

    Dim n1 As Long: n1 = ws1.Cells(ws1.Rows.Count,ci1_name).End(xlUp).Row - hdrRow
    Dim n2 As Long: n2 = ws2.Cells(ws2.Rows.Count,ci2_name).End(xlUp).Row - hdrRow
    If n1<=0 Or n2<=0 Then MsgBox "No data found.", vbExclamation: GoTo Cleanup

    ReDim d1(1 To n1, 1 To 7) As Variant
    ReDim d2(1 To n2, 1 To 7) As Variant

    Dim i As Long
    For i = 1 To n1
        Dim rr As Long: rr = hdrRow + i
        d1(i,1)=Trim(CStr(Nz2(ws1.Cells(rr,ci1_name).Value,"")))
        d1(i,2)=SafeDbl(ws1.Cells(rr,ci1_x).Value)
        d1(i,3)=SafeDbl(ws1.Cells(rr,ci1_y).Value)
        d1(i,4)=SafeDbl(ws1.Cells(rr,ci1_z).Value)
        d1(i,5)=IIf(ci1_i>0,SafeDbl(ws1.Cells(rr,ci1_i).Value),Empty)
        d1(i,6)=IIf(ci1_j>0,SafeDbl(ws1.Cells(rr,ci1_j).Value),Empty)
        d1(i,7)=IIf(ci1_k>0,SafeDbl(ws1.Cells(rr,ci1_k).Value),Empty)
    Next i
    For i = 1 To n2
        Dim rr2 As Long: rr2 = hdrRow + i
        d2(i,1)=Trim(CStr(Nz2(ws2.Cells(rr2,ci2_name).Value,"")))
        d2(i,2)=SafeDbl(ws2.Cells(rr2,ci2_x).Value)
        d2(i,3)=SafeDbl(ws2.Cells(rr2,ci2_y).Value)
        d2(i,4)=SafeDbl(ws2.Cells(rr2,ci2_z).Value)
        d2(i,5)=IIf(ci2_i>0,SafeDbl(ws2.Cells(rr2,ci2_i).Value),Empty)
        d2(i,6)=IIf(ci2_j>0,SafeDbl(ws2.Cells(rr2,ci2_j).Value),Empty)
        d2(i,7)=IIf(ci2_k>0,SafeDbl(ws2.Cells(rr2,ci2_k).Value),Empty)
    Next i

    Dim res() As Variant
    Dim nRes As Long
    RunCompare d1, n1, d2, n2, tolXYZ, tolIJK, useIJK, res, nRes

    Dim sfx(6) As String
    sfx(0)="Overview": sfx(1)="All Results": sfx(2)="Match"
    sfx(3)="Name Changed": sfx(4)="Coord Changed": sfx(5)="Deleted": sfx(6)="Added"
    Dim s As Long
    For s=0 To 6: DelSheet wb, outPfx & sfx(s): Next s

    WriteOverview  wb, res, nRes, n1, n2, sh1, sh2, tolXYZ, outPfx, cMatch, cName, cCoord, cDel, cAdd, cHdrBg, cHdrFg
    WriteResults   wb, outPfx&"All Results",    res, nRes, "",             cMatch,cName,cCoord,cDel,cAdd,cDiff,cHdrBg,cHdrFg
    WriteResults   wb, outPfx&"Match",          res, nRes, "MATCH",        cMatch,cName,cCoord,cDel,cAdd,cDiff,cHdrBg,cHdrFg
    WriteResults   wb, outPfx&"Name Changed",   res, nRes, "NAME_CHANGED", cMatch,cName,cCoord,cDel,cAdd,cDiff,cHdrBg,cHdrFg
    WriteResults   wb, outPfx&"Coord Changed",  res, nRes, "COORD_CHANGED",cMatch,cName,cCoord,cDel,cAdd,cDiff,cHdrBg,cHdrFg
    WriteResults   wb, outPfx&"Deleted",        res, nRes, "DELETED",      cMatch,cName,cCoord,cDel,cAdd,cDiff,cHdrBg,cHdrFg
    WriteResults   wb, outPfx&"Added",          res, nRes, "ADDED",        cMatch,cName,cCoord,cDel,cAdd,cDiff,cHdrBg,cHdrFg

    wb.Sheets(outPfx & "Overview").Activate
    Dim elapsed As Double: elapsed = Round(Timer-t0,1)
    MsgBox "Done. " & nRes & " rows processed in " & elapsed & " sec.", vbInformation, "Comparison Complete"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'===========================================================
' COMPARISON ENGINE
'===========================================================
Sub RunCompare(d1() As Variant, n1 As Long, d2() As Variant, n2 As Long, _
               tolXYZ As Double, tolIJK As Double, useIJK As Boolean, _
               ByRef res() As Variant, ByRef nRes As Long)
    ReDim res(1 To n1+n2, 1 To R_COLS)
    nRes = 0
    Dim dName  As Object: Set dName  = CreateObject("Scripting.Dictionary")
    Dim dCoord As Object: Set dCoord = CreateObject("Scripting.Dictionary")
    ReDim matched2(1 To n2) As Boolean

    Dim i As Long
    For i = 1 To n2
        Dim nm As String: nm = CStr(d2(i,1))
        If Len(nm)>0 And Not dName.Exists(nm) Then dName(nm)=i
        Dim ck As String
        ck = CKStr(d2(i,2),d2(i,3),d2(i,4),d2(i,5),d2(i,6),d2(i,7),tolXYZ,tolIJK,useIJK)
        If Not dCoord.Exists(ck) Then dCoord(ck)=i
    Next i

    For i = 1 To n1
        Dim name1 As String: name1 = CStr(d1(i,1))
        Dim ck1 As String
        ck1 = CKStr(d1(i,2),d1(i,3),d1(i,4),d1(i,5),d1(i,6),d1(i,7),tolXYZ,tolIJK,useIJK)
        Dim byName As Boolean:  byName  = dName.Exists(name1)
        Dim byCoord As Boolean: byCoord = dCoord.Exists(ck1)
        nRes=nRes+1: CopyF1 d1,i,res,nRes
        If byName Then
            Dim j As Long: j=dName(name1): matched2(j)=True: CopyF2 d2,j,res,nRes
            Dim ck2 As String
            ck2=CKStr(d2(j,2),d2(j,3),d2(j,4),d2(j,5),d2(j,6),d2(j,7),tolXYZ,tolIJK,useIJK)
            If ck1=ck2 Then
                res(nRes,R_STATUS)="MATCH"
            Else
                res(nRes,R_STATUS)="COORD_CHANGED"
                CalcDiffs d1,i,d2,j,res,nRes,tolXYZ,tolIJK
            End If
        ElseIf byCoord Then
            Dim j2 As Long: j2=dCoord(ck1)
            If Not matched2(j2) Then
                matched2(j2)=True: CopyF2 d2,j2,res,nRes
                res(nRes,R_STATUS)="NAME_CHANGED"
                res(nRes,R_NAME_DIFF)=d1(i,1) & " -> " & d2(j2,1)
                res(nRes,R_DIFF_FLD)="Name"
                res(nRes,R_DX)=0:res(nRes,R_DY)=0:res(nRes,R_DZ)=0
            Else
                res(nRes,R_STATUS)="DELETED"
            End If
        Else
            res(nRes,R_STATUS)="DELETED"
        End If
    Next i
    Dim k As Long
    For k=1 To n2
        If Not matched2(k) Then
            nRes=nRes+1: res(nRes,R_STATUS)="ADDED": CopyF2 d2,k,res,nRes
        End If
    Next k
End Sub

Function CKStr(x,y,z,ii,jj,kk,tolXYZ As Double,tolIJK As Double,useIJK As Boolean) As String
    Dim s As String
    s=RK(x,tolXYZ)&"|"&RK(y,tolXYZ)&"|"&RK(z,tolXYZ)
    If useIJK Then s=s&"|"&RK(ii,tolIJK)&"|"&RK(jj,tolIJK)&"|"&RK(kk,tolIJK)
    CKStr=s
End Function

Function RK(v,tol As Double) As String
    If IsEmpty(v) Or IsNull(v) Or IsError(v) Then RK="?": Exit Function
    On Error Resume Next
    RK=CStr(CLng(CDbl(v)/tol))
    If Err.Number<>0 Then RK="?": Err.Clear
    On Error GoTo 0
End Function

Sub CalcDiffs(d1() As Variant,i1 As Long,d2() As Variant,i2 As Long, _
              res() As Variant,nr As Long,tolXYZ As Double,tolIJK As Double)
    Dim diffs As String: diffs=""
    Dim fXYZ(2) As String: fXYZ(0)="X": fXYZ(1)="Y": fXYZ(2)="Z"
    Dim cXYZ(2) As Long:   cXYZ(0)=2: cXYZ(1)=3: cXYZ(2)=4
    Dim rXYZ(2) As Long:   rXYZ(0)=R_DX: rXYZ(1)=R_DY: rXYZ(2)=R_DZ
    Dim f As Long
    For f=0 To 2
        Dim dv As Double: dv=SafeDbl(d2(i2,cXYZ(f)))-SafeDbl(d1(i1,cXYZ(f)))
        res(nr,rXYZ(f))=Round(dv,4)
        If Abs(dv)>tolXYZ Then diffs=diffs&IIf(Len(diffs)>0,", ","")&fXYZ(f)
    Next f
    Dim fIJK(2) As String: fIJK(0)="I": fIJK(1)="J": fIJK(2)="K"
    Dim cIJK(2) As Long:   cIJK(0)=5: cIJK(1)=6: cIJK(2)=7
    Dim rIJK(2) As Long:   rIJK(0)=R_DI: rIJK(1)=R_DJ: rIJK(2)=R_DK
    For f=0 To 2
        Dim dv2 As Double: dv2=SafeDbl(d2(i2,cIJK(f)))-SafeDbl(d1(i1,cIJK(f)))
        res(nr,rIJK(f))=Round(dv2,6)
        If Abs(dv2)>tolIJK Then diffs=diffs&IIf(Len(diffs)>0,", ","")&fIJK(f)
    Next f
    res(nr,R_DIFF_FLD)=diffs
End Sub

Function MissingColsMsg(ciName As Long, ciX As Long, ciY As Long, ciZ As Long, _
                        ciI As Long, ciJ As Long, ciK As Long) As String
    Dim msg As String: msg = ""
    If ciName = 0 Then msg = AppendMissing(msg, "Name")
    If ciX = 0 Then msg = AppendMissing(msg, "X")
    If ciY = 0 Then msg = AppendMissing(msg, "Y")
    If ciZ = 0 Then msg = AppendMissing(msg, "Z")
    If ciI = 0 Then msg = AppendMissing(msg, "I")
    If ciJ = 0 Then msg = AppendMissing(msg, "J")
    If ciK = 0 Then msg = AppendMissing(msg, "K")
    MissingColsMsg = msg
End Function

Function AppendMissing(msg As String, item As String) As String
    If Len(msg) = 0 Then
        AppendMissing = item
    Else
        AppendMissing = msg & ", " & item
    End If
End Function

Sub CopyF1(d() As Variant,idx As Long,res() As Variant,nr As Long)
    res(nr,R_F1_NAME)=d(idx,1):res(nr,R_F1_X)=d(idx,2):res(nr,R_F1_Y)=d(idx,3)
    res(nr,R_F1_Z)=d(idx,4):res(nr,R_F1_I)=d(idx,5):res(nr,R_F1_J)=d(idx,6):res(nr,R_F1_K)=d(idx,7)
End Sub

Sub CopyF2(d() As Variant,idx As Long,res() As Variant,nr As Long)
    res(nr,R_F2_NAME)=d(idx,1):res(nr,R_F2_X)=d(idx,2):res(nr,R_F2_Y)=d(idx,3)
    res(nr,R_F2_Z)=d(idx,4):res(nr,R_F2_I)=d(idx,5):res(nr,R_F2_J)=d(idx,6):res(nr,R_F2_K)=d(idx,7)
End Sub

'===========================================================
' OUTPUT WRITERS
'===========================================================
Function GetBgColor(st As String,cM,cN,cC,cD,cA As Long) As Long
    Select Case st
        Case "MATCH":         GetBgColor=cM
        Case "NAME_CHANGED":  GetBgColor=cN
        Case "COORD_CHANGED": GetBgColor=cC
        Case "DELETED":       GetBgColor=cD
        Case "ADDED":         GetBgColor=cA
        Case Else:            GetBgColor=RGB(240,240,240)
    End Select
End Function

Sub WriteResults(wb As Workbook, shName As String, _
                 res() As Variant, nRes As Long, filterSt As String, _
                 cM,cN,cC,cD,cA,cDiff,cHBg,cHFg As Long)
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = shName

    Dim hdrs As Variant
    hdrs=Array("Status","F1_Name","F1_X","F1_Y","F1_Z","F1_I","F1_J","F1_K", _
               "F2_Name","F2_X","F2_Y","F2_Z","F2_I","F2_J","F2_K", _
               "NAME_diff","dX","dY","dZ","dI","dJ","dK","DIFF_Fields")

    Dim c As Long
    For c=0 To UBound(hdrs)
        With ws.Cells(1,c+1)
            .Value=hdrs(c):.Font.Name="Calibri":.Font.Bold=True
            .Font.Color=cHFg:.Interior.Color=cHBg
            .HorizontalAlignment=xlCenter:.VerticalAlignment=xlCenter
            .Borders.LineStyle=xlContinuous:.Borders.Color=RGB(180,180,180)
        End With
    Next c
    ws.Rows(1).RowHeight=28

    Dim isD(22) As Boolean
    isD(15)=True:isD(16)=True:isD(17)=True:isD(18)=True
    isD(19)=True:isD(20)=True:isD(21)=True:isD(22)=True

    Dim outR As Long: outR=2
    Dim i As Long
    For i=1 To nRes
        Dim st As String: st=CStr(res(i,R_STATUS))
        If filterSt<>"" And st<>filterSt Then GoTo Skip
        Dim bgC As Long: bgC=GetBgColor(st,cM,cN,cC,cD,cA)
        For c=0 To UBound(hdrs)
            Dim v As Variant: v=res(i,c+1)
            If IsEmpty(v) Or IsNull(v) Then v=""
            With ws.Cells(outR,c+1)
                .Value=v:.Font.Name="Calibri":.Font.Size=10
                .Interior.Color=bgC
                .HorizontalAlignment=xlLeft:.VerticalAlignment=xlCenter
                .Borders.LineStyle=xlContinuous:.Borders.Color=RGB(200,200,200)
                If isD(c) Then
                    Dim sig As Boolean: sig=False
                    If (c=15 Or c=22) And Len(CStr(v))>0 Then sig=True
                    If c>=16 And c<=21 Then
                        On Error Resume Next
                        If Abs(CDbl(v))>0.000001 Then sig=True
                        On Error GoTo 0
                    End If
                    If sig Then .Font.Bold=True:.Font.Color=cDiff
                End If
            End With
        Next c
        ws.Rows(outR).RowHeight=18
        outR=outR+1
Skip:
    Next i

    ws.Columns.AutoFit
    Dim col As Range
    For Each col In ws.UsedRange.Columns
        If col.ColumnWidth>45 Then col.ColumnWidth=45
        If col.ColumnWidth<8  Then col.ColumnWidth=8
    Next col
    ws.Rows(1).AutoFilter
    ws.Cells(2,1).Select
    ActiveWindow.FreezePanes=True
End Sub

Sub WriteOverview(wb As Workbook,res() As Variant,nRes As Long, _
                  n1 As Long,n2 As Long,sh1 As String,sh2 As String, _
                  tolXYZ As Double,outPfx As String, _
                  cM,cN,cC,cD,cA,cHBg,cHFg As Long)
    Dim ws As Worksheet
    Set ws=wb.Sheets.Add(Before:=wb.Sheets(1))
    ws.Name=outPfx&"Overview"
    ws.Columns("A").ColumnWidth=4:ws.Columns("B").ColumnWidth=28
    ws.Columns("C").ColumnWidth=14:ws.Columns("D").ColumnWidth=10:ws.Columns("E").ColumnWidth=40

    ws.Range("B2:E2").Merge
    With ws.Range("B2")
        .Value="Points Comparison Report":.Font.Name="Calibri":.Font.Size=16:.Font.Bold=True:.Font.Color=RGB(31,56,100)
    End With
    ws.Rows(2).RowHeight=28
    ws.Range("B3:E3").Merge
    ws.Range("B3").Value="Sheet 1: "&sh1&"   |   Sheet 2: "&sh2&"   |   "&Format(Now,"YYYY-MM-DD HH:MM")
    ws.Range("B3").Font.Italic=True:ws.Range("B3").Font.Color=RGB(130,130,130)
    ws.Range("B4:E4").Merge
    ws.Range("B4").Value="Rows in sheet 1: "&n1&"   |   Rows in sheet 2: "&n2&"   |   XYZ tolerance: "&tolXYZ&" mm"
    ws.Range("B4").Font.Italic=True:ws.Range("B4").Font.Color=RGB(160,160,160)

    Dim hRow As Long: hRow=6
    Dim hH As Variant: hH=Array("Status","Count","Color","Description")
    Dim c As Long
    For c=0 To 3
        With ws.Cells(hRow,c+2)
            .Value=hH(c):.Font.Bold=True:.Font.Color=cHFg:.Interior.Color=cHBg
            .HorizontalAlignment=xlCenter:.Borders.LineStyle=xlContinuous
        End With
    Next c
    ws.Rows(hRow).RowHeight=24

    Dim codes(4) As String
    codes(0)="MATCH":codes(1)="NAME_CHANGED":codes(2)="COORD_CHANGED":codes(3)="DELETED":codes(4)="ADDED"
    Dim lbls(4) As String
    lbls(0)="✔ Match":lbls(1)="✎ Name Changed"
    lbls(2)="⚠ Coordinates Changed":lbls(3)="✖ Deleted":lbls(4)="＋ Added"

    Dim r As Long
    For r=0 To 4
        Dim cnt As Long: cnt=0: Dim i As Long
        For i=1 To nRes
            If CStr(res(i,1))=codes(r) Then cnt=cnt+1
        Next i
        Dim row As Long: row=hRow+1+r
        ws.Rows(row).RowHeight=20
        Dim bg As Long: bg=GetBgColor(codes(r),cM,cN,cC,cD,cA)
        With ws.Cells(row,2): .Value=codes(r):.Interior.Color=bg:.Font.Bold=True:.Borders.LineStyle=xlContinuous:.HorizontalAlignment=xlCenter: End With
        With ws.Cells(row,3): .Value=cnt:.Interior.Color=bg:.Font.Bold=True:.Borders.LineStyle=xlContinuous:.HorizontalAlignment=xlCenter: End With
        With ws.Cells(row,4): .Interior.Color=bg:.Borders.LineStyle=xlContinuous: End With
        With ws.Cells(row,5): .Value=lbls(r):.Interior.Color=bg:.Borders.LineStyle=xlContinuous: End With
    Next r

    Dim tr As Long: tr=hRow+6: ws.Rows(tr).RowHeight=20
    With ws.Cells(tr,2): .Value="TOTAL":.Font.Bold=True:.Borders.LineStyle=xlContinuous:.HorizontalAlignment=xlCenter: End With
    With ws.Cells(tr,3): .Value=nRes:.Font.Bold=True:.Borders.LineStyle=xlContinuous:.HorizontalAlignment=xlCenter: End With
End Sub

'===========================================================
' UTILITIES
'===========================================================
Function GetConfigSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = "⚙ Settings" Then
            Set GetConfigSheet = ws
            Exit Function
        End If
    Next ws
    For Each ws In wb.Worksheets
        If Left$(ws.Name, 1) = "⚙" Then
            Set GetConfigSheet = ws
            Exit Function
        End If
    Next ws
    Set GetConfigSheet = Nothing
End Function

Function SheetExists(wb As Workbook, nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next: Set ws=wb.Sheets(nm): On Error GoTo 0
    SheetExists=Not ws Is Nothing
End Function

Sub DelSheet(wb As Workbook, nm As String)
    If SheetExists(wb,nm) Then
        Application.DisplayAlerts=False: wb.Sheets(nm).Delete: Application.DisplayAlerts=True
    End If
End Sub

Function SafeDbl(v As Variant) As Double
    On Error Resume Next: SafeDbl=CDbl(v)
    If Err.Number<>0 Then SafeDbl=0: Err.Clear
    On Error GoTo 0
End Function

Function Nz2(v As Variant, def As Variant) As Variant
    If IsNull(v) Or IsEmpty(v) Then Nz2=def Else Nz2=v
End Function
