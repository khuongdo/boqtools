Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports System.Text.RegularExpressions

Public Class BOQBeamTools
    Public acApp As AcadApplication
    Public acDoc As AcadDocument
    Public xlApp As Application
    Public xlWB As Workbook
    Public xlWS As Worksheet
    Public Const APPNAME = "BOQTools(c)DHK"
    Public dbl_Lnhip As Double ' CHIEU DAI THEP DAM KHONG TRU COT
    Public dbl_b As Double ' BE RONG DAM
    Public dbl_h As Double ' CHIEU CAO DAM
    Public str_loaithep As String ' KIEU THEP NE-NE TD-TD
    Public str_tenthep As String
    Public TextHeight As Double 'CHIEU CAO TEXT BEAM 
    Public dbl_Wcot As Double 'BE RONG COT
    Public int_sonhip As Integer 'SO NHIP DAM
    Public objHandle As String ' handle cua text dam
    
    ' THU TUC GOI EXCEL VA ACAD
    Sub StartACAD()
        Try
            acApp = GetObject(, "Autocad.Application")
            acDoc = acApp.ActiveDocument
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub StartExcel()
        Try
            xlApp = GetObject(, "Excel.Application")
            xlWB = xlApp.ActiveWorkbook
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
    End Sub
    ' XU LY TRONG EXCEL
    Sub AddBeamLV1()
        Dim NextRow As Long
        Dim temptxt As String
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        xlWS = xlWB.Sheets("dulieu_insert")
        temptxt = InputBox("Nhap ten: ", "khuong.do")
        If Not temptxt = "" Then
            xlWS.Range("A3").Value = temptxt
            xlWS.Range("A48").Value = temptxt
        Else : Exit Sub
        End If
        'Chen vao sheet Rein
        xlWS = xlApp.Sheets("Reinforcement")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("3:4").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        'Chen vao sheet Concrete
        xlWS = xlApp.Sheets("Concrete")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("48:49").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        'Chen vao sheet formwork
        xlWS = xlApp.Sheets("Formwork")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("48:49").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        xlWB.Sheets("Reinforcement").Activate()
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub
    Sub AddBeamLV2()
        Dim TextLv1 As String
        Dim Cell, WorkRange As Range
        Dim StartRow, EndRow As Long
        Dim temptxt As String
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        temptxt = InputBox("Nhập tên LV2: ", APPNAME)
        If Not temptxt = "" Then
            xlWB.Sheets("dulieu_insert").Range("A6").Value = temptxt
            xlWB.Sheets("dulieu_insert").Range("A50").Value = temptxt
        Else : Exit Sub
        End If
        'lay du lieu dam LV1 hien tai tu sheet LIST------------------
        xlWS = xlWB.Sheets("List")
        TextLv1 = xlWS.Range("XFB1").Text
        'Tim vi tri ket thuc Beam LV1 sheet Reinforcement------------------
        xlWS = xlWB.Sheets("Reinforcement")
        xlWS.Activate()
        WorkRange = xlApp.Intersect(xlWS.Columns("P"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas) ' column chua lv1
        For Each Cell In WorkRange
            If Cell.Text <> "" Then
                If Cell.Text = "/" & TextLv1 Then
                    EndRow = Cell.Row
                    Exit For
                End If
            End If
        Next
        xlWB.Sheets("dulieu_insert").Rows("6:7").Copy() ' Chen du lieu vao sheet Rein
        xlWS.Range("A" & EndRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        'Tim vi tri bat dau va ket thuc Beam LV1 sheet Concrete----------
        xlWS = xlWB.Sheets("Concrete")
        xlWS.Activate()
        WorkRange = xlApp.Intersect(xlWS.Columns("M"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas) ' column chua lv1
        For Each Cell In WorkRange
            Select Case Cell.Text
                Case TextLv1
                    StartRow = Cell.Row
                Case "/" & TextLv1
                    EndRow = Cell.Row
                    Exit For
            End Select
        Next Cell
        ' Chen beam LV2 vao sheet concrete
        xlWB.Sheets("dulieu_insert").Rows("50:51").Copy()
        xlWS.Range("A" & EndRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        '---------------------------
        'Tim vi tri bat dau va ket thuc Beam LV1 sheet Formwork----------
        xlWS = xlWB.Sheets("Formwork")
        xlWS.Activate()
        WorkRange = xlApp.Intersect(xlWS.Columns("M"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas)

        For Each Cell In WorkRange
            Select Case Cell.Value
                Case TextLv1
                    StartRow = Cell.Row
                Case "/" & TextLv1
                    EndRow = Cell.Row
                    Exit For
            End Select
        Next Cell

        ' Chen beam LV3 vao sheet Formwork
        xlWB.Sheets("dulieu_insert").Rows("50:51").Copy()
        xlWS.Range("A" & EndRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        '---------------------------
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub
    Sub CalculateBeamLV3_1(str_BeamLV2 As String)
        Dim InsertRow, CurrRow As Long
        On Error Resume Next
        Call StartExcel()
        CurrRow = FindNextBeam() - 1 ' Vi thu tuc nay goi sau khi da add du lieu vao sheet beamdata nen phai -1
        With xlWB.Sheets("dulieu_insert") ' Ghi ten dam vao sheet dulieu_insert truoc khi insert 
            .Range("A9") = "BEAM" & CurrRow
            .Range("A21") = "BEAM" & CurrRow
            .Range("A23") = "BEAM" & CurrRow
        End With
        ' Chen dam LV3 vao sheet Concrete
        xlWS = xlWB.Sheets("Concrete")
        InsertRow = FindInsertRow("Concrete", "M", str_BeamLV2)
        xlWB.Sheets("dulieu_insert").Rows(21).Copy()
        xlWS.Activate()
        xlWS.Range("A" & InsertRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        ' Chen dam LV3 vao sheet Formwork
        xlWS = xlWB.Sheets("Formwork")
        InsertRow = FindInsertRow("Formwork", "M", str_BeamLV2)
        xlWB.Sheets("dulieu_insert").Rows(23).Copy()
        xlWS.Activate()
        xlWS.Range("A" & InsertRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        ' Chen dam LV3 vao sheet Reinforcement
        xlWS = xlWB.Sheets("Reinforcement")
        InsertRow = FindInsertRow("Reinforcement", "P", str_BeamLV2)
        xlWB.Sheets("dulieu_insert").Rows("9:18").Copy()
        xlWS.Activate()
        xlWS.Range("A" & InsertRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
    End Sub
    ' XU LY ADD BEAM KIEU 1 TRONG AUTOCAD
    Sub AddBeamLV3FromACAD_1(str_BeamlV2 As String, SlabThk As String)
        On Error Resume Next
        Dim sset(0 To 10) As AcadSelectionSet
        Dim str_beamname As String
        Dim acText As AcadEntity
        Dim int_sck As Integer
        Dim xlws_insert As Worksheet
        Dim str_ttren(1), str_tduoi(1), str_tdai(0 To 1), str_tcgoi, str_tcnhip, str_tgia As String
        Dim Sdai1, Sdai2 As Double
        Dim dbl_LnhipTB, dbl_WcotTB As Double
        Dim acTextHandle As String
        Dim SLtheptren, SLthepduoi As Integer
        Call StartACAD()
        Call StartExcel()
        xlws_insert = xlWB.Sheets("dulieu_insert")
        Dim i As Integer
        ' XOA HET DU LIEU TREN SHEET ADDBEAMLV3
        xlWS = xlWB.Sheets("addbeamlv3")
        xlWS.UsedRange.EntireRow.Delete(Shift:=XlDirection.xlUp)
        For i = 0 To 10
            sset(i) = acDoc.SelectionSets.Add("Acadfile1" & i)
        Next
        With acDoc.Utility
            ' Chon ten dam
            .Prompt(vbCrLf & "Chọn tên dầm: ")
            sset(0).SelectOnScreen()
            acText = sset(0).Item(0)
            str_beamname = acText.Textstring
            str_beamname = RemoveLeftString(str_beamname, "%%U")
            TextHeight = acText.Height
            acText.color = ACAD_COLOR.acRed
            acTextHandle = acText.Handle

            ' chon BxH
            Call ChonBH()
            ' chon thep chu tren
            .Prompt("Chọn thép chủ trên: ")
            sset(2).SelectOnScreen()
            SLtheptren = 0
            str_ttren(0) = sset(2).Item(0).TextString
            str_ttren(1) = sset(2).Item(1).TextString
            If str_ttren(1) <> "" Then
                SLtheptren = CType(SplitText(str_ttren(0), 1), Integer) + CType(SplitText(str_ttren(1), 1), Integer)
            Else
                SLtheptren = CType(SplitText(str_ttren(0), 1), Integer)
            End If
            ' Chon thep chu lop duoi
            .Prompt("Chọn thép chủ dưới: ")
            sset(3).SelectOnScreen()
            SLthepduoi = 0
            str_tduoi(0) = sset(3).Item(0).TextString
            str_tduoi(1) = sset(3).Item(1).TextString
            If str_tduoi(1) <> "" Then
                SLthepduoi = CType(SplitText(str_tduoi(0), 1), Integer) + CType(SplitText(str_tduoi(1), 1), Integer)
            Else
                SLthepduoi = CType(SplitText(str_tduoi(0), 1), Integer)
            End If
            ' Chon thep dai
            .Prompt("Chọn thép đai: ")
            sset(4).SelectOnScreen()
            str_tdai(0) = sset(4).Item(0).TextString
            str_tdai(1) = sset(4).Item(1).TextString
            Sdai1 = 0
            Sdai2 = 0
            Sdai1 = SplitText(str_tdai(0), 2)
            Sdai2 = SplitText(str_tdai(1), 2)
        End With
        ' 2. ADD TRUOC THONG TIN DAM VA THEP CHU & THEP DAI VAO SHEET ADDBEAMLV3
        xlws_insert.Rows("26:30").Copy()
        xlWS = xlWB.Sheets("addbeamlv3")
        xlWS.Activate()
        xlWS.Range("A1").Insert(XlDirection.xlDown)
        xlApp.CutCopyMode = False
        ' NHAP THEP TC GOI, NHIP,GIA
        With acDoc.Utility
            ' Chon thep tang cuong goi
            .Prompt("Chọn thép TC gối [ESC để bỏ qua]: ")
            sset(5).SelectOnScreen()
            str_tcgoi = ""
            str_tcgoi = sset(5).Item(0).TextString
            ' Chon thep tang cuong nhịp
            .Prompt("Chọn thép tăng cường nhịp [ESC để bỏ qua]: ")
            sset(6).SelectOnScreen()
            str_tcnhip = ""
            str_tcnhip = sset(6).Item(0).TextString
            ' Chon thep gia
            .Prompt("Chọn thép giá [ESC để bỏ qua]: ")
            sset(7).SelectOnScreen()
            str_tgia = sset(7).Item(0).TextString
        End With

        ' NHAP SO CK, SO NHIP
        With acDoc.Utility
            ' Nhap so ck
            int_sck = 1
            int_sck = CType(.GetInteger("Nhập số CK: "), Integer)
            '' TIM TEXT
            Call FindtextInAcad()
            ' Nhap so nhip
            int_sonhip = 1
            int_sonhip = CType(.GetInteger("Nhập số nhịp: "), Integer)
        End With
        ' VE LINE DAM
        dbl_Wcot = 0
        Call VeLdam()
        ' TIEP TUC ADD DU LIEU VAO SHEET ADDBEAMLV3
        dbl_LnhipTB = (dbl_Lnhip - dbl_Wcot) / int_sonhip
        dbl_WcotTB = dbl_Wcot / (int_sonhip - 1)
        xlWS = xlWB.Sheets("addbeamlv3")
        xlws_insert.Range("A32:A37").Clear()
        Dim R As Long
        Dim LastRow As Long
        ' NHAP 2 THEP TANG CUONG GOI 
        If Not str_tcgoi = "" Then
            For i = 1 To 2
                R = 32 ' ne-tu
                LastRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                xlws_insert.Range("F" & R).Value = SplitText(str_tcgoi, 2) ' Fi
                xlws_insert.Range("G" & R).Value = SplitText(str_tcgoi, 1) ' SL
                xlws_insert.Range("C" & R).FormulaR1C1 = "=Goi_ThepTren*" & Math.Round(dbl_LnhipTB, 0)
                xlws_insert.Rows(R).Copy()
                xlWS.Activate()
                LastRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                xlWS.Range("A" & LastRow).Insert()
                xlApp.CutCopyMode = False
            Next
            If int_sonhip > 1 Then ' 
                R = 33 ' tu-tu
                For i = 1 To int_sonhip - 1
                    LastRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                    xlws_insert.Range("F" & R).Value = SplitText(str_tcgoi, 2) ' Fi
                    xlws_insert.Range("G" & R).Value = SplitText(str_tcgoi, 1) ' SL
                    xlws_insert.Range("C" & R).FormulaR1C1 = "=Goi_ThepTren*" & Math.Round(dbl_LnhipTB, 0) & "+" & Math.Round(dbl_WcotTB, 0) & "+Goi_ThepTren*" & Math.Round(dbl_LnhipTB, 0)
                    xlws_insert.Rows(R).Copy()
                    xlWS.Activate()
                    xlWS.Range("A" & LastRow).Insert()
                    xlApp.CutCopyMode = False
                Next
            End If
        End If
        ' NHAP 2 THEP TANG CUONG NHIP NEU CO
        If Not str_tcnhip = "" Then
            R = 36 ' tu-tu
            For i = 1 To int_sonhip
                LastRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                xlws_insert.Range("F" & R).Value = SplitText(str_tcnhip, 2) ' Fi
                xlws_insert.Range("G" & R).Value = SplitText(str_tcnhip, 1) ' SL
                xlws_insert.Range("C" & R).FormulaR1C1 = "=(1-2*Goi_ThepDuoi)*" & Math.Round(dbl_LnhipTB, 0)
                xlws_insert.Rows(R).Copy()
                xlWS.Activate()
                xlWS.Range("A" & LastRow).Insert()
                xlApp.CutCopyMode = False
            Next
        End If
        ' Zoom toi doi tuong ten dam
        Dim p1, p2
        acText = acDoc.HandleToObject(acTextHandle)
        acText.GetBoundingBox(p1, p2)
        acApp.ZoomWindow(p1, p2)
        Call ThepC_1()
        ' CHEN HANG KET THUC BEAM
        xlws_insert.Rows("38").copy()
        LastRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlWS.Range("A" & LastRow).Insert()
        xlApp.CutCopyMode = False
        ' CO NHAP DAM NAY HAY KHONG
        Dim opt As String
        opt = "Yes"
        acDoc.Utility.InitializeUserInput(0, "Yes No")
        opt = acDoc.Utility.GetKeyword("Nhập dầm này Yes[No]: ")
        Select Case opt
            Case "No"
                GoTo Exit_Sub
        End Select
        ' XUAT DU LIEU SANG SHEET BEAM DATA
        xlWS = xlWB.Sheets("beamdata")
        Dim NextBeam As Long = FindCurrBeam() + 1
        xlWS.Range("A" & NextBeam).Value = "BEAM" & NextBeam ' CODE
        xlWS.Range("B" & NextBeam).Value = str_beamname ' TEN DAM
        xlWS.Range("C" & NextBeam).Value = dbl_b 'B
        xlWS.Range("D" & NextBeam).Value = dbl_h 'H
        xlWS.Range("E" & NextBeam).Value = dbl_Lnhip ' L NHIP
        xlWS.Range("F" & NextBeam).FormulaR1C1 = "=RC[-1]-RC[3]"
        xlWS.Range("G" & NextBeam).Value = int_sck ' SO CAU KIEN
        xlWS.Range("H" & NextBeam).Value = int_sonhip ' SO NHIP
        xlWS.Range("I" & NextBeam).Value = dbl_Wcot ' BE RONG COT
        xlWS.Range("J" & NextBeam).Value = SLtheptren 'sl THEP TREN
        xlWS.Range("K" & NextBeam).Value = SplitText(str_ttren(0), 2) 'FI THEP TREN
        xlWS.Range("L" & NextBeam).Value = SLthepduoi 'SL THEP DUOI
        xlWS.Range("M" & NextBeam).Value = SplitText(str_tduoi(0), 2) ' FI THEP DUOI
        xlWS.Range("O" & NextBeam).Value = SplitText(str_tdai(0), 1) ' fi thep dai
        xlWS.Range("N" & NextBeam).FormulaR1C1 = "=ROUND(RC[-8]/" & "AVERAGE(" & Sdai1 & "," & Sdai2 & "),0)+1"  ' sl thep dai
        xlWS.Range("X" & NextBeam).Value = SlabThk ' SAN
        xlWS.Range("Y" & NextBeam).FormulaR1C1 = "=VLOOKUP(RC[-1],SlabsThkTable,2,0)"
        xlWS.Range("T" & NextBeam).Value = SplitText(str_tgia, 1) ' SL THEP GIA
        xlWS.Range("U" & NextBeam).Value = SplitText(str_tgia, 2) ' FI THEP GIA
        xlWS.Range("Z" & NextBeam).Value = objHandle
Calculate:
        acDoc.ModelSpace.AddLine(p1, p2)
        Call AddBeamLV3_2(str_BeamlV2)
Exit_Sub:
        For i = 0 To 10
            sset(i).Delete()
        Next i
        Call ReleaseAcadObj()
        Call ReleaseExcelObj()
    End Sub
    Sub ThepC_1()
        Dim sset(2) As AcadSelectionSet
        Dim dbl_khoangRaiC As Double
        Dim str_thepc As String
        Dim int_fiC As Integer
        Dim int_AThepC As Integer
        Dim WthepC As Double
        Dim int_SL As Integer
        Try
            Do
retry:
                ' NHAP TRONG ACAD
                For i As Integer = 0 To 2
                    sset(i) = acDoc.SelectionSets.Add("ThepC" & i)
                Next
                With acDoc.Utility
                    .Prompt("Chọn ký hiệu thép C: ")
                    sset(0).SelectOnScreen()
                    str_thepc = sset(0).Item(0).TextString
                    int_fiC = CType(SplitText(str_thepc, 1), Integer)
                    int_AThepC = CType(SplitText(str_thepc, 2), Integer)
                    .Prompt("Chọn DIM bề rộng thép C: ")
                    sset(1).SelectOnScreen()
                    Dim ent As AcadEntity
                    ent = sset(1).Item(0)
                    If ent.TextOverride <> "" Then
                        WthepC = ent.TextOVerride
                    Else
                        WthepC = ent.Measurement
                    End If
                    '.Prompt("Chọn khoảng rải thép C")
                    'sset(2).SelectOnScreen()
                    'dbl_khoangRaiC = 0
                    'For Each acEnt In sset(2) ' KHOANG RAI THEP C
                    '    If acEnt.TextOverride <> "" Then
                    '        dbl_khoangRaiC = dbl_khoangRaiC + CType(acEnt.TextOverride, Double)
                    '    Else
                    '        dbl_khoangRaiC = dbl_khoangRaiC + CType(acEnt.Measurement, Double)
                    '    End If
                    'Next
                    int_SL = .GetInteger("Số lượng: ")
                End With
                ' XUAT SANG EXCEL
                xlWS = xlWB.Sheets("dulieu_insert")
                xlWS.Range("C31").FormulaR1C1 = "=" & WthepC & "-2*be" ' chieu dai thep c
                xlWS.Range("F31").Value = int_fiC ' fi thep c
                xlWS.Range("G31").FormulaR1C1 = "=" & int_SL & "*(ROUND(" & dbl_Lnhip - dbl_Wcot & "/" & int_AThepC & ",0)+1)" 'SL
                xlWS.Rows("31").Copy()
                xlWS = xlWB.Sheets("addbeamlv3")
                Dim LastRow As Long
                LastRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                xlWS.Range("A" & LastRow).Insert()
                xlApp.CutCopyMode = False
                'XOA SELECTION SET
                For i As Integer = 0 To 2
                    sset(i).Delete()
                Next
            Loop
        Catch ex As Exception
            For i As Integer = 0 To 2
                sset(i).Delete()
            Next
            Exit Try
        End Try
    End Sub
    ' XU KY ADD BEAM KIEU 2 TRONG ACAD
    Sub AddBeamLV3FromACAD_2(str_beamlv2 As String, SlabTHK As String)
        On Error Resume Next
        Call StartExcel()
        Call StartACAD()
        Dim xlWS_insert As Worksheet
        Dim acEnt As AcadEntity
        Dim sset(0 To 10) As AcadSelectionSet
        Dim str_beamname As String
        Dim int_sck As Integer
        Dim str_ttren, str_tduoi, str_tdai(1), str_tcgoi, str_tcnhip, str_tgia, str_thepc As String
        Dim dbl_Ldaigiuanhip As Double = 0
        Dim dbl_Ldaikegoi As Double = 0
        Dim dbl_Lraidai As Double = 0
        Dim Sdai1, Sdai2 As Integer
        Dim fiDai As Integer
        Dim i As Integer
        Dim int_fiC As Integer
        Dim int_AThepC As Integer
        Dim acText_tendam As AcadEntity
        Dim acTextHandle As String
        Dim dbl_KhoangRaiC As Double = 0
        xlWS_insert = xlWB.Sheets("dulieu_insert")
        ' XOA HET DU LIEU TREN SHEET ADDBEAMLV3
        xlWS = xlWB.Sheets("addbeamlv3")
        xlWS.UsedRange.EntireRow.Delete(Shift:=XlDirection.xlUp)
        For i = 0 To 10
            sset(i) = acDoc.SelectionSets.Add("acadfile2" & i)
        Next
        ' 1. CHON TEN DAM, BXH, THEP CHU TREN, THEP CHU DUOI, THEP DAI
        With acDoc.Utility
            ' Chon ten dam
            .Prompt(vbCrLf & "Chọn tên dầm: ")
            sset(0).SelectOnScreen()
            Dim temptxt As String
            acText_tendam = sset(0).Item(0)
            temptxt = sset(0).Item(0).Textstring
            str_beamname = RemoveLeftString(temptxt, "%%U")
            TextHeight = sset(0).Item(0).Height
            acTextHandle = sset(0).Item(0).Handle
            ' chon BxH
            Call ChonBH()
            ' CHON THEP CHU TREN
            .Prompt("Chọn thép chủ trên: ")
            sset(2).SelectOnScreen()
            str_ttren = sset(2).Item(0).TextString
            ' CHON THEP CHU DUOI
            .Prompt("Chọn thép chủ dưới: ")
            sset(3).SelectOnScreen()
            str_tduoi = sset(3).Item(0).TextString
        End With

        '    .Prompt("Chọn thép đai: ")
        '    sset(4).SelectOnScreen()
        '    str_tdai(0) = sset(4).Item(0).TextString
        '    str_tdai(1) = sset(4).Item(1).TextString
        '    Sdai1 = 0
        '    Sdai2 = 0
        '    Sdai1 = CType(SplitText(str_tdai(0), 2), Integer)
        '    Sdai2 = CType(SplitText(str_tdai(1), 2), Integer)
        '    fiDai = SplitText(str_tdai(0), 1)
        '    Dim int_temp As Integer
        '    If Sdai2 <> 0 Then
        '        If Sdai1 < Sdai2 Then ' Adai1 > Adai2
        '            int_temp = Sdai1
        '            Sdai1 = Sdai2
        '            Sdai2 = int_temp
        '        End If
        '        .Prompt("Chọn dim khoảng rải đai giữa nhịp: ")
        '        sset(5).SelectOnScreen() ' chon khoang rai dai dau nhip
        '        For Each acEnt In sset(5) ' tinh khoang rai dai giua nhip
        '            If acEnt.TextOverride <> "" Then
        '                dbl_Ldaigiuanhip = dbl_Ldaigiuanhip + CType(acEnt.TextOverride, Double)
        '            Else
        '                dbl_Ldaigiuanhip = dbl_Ldaigiuanhip + acEnt.Measurement
        '            End If
        '        Next
        '        .Prompt("Chọn dim khoảng rải đai kể gối: ")
        '        sset(7).SelectOnScreen()
        '        For Each acEnt In sset(7)
        '            If acEnt.TextOverride <> "" Then
        '                dbl_Ldaikegoi = dbl_Ldaikegoi + CType(acEnt.TextOverride, Double)
        '            Else
        '                dbl_Ldaikegoi = dbl_Ldaikegoi + acEnt.Measurement
        '            End If
        '        Next
        '    Else
        '        .Prompt("Chọn dim khoảng rải đai: ")
        '        sset(8).SelectOnScreen()
        '        For Each acEnt In sset(8)
        '            If acEnt.TextOverride <> "" Then
        '                dbl_Lraidai = dbl_Lraidai + CType(acEnt.TextOverride, Double)
        '            Else
        '                dbl_Lraidai = dbl_Lraidai + acEnt.Measurement
        '            End If
        '        Next
        '    End If
        'End With
        ' 2. ADD TRUOC THONG TIN DAM VA THEP CHU & THEP DAI VAO SHEET ADDBEAMLV3
        xlWS_insert.Rows("26:29").Copy()
        xlWS.Activate()
        xlWS.Range("A1").Insert(XlDirection.xlDown)
        xlApp.CutCopyMode = False
        ' CHON THEP DAI
        Call NhapThepDai()
        ' CHON THEP TANG CUONG GOI
        Call ThepTCGoi()
        ' CHON THEP TANG CUONG NHIP
        Call ThepTCNhip()

        ' 3. CHON THEP GIA
        acDoc.Utility.Prompt("Chọn thép giá: ")
        sset(6).SelectOnScreen()
        str_tgia = sset(6).Item(0).textstring
        ' CHON THEP C
        Call ThepC_2()
        ' CHEN HANG KET THUC BEAM
        Dim LastRow As Long
        xlWS = xlWB.Sheets("addbeamlv3")
        xlWS_insert.Rows("38").copy()
        LastRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlWS.Range("A" & LastRow).Insert()
        xlApp.CutCopyMode = False
        With acDoc.Utility
            ' NHAP SO CK
            int_sck = 1
            int_sck = CType(.GetInteger("Nhập số CK: "), Integer)
            Call FindtextInAcad() ' tim dam tren mat bang ket cau
            ' NHAP SO NHIP
            int_sonhip = 1
            int_sonhip = CType(.GetInteger("Nhập số nhịp: "), Integer)
        End With
        ' 4. VE CHIEU DAI DAM
        dbl_Wcot = 0
        Call VeLdam()
        ' 5. CO NHAP DAM NAY HAY KO
        Dim opt As String
        opt = "Yes"
        acDoc.Utility.InitializeUserInput(0, "Yes No")
        opt = acDoc.Utility.GetKeyword("Nhập dầm này Yes[No]: ")
        If opt = "No" Then
            For i = 0 To 10
                sset(i).Delete()
            Next
            Call ReleaseAcadObj()
            Call ReleaseExcelObj()
            Exit Sub
        End If
        ' 6. NHAP VAO SHEET BEAM DATA EXCEL
        xlWS = xlWB.Sheets("beamdata")
        Dim NextBeam As Long = FindCurrBeam() + 1
        xlWS.Range("A" & NextBeam).Value = "BEAM" & NextBeam ' CODE
        xlWS.Range("B" & NextBeam).Value = str_beamname ' TEN DAM
        xlWS.Range("C" & NextBeam).Value = dbl_b 'B
        xlWS.Range("D" & NextBeam).Value = dbl_h 'H
        xlWS.Range("E" & NextBeam).Value = dbl_Lnhip ' L NHIP
        xlWS.Range("F" & NextBeam).FormulaR1C1 = "=RC[-1]-RC[3]"
        xlWS.Range("G" & NextBeam).Value = int_sck ' SO CAU KIEN
        xlWS.Range("H" & NextBeam).Value = int_sonhip ' SO NHIP
        xlWS.Range("I" & NextBeam).Value = dbl_Wcot ' BE RONG COT
        xlWS.Range("J" & NextBeam).Value = SplitText(str_ttren, 1)
        xlWS.Range("K" & NextBeam).Value = SplitText(str_ttren, 2)
        xlWS.Range("L" & NextBeam).Value = SplitText(str_tduoi, 1)
        xlWS.Range("M" & NextBeam).Value = SplitText(str_tduoi, 2)
        xlWS.Range("X" & NextBeam).Value = SlabTHK ' SAN
        xlWS.Range("Y" & NextBeam).FormulaR1C1 = "=VLOOKUP(RC[-1],SlabsThkTable,2,0)"
        xlWS.Range("T" & NextBeam).Value = SplitText(str_tgia, 1) ' SL THEP GIA
        xlWS.Range("U" & NextBeam).Value = SplitText(str_tgia, 2) ' FI THEP GIA
        xlWS.Range("Z" & NextBeam).Value = objHandle
        ' 7. ADD DU LIEU VAO CAC SHEET TINH TOAN
        Dim p1, p2
        acText_tendam.GetBoundingBox(p1, p2)
        acDoc.ModelSpace.AddLine(p1, p2)
        acDoc.HandleToObject(acTextHandle)
        acApp.ZoomWindow(p1, p2)
        Call AddBeamLV3_2(str_beamlv2)
exit_sub:
        ' 8. XOA SSET
        For i = 0 To 10
            sset(i).Delete()
        Next
        ' 9. RELEASE MEMORY
        Call ReleaseAcadObj()
        Call ReleaseExcelObj()
    End Sub
    Sub AddBeamLV3_2(str_BeamLV2 As String)
        Dim InsertRow, CurrRow As Long
        On Error Resume Next
        CurrRow = FindNextBeam() - 1 ' Vi thu tuc nay goi sau khi da add du lieu vao sheet beamdata nen phai -1
        xlWS = xlWB.Sheets("addbeamlv3") : xlWS.Range("A1").Value = "BEAM" & CurrRow
        With xlWB.Sheets("dulieu_insert") ' Ghi ten dam vao sheet dulieu_insert truoc khi insert 
            .Range("A21") = "BEAM" & CurrRow
            .Range("A23") = "BEAM" & CurrRow
        End With
        ' Chen dam LV3 vao sheet Concrete
        xlWS = xlWB.Sheets("Concrete")
        InsertRow = FindInsertRow("Concrete", "M", str_BeamLV2)
        xlWB.Sheets("dulieu_insert").Rows(21).Copy()
        xlWS.Activate()
        xlWS.Range("A" & InsertRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        ' Chen dam LV3 vao sheet Formwork
        xlWS = xlWB.Sheets("Formwork")
        InsertRow = FindInsertRow("Formwork", "M", str_BeamLV2)
        xlWB.Sheets("dulieu_insert").Rows(23).Copy()
        xlWS.Activate()
        xlWS.Range("A" & InsertRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        ' Chen dam LV3 vao sheet Reinforcement
        xlWB.Sheets("addbeamlv3").UsedRange.EntireRow.Copy()
        xlWS = xlWB.Sheets("Reinforcement")
        xlWS.Activate()
        InsertRow = FindInsertRow("Reinforcement", "P", str_BeamLV2)
        xlWS.Range("A" & InsertRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
    End Sub
    ' XU LY ADD BEAM KIEU 3 TRONG AUTOCAD
    Sub AddBeamLV3FromACAD_3(str_BeamlV2 As String, SlabThk As String)
        On Error Resume Next
        Dim sset(0 To 10) As AcadSelectionSet
        Dim str_beamname As String
        Dim acText As AcadEntity
        Dim int_sck As Integer
        Dim xlws_insert As Worksheet
        Dim str_ttren(1), str_tduoi(1), str_tdai(0 To 1), str_tcgoi, str_tcnhip, str_tgia As String
        Dim Sdai1, Sdai2 As Double
        Dim dbl_LnhipTB, dbl_WcotTB As Double
        Dim acTextHandle As String
        Dim SLtheptren, SLthepduoi As Integer
        Call StartACAD()
        Call StartExcel()
        xlws_insert = xlWB.Sheets("dulieu_insert")
        Dim i As Integer
        ' XOA HET DU LIEU TREN SHEET ADDBEAMLV3
        xlWS = xlWB.Sheets("addbeamlv3")
        xlWS.UsedRange.EntireRow.Delete(Shift:=XlDirection.xlUp)
        For i = 0 To 10
            sset(i) = acDoc.SelectionSets.Add("Acadfile1" & i)
        Next
        With acDoc.Utility
            ' Chon ten dam
            .Prompt(vbCrLf & "Chọn tên dầm: ")
            sset(0).SelectOnScreen()
            acText = sset(0).Item(0)
            str_beamname = acText.Textstring
            str_beamname = RemoveLeftString(str_beamname, "%%U")
            TextHeight = acText.Height
            acText.color = ACAD_COLOR.acRed
            acTextHandle = acText.Handle
        End With
        ' NHAP BH
        dbl_b = acDoc.Utility.GetReal("Nhập b:")
        dbl_h = acDoc.Utility.GetReal("Nhập h:")
        ' NHAP SO CK, SO NHIP
        With acDoc.Utility
            ' Nhap so ck
            int_sck = 1
            int_sck = CType(.GetInteger("Nhập số CK: "), Integer)
            '' TIM TEXT
            'Call FindtextInAcad()
            ' Nhap so nhip
            int_sonhip = 1
            int_sonhip = CType(.GetInteger("Nhập số nhịp: "), Integer)
        End With
        ' VE LINE DAM
        dbl_Wcot = 0
        dbl_Lnhip = 0
        Call VeLdam()
        ' TIEP TUC ADD DU LIEU VAO SHEET ADDBEAMLV3
        Dim R As Long
        Dim LastRow As Long
        ' CO NHAP DAM NAY HAY KHONG
        Dim opt As String
        opt = "Yes"
        acDoc.Utility.InitializeUserInput(0, "Yes No")
        opt = acDoc.Utility.GetKeyword("Nhập dầm này Yes[No]: ")
        Select Case opt
            Case "No"
                GoTo Exit_Sub
        End Select
        ' XUAT DU LIEU SANG SHEET BEAM DATA
        xlWS = xlWB.Sheets("beamdata")
        Dim NextBeam As Long = FindCurrBeam() + 1
        xlWS.Range("A" & NextBeam).Value = "BEAM" & NextBeam ' CODE
        xlWS.Range("B" & NextBeam).Value = str_beamname ' TEN DAM
        xlWS.Range("C" & NextBeam).Value = dbl_b 'B
        xlWS.Range("D" & NextBeam).Value = dbl_h 'H
        xlWS.Range("E" & NextBeam).Value = dbl_Lnhip ' L NHIP
        xlWS.Range("F" & NextBeam).FormulaR1C1 = "=RC[-1]-RC[3]"
        xlWS.Range("G" & NextBeam).Value = int_sck ' SO CAU KIEN
        xlWS.Range("H" & NextBeam).Value = int_sonhip ' SO NHIP
        xlWS.Range("I" & NextBeam).Value = dbl_Wcot ' BE RONG COT
        xlWS.Range("J" & NextBeam).Value = SLtheptren 'sl THEP TREN
        xlWS.Range("K" & NextBeam).Value = SplitText(str_ttren(0), 2) 'FI THEP TREN
        xlWS.Range("L" & NextBeam).Value = SLthepduoi 'SL THEP DUOI
        xlWS.Range("M" & NextBeam).Value = SplitText(str_tduoi(0), 2) ' FI THEP DUOI
        xlWS.Range("O" & NextBeam).Value = SplitText(str_tdai(0), 1) ' fi thep dai
        xlWS.Range("N" & NextBeam).FormulaR1C1 = "=ROUND(RC[-8]/" & "AVERAGE(" & Sdai1 & "," & Sdai2 & "),0)+1"  ' sl thep dai
        xlWS.Range("X" & NextBeam).Value = SlabThk ' SAN
        xlWS.Range("Y" & NextBeam).FormulaR1C1 = "=VLOOKUP(RC[-1],SlabsThkTable,2,0)"
        xlWS.Range("T" & NextBeam).Value = SplitText(str_tgia, 1) ' SL THEP GIA
        xlWS.Range("U" & NextBeam).Value = SplitText(str_tgia, 2) ' FI THEP GIA
        xlWS.Range("Z" & NextBeam).Value = objHandle
Calculate:
        Call AddBeamLV3_2(str_BeamlV2)
Exit_Sub:
        For i = 0 To 10
            sset(i).Delete()
        Next i
        Call ReleaseAcadObj()
        Call ReleaseExcelObj()
    End Sub
    ' THU TUC XU LY THEP DAM
    Sub NhapThepDai()
        Dim sset(2) As AcadSelectionSet
        Dim kyhieu As String
        Dim i As Integer
        Dim ent As AcadEntity
        Dim dbl_khoangraidai As Double
        Dim fi As Integer
        Dim S As Double
        Try
            Do
                For i = 0 To 2
                    sset(i) = acDoc.SelectionSets.Add("ThepDai" & i)
                Next
                With acDoc.Utility
                    .Prompt("Chọn ký hiệu thép đai: ")
                    sset(0).SelectOnScreen()
                    kyhieu = sset(0).Item(0).textstring
                    fi = CType(SplitText(kyhieu, 1), Integer)
                    S = CType(SplitText(kyhieu, 2), Double)
                    .Prompt("Chọn khoảng rải: ")
                    sset(1).SelectOnScreen()
                End With
                dbl_khoangraidai = 0
                For Each ent In sset(1)
                    If Not ent.textoverride = "" Then
                        dbl_khoangraidai = dbl_khoangraidai + CType(ent.textoverride, Double)
                    Else
                        dbl_khoangraidai = dbl_khoangraidai + CType(ent.measurement, Double)
                    End If
                Next
                ' XUAT SANG SHEET dulieu_insert
                xlWS = xlWB.Sheets("dulieu_insert")
                xlWS.Range("B30").FormulaR1C1 = "=" & dbl_b & "-2*be"
                xlWS.Range("C30").FormulaR1C1 = "=" & dbl_h & "-2*be"
                xlWS.Range("D30").FormulaR1C1 = "=VLOOKUP(RC[2],hooks,2,0)"
                xlWS.Range("F30").Value = fi
                xlWS.Range("G30").FormulaR1C1 = "=ROUND(" & dbl_khoangraidai & "/" & S & ",0)+1"
                xlWS.Rows("30").copy()
                ' XUAT SANG SHEET ADDBEAM LV3
                xlWS = xlWB.Sheets("addbeamlv3")
                Dim nextrow As Long
                nextrow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                xlWS.Range("A" & nextrow).Insert()
                xlApp.CutCopyMode = False
                For i = 0 To 2
                    sset(i).Delete()
                Next
            Loop
        Catch ex As Exception
            For i = 0 To 2
                sset(i).Delete()
            Next
            Exit Try
        End Try
    End Sub
    Sub ChonBH()
        Dim sset As AcadSelectionSet
        Dim i As Integer
        Try
retry:
            sset = acDoc.SelectionSets.Add("bh" & i)
            With acDoc.Utility
                .Prompt(vbCrLf & "Chọn BxH: ")
                sset.SelectOnScreen()
                dbl_b = sset.Item(0).Measurement
                dbl_h = sset.Item(1).Measurement
            End With
            sset.Delete()
        Catch EX As ArgumentException
            sset.Delete()
            acDoc.Utility.Prompt("Vui lòng nhập lại")
            GoTo retry
        End Try
        sset.Delete()
    End Sub
    Sub ThepTCGoi()
        Dim sset(0 To 3) As AcadSelectionSet
        Dim i As Integer
        Dim str_kyhieuthep As String
        Dim fi As Integer = 0
        Dim Num As Integer = 0
        Dim L As Double = 0
        Dim NextRow As Long
        Dim opt As String
        Dim sl As Integer
        Try
try1:
            Do
                For i = 0 To 3
                    sset(i) = acDoc.SelectionSets.Add("thepTCgoi" & i)
                Next
                ' Chon loai thep
try4:
                With acDoc.Utility
                    .Prompt("Chọn loại thép TC gối: ")
                    .InitializeUserInput(0, "1 2 3")
                    opt = .GetKeyword("1<td-td> 2<ne-tu> 3<ne-ne>")
                    Select Case opt
                        Case 1
                            str_loaithep = "td-td"
                        Case 2
                            str_loaithep = "ne-td"
                        Case 3
                            str_loaithep = "ne-ne"
                    End Select
                End With
                ' CHON TEN THEP
                acDoc.Utility.Prompt("Chọn tên thép TC gối[ESC để bỏ qua]: ")
                sset(0).SelectOnScreen()
                If Not sset(0).Item(0) Is Nothing Then str_tenthep = sset(0).Item(0).TextString
try2:           ' CHON KY HIEU THEP
                acDoc.Utility.Prompt("Chọn ký hiệu thép TC gối[ESC để thoát] : ")
                sset(1).SelectOnScreen()
                str_kyhieuthep = sset(1).Item(0).TextString
                fi = CType(SplitText(str_kyhieuthep, 1), Integer)
                SL = CType(SplitText(str_kyhieuthep, 2), Integer)
Try3:           ' Chon dim the hien chieu dai cay thep
                acDoc.Utility.Prompt("Chọn DIM hoặc LINE : ")
                sset(2).SelectOnScreen()
                L = 0
                Dim ent As AcadEntity
                For Each ent In sset(2)
                    If TypeOf ent Is AcadLine Or TypeOf ent Is AcadPolyline Or _
                        TypeOf ent Is AcadLWPolyline Or TypeOf ent Is Acad3DPolyline Then
                        L = L + CType(ent.Length, Double)
                    Else
                        If ent.textoverride <> "" Then
                            L = L + CType(ent.TextOverride, Double)
                        Else
                            L = L + ent.Measurement
                        End If
                    End If
                Next

                '   ADD VAO EXCEL
                xlWS = xlWB.Sheets("dulieu_insert")
                Dim R As Long
                Select Case str_loaithep
                    Case "ne-td"
                        R = 32
                    Case "td-td"
                        R = 33
                    Case "ne-ne"
                        R = 34
                End Select
                xlWS.Range("A" & R).Value = str_tenthep
                xlWS.Range("F" & R).Value = SplitText(str_kyhieuthep, 2) ' Fi
                xlWS.Range("G" & R).Value = SplitText(str_kyhieuthep, 1) ' SL
                xlWS.Range("C" & R).Value = L
                xlWS.Rows(R).Copy()
                xlWS = xlWB.Sheets("addbeamlv3")
                xlWS.Activate()
                NextRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                xlWS.Range("A" & NextRow).Insert()
                xlApp.CutCopyMode = False
                For i = 0 To 3
                    sset(i).Delete()
                Next
            Loop
        Catch ex As Exception ' KHI NGUOI DUNG NHAN ESC 

            For i = 0 To 3
                sset(i).Delete()
            Next
            Exit Try
        End Try
    End Sub
    Sub ThepTCNhip()
        Dim sset(0 To 3) As AcadSelectionSet
        Dim i As Integer
        Dim str_kyhieuthep As String
        Dim fi As Integer = 0
        Dim sl As Integer
        Dim Num As Integer = 0
        Dim L As Double = 0
        Dim NextRow As Long
        Dim opt As String
        Dim Lstring As String
        Try
try1:
            Do
                For i = 0 To 3
                    sset(i) = acDoc.SelectionSets.Add("thepTCnhip" & i)
                Next
                ' Chon loai thep
try4:
                With acDoc.Utility
                    .Prompt("Chọn loại thép TC nhịp: ")
                    .InitializeUserInput(0, "1 2 3")
                    opt = .GetKeyword("1<td-td> 2<ne-tu> 3<ne-ne>")
                    Select Case opt
                        Case 1
                            str_loaithep = "td-td"
                        Case 2
                            str_loaithep = "ne-td"
                        Case 3
                            str_loaithep = "ne-ne"
                    End Select
                End With
                ' CHON TEN THEP
                acDoc.Utility.Prompt("Chọn tên thép TC nhịp[ESC để bỏ qua]: ")
                sset(0).SelectOnScreen()
                If Not sset(0).Item(0) Is Nothing Then str_tenthep = sset(0).Item(0).TextString
try2:           ' CHON KY HIEU THEP
                acDoc.Utility.Prompt("Chọn ký hiệu thép TC nhịp[ESC để thoát] : ")
                sset(1).SelectOnScreen()
                str_kyhieuthep = sset(1).Item(0).TextString
                fi = CType(SplitText(str_kyhieuthep, 1), Integer)
                SL = CType(SplitText(str_kyhieuthep, 2), Integer)
Try3:           ' Chon dim the hien chieu dai cay thep
                acDoc.Utility.InitializeUserInput(0, "1 2")
                opt = acDoc.Utility.GetKeyword("1<vẽ LINE hoặc chọn DIM> 2<Tính tay>")
                Select Case opt
                    Case 1
                        acDoc.Utility.Prompt("Chọn DIM hoặc LINE : ")
                        sset(2).SelectOnScreen()
                        L = 0
                        Dim ent As AcadEntity
                        For Each ent In sset(2)
                            If TypeOf ent Is AcadLine Or TypeOf ent Is AcadPolyline Or _
                                TypeOf ent Is AcadLWPolyline Or TypeOf ent Is Acad3DPolyline Then
                                L = L + CType(ent.Length, Double)
                            Else
                                If ent.textoverride <> "" Then
                                    L = L + CType(ent.TextOverride, Double)
                                Else
                                    L = L + ent.Measurement
                                End If
                            End If
                        Next
                    Case 2
                        Lstring = acDoc.Utility.GetString(False, "nhập công thức: ")
                End Select
                '   ADD VAO EXCEL
                xlWS = xlWB.Sheets("dulieu_insert")
                Dim R As Long
                Select Case str_loaithep
                    Case "ne-td"
                        R = 35
                    Case "td-td"
                        R = 36
                    Case "ne-ne"
                        R = 37
                End Select
                xlWS.Range("A" & R).Value = str_tenthep
                xlWS.Range("F" & R).Value = SplitText(str_kyhieuthep, 2) ' Fi
                xlWS.Range("G" & R).Value = SplitText(str_kyhieuthep, 1) ' SL
                Select Case opt
                    Case 1
                        xlWS.Range("C" & R).Value = L
                    Case 2
                        xlWS.Range("C" & R).FormulaR1C1 = "=" & Lstring
                End Select
                xlWS.Rows(R).Copy()
                xlWS = xlWB.Sheets("addbeamlv3")
                xlWS.Activate()
                NextRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                xlWS.Range("A" & NextRow).Insert(XlDirection.xlDown)
                xlApp.CutCopyMode = False
                For i = 0 To 3
                    sset(i).Delete()
                Next

            Loop
        Catch ex As Exception ' KHI NGUOI DUNG NHAN ESC 
            For i = 0 To 3
                sset(i).Delete()
            Next
            Exit Try
        End Try
    End Sub
    Sub VeLdam()
        Dim sset As AcadSelectionSet
        Dim pt1, pt2
        Dim acLine_Beam As AcadLine
        Dim Midpt(0 To 2) As Double
        Dim acText_BeamName As AcadText
        ' TAO LAYER MOI
        Call AddLayer("BOQTools(c)DHK_BEAMTOOL")
        ' BAT DAU VE DAM
        pt1 = CType(acDoc.Utility.GetPoint(, "Chọn điểm đầm dầm: "), Double())
        pt2 = CType(acDoc.Utility.GetPoint(pt1, "Chọn điểm cuối dầm: "), Double())
        acLine_Beam = acDoc.ModelSpace.AddLine(pt1, pt2)
        acLine_Beam.Lineweight = ACAD_LWEIGHT.acLnWt060 ' LINE WEIGHT
        acLine_Beam.color = ACAD_COLOR.acMagenta ' LINE COLOR
        dbl_Lnhip = 0
        dbl_Lnhip = acLine_Beam.Length
        ' Add text tren line
        Midpt(0) = (acLine_Beam.EndPoint(0) - acLine_Beam.StartPoint(0)) / 2.0# + acLine_Beam.StartPoint(0)
        Midpt(1) = (acLine_Beam.EndPoint(1) - acLine_Beam.StartPoint(1)) / 2.0# + acLine_Beam.StartPoint(1)
        Midpt(2) = (acLine_Beam.EndPoint(2) - acLine_Beam.StartPoint(2)) / 2.0# + acLine_Beam.StartPoint(2)
        Dim str As String = "BEAM" & FindCurrBeam() + 1
        acText_BeamName = acDoc.ModelSpace.AddText(str, Midpt, TextHeight)
        acText_BeamName.color = ACAD_COLOR.acRed
        acText_BeamName.Update()
        objHandle = acText_BeamName.Handle
        ' TIM BE RONG COT
        If int_sonhip > 1 Then
            Dim opt As String
            acDoc.Utility.InitializeUserInput(0, "1 2")
            opt = acDoc.Utility.GetKeyword("1<Chon dim>|2<Ve line>:")
            Select Case opt
                Case 1
                    acDoc.Utility.Prompt("Chọn dim: ")
                    acDoc.SelectionSets.Add("ChonDIMCot")
                    sset.SelectOnScreen()
                    dbl_Wcot = 0
                    For Each acEnt In sset
                        If acEnt.TextOverride <> "" Then
                            dbl_Wcot = dbl_Wcot + CType(acEnt.TextOverride, Double)
                        Else
                            dbl_Wcot = dbl_Wcot + acEnt.Measurement
                        End If
                    Next
                    sset.Delete()
                Case 2
                    dbl_Wcot = 0
                    dbl_Wcot = VeLineCot()
            End Select
        End If

    End Sub
    Sub ThepC_2()
        Dim sset(2) As AcadSelectionSet
        Dim dbl_khoangRaiC As Double
        Dim str_thepc As String
        Dim int_fiC As Integer
        Dim int_AThepC As Integer
        Dim WthepC As Double
        Dim int_SL As Integer
        Try
            Do
retry:
                ' NHAP TRONG ACAD
                For i As Integer = 0 To 2
                    sset(i) = acDoc.SelectionSets.Add("ThepC" & i)
                Next
                With acDoc.Utility
                    .Prompt("Chọn ký hiệu thép C: ")
                    sset(0).SelectOnScreen()
                    str_thepc = sset(0).Item(0).TextString
                    int_fiC = CType(SplitText(str_thepc, 1), Integer)
                    int_AThepC = CType(SplitText(str_thepc, 2), Integer)
                    .Prompt("Chọn DIM bề rộng thép C: ")
                    sset(1).SelectOnScreen()
                    Dim ent As AcadEntity
                    ent = sset(1).Item(0)
                    If ent.TextOverride <> "" Then
                        WthepC = ent.TextOVerride
                    Else
                        WthepC = ent.Measurement
                    End If
                    .Prompt("Chọn khoảng rải thép C")
                    sset(2).SelectOnScreen()
                    dbl_khoangRaiC = 0
                    For Each acEnt In sset(2) ' KHOANG RAI THEP C
                        If acEnt.TextOverride <> "" Then
                            dbl_khoangRaiC = dbl_khoangRaiC + CType(acEnt.TextOverride, Double)
                        Else
                            dbl_khoangRaiC = dbl_khoangRaiC + CType(acEnt.Measurement, Double)
                        End If
                    Next
                    int_SL = .GetInteger("Số lượng: ")
                End With
                ' XUAT SANG EXCEL
                xlWS = xlWB.Sheets("dulieu_insert")
                xlWS.Range("C31").FormulaR1C1 = "=" & WthepC & "-2*be" ' chieu dai thep c
                xlWS.Range("F31").Value = int_fiC ' fi thep c
                xlWS.Range("G31").FormulaR1C1 = "=" & int_SL & "*(ROUND(" & dbl_khoangRaiC & "/" & int_AThepC & ",0)+1)" 'SL
                xlWS.Rows("31").Copy()
                xlWS = xlWB.Sheets("addbeamlv3")
                Dim LastRow As Long
                LastRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                xlWS.Range("A" & LastRow).Insert()
                xlApp.CutCopyMode = False
                'XOA SELECTION SET
                For i As Integer = 0 To 2
                    sset(i).Delete()
                Next
            Loop
        Catch ex As Exception
            For i As Integer = 0 To 2
                sset(i).Delete()
            Next
            Exit Try
        End Try
    End Sub
    Function VeLineCot() As Double
        Dim pt1, pt2
        Dim Col As Double = 0
        Dim Wcot As AcadLine
        With acDoc.Utility
            Do
                Try
                    pt1 = CType(.GetPoint(, "Nhập điểm đầu cột[ESC để thoát]: "), Double())
                    pt2 = CType(.GetPoint(pt1, "Nhập điểm cuối cột[ESC để thoát]: "), Double())
                    Wcot = acDoc.ModelSpace.AddLine(pt1, pt2)
                    Wcot.color = ACAD_COLOR.acBlue
                    Wcot.Lineweight = ACAD_LWEIGHT.acLnWt050
                    Col = Col + Wcot.Length
                Catch ex As Exception
                    If ex.Message Like "*0x80020009*" Then Exit Do
                End Try
            Loop
        End With
        Return Col
        GC.Collect()
    End Function
    Sub BeRongCot()
        Dim opt As String
        dbl_Wcot = 0
        If int_sonhip > 1 Then
            With acDoc.Utility
                .InitializeUserInput(0, "1 2")
                opt = .GetKeyword("1<Nhập tay> 2<Vẽ LINE>: ")
                Select Case opt
                    Case 1
                        dbl_Wcot = .GetInteger("Nhập bề rộng TB 1 cột: ")
                        dbl_Wcot = dbl_Wcot * (int_sonhip - 1)
                    Case 2
                        dbl_Wcot = VeLineCot()
                End Select
            End With
        End If
    End Sub
    ' CAC THU TUC HO TRO
    Sub AcquireOutline()
        Dim NextRow As Long
        Dim top, bot, j, i As Long
        Dim countlv1, countlv2 As Long
        Dim Cell As Range = Nothing
        Dim WorkRange As Range = Nothing
        Dim RangeTop(100) As Range
        Dim RangeBot(100) As Range
        Dim Listlv1(100), Listlv2(1000) As String
        Dim WriteRange As Range
        Dim sheet_Outline As Worksheet
        Dim RangeLV2 As Range
        Dim LastRow As Long
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        xlWS = xlWB.Sheets("Formwork")
        '   Xoa du lieu cu
        sheet_Outline = xlWB.Sheets("List")
        sheet_Outline.Range("A1").CurrentRegion.ClearContents()
        '   Xac dinh dau muc LV1 trong sheet Reinforcement va Ghi vao Sheet List
        WriteRange = sheet_Outline.Range("A1")
        WorkRange = xlApp.Intersect(xlWS.Columns("M"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas) ' Column chua thong tin Beam LV1
        top = 0
        bot = 0
        For Each Cell In WorkRange
            If Cell.Text <> "" Then
                If Left(Cell.Text, 1) <> "/" Then
                    top = top + 1
                    RangeTop(top) = Cell
                    Listlv1(top) = Cell.Text
                Else
                    bot = bot + 1
                    RangeBot(bot) = Cell
                End If
            End If
        Next Cell
        '   Ghi dau muc lv1 vao sheet
        For j = 1 To top
            NextRow = sheet_Outline.Range("A" & sheet_Outline.Rows.Count).End(XlDirection.xlUp).Row + 1
            sheet_Outline.Cells(NextRow, WriteRange.Column) = Listlv1(j)
        Next j
        '   Xac dinh dau muc level 2
        countlv2 = 0
        For j = 1 To top
            RangeLV2 = xlWS.Range(RangeTop(j).Offset(0, 1), RangeBot(j).Offset(0, 1))
            For Each Cell In RangeLV2
                If Left(Cell.Text, 1) <> "/" And Cell.Text <> "" Then
                    countlv2 = countlv2 + 1
                    Listlv2(countlv2) = Listlv1(j) & "/" & Cell.Text
                End If
            Next Cell
        Next j
        '   Ghi dau muc level 2 vao sheet
        For j = 1 To countlv2
            NextRow = sheet_Outline.Range("B" & sheet_Outline.Rows.Count).End(XlDirection.xlUp).Row + 1
            sheet_Outline.Range("B" & NextRow).Value = Listlv2(j)
        Next j
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub
    Function FindCurrBeam()
        Dim Curr As Long
        xlWS = xlWB.Sheets("beamdata")
        Curr = xlWS.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row
        Return Curr
    End Function
    Sub ZoomToObject()
        Dim ent As AcadEntity
        Dim pt1, pt2
        Call StartACAD()
        Call StartExcel()
        Try
            pt1 = Nothing : pt2 = Nothing
            Dim handleID As String = Nothing
            Dim str_tendam As String
            str_tendam = xlApp.ActiveCell.Value
            xlWS = xlWB.Sheets("beamdata")
            Dim WorkRange As Excel.Range = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("A"))
            For Each Cell As Excel.Range In WorkRange
                If TypeOf Cell.Value Is String Then
                    If CType(Cell.Value, String) = str_tendam Then
                        handleID = CType(xlWS.Range("Z" & Cell.Row).Value, String)
                    End If
                End If
            Next
            ent = acDoc.HandleToObject(handleID)
            ent.GetBoundingBox(pt1, pt2)
            acApp.ZoomWindow(pt1, pt2)
            Call ReleaseAcadObj()
            Call ReleaseExcelObj()
        Catch ex As Exception
            MsgBox("Đối tượng không tồn tại hoặc bản vẽ hiện hành không đúng", vbCritical, "BOQ-Tool@DHK")
        End Try
    End Sub
    Function FindNextBeam() As Long
        Dim NextRow As Long
        xlWS = xlWB.Sheets("beamdata")
        NextRow = xlWS.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        Return NextRow
    End Function
    Function FindInsertRow(SheetName As String, ColLV1 As String, str_BeamLV2 As String) As Long
        Dim TextLv1, TextLv2 As String
        Dim Cell, WorkRange As Range
        Dim StartCellLV1, EndCellLV1 As Range
        Dim CurrBeam As Long
        Dim InsertRow As Long
        On Error Resume Next
        CurrBeam = FindNextBeam() - 1
        If CurrBeam <= 4 Then CurrBeam = 5
        '-----------------------------
        '   Lay Current Index cua beam LV1,LV2
        Dim arr1(2) As String
        arr1 = Split(str_BeamLV2, "/")
        TextLv1 = arr1(0)
        TextLv2 = arr1(1)

        xlWS = xlWB.Sheets(SheetName)
        xlWS.Activate()
        ' Tim vung LV1
        Dim WR As Range
        WR = xlApp.Intersect(xlWS.Columns(ColLV1), xlWS.UsedRange)
        WorkRange = WR.SpecialCells(XlCellType.xlCellTypeFormulas)
        For Each Cell In WorkRange
            Select Case Cell.Text
                Case TextLv1
                    StartCellLV1 = Cell
                Case "/" & TextLv1
                    EndCellLV1 = Cell
                    Exit For
            End Select
        Next Cell
        ' Tim vung LV2
        WR = xlWS.Range(StartCellLV1.Offset(0, 1), EndCellLV1.Offset(0, 1))
        WorkRange = WR.SpecialCells(XlCellType.xlCellTypeFormulas)
        TextLv2 = "/" & TextLv2
        For Each Cell In WorkRange
            If Cell.Text = TextLv2 Then
                InsertRow = Cell.Row
                Exit For
            End If
        Next Cell
        Return InsertRow
    End Function
    Sub AddLayer(LayerName As String)
        Dim Layer, Layer2 As AcadLayer
        On Error Resume Next
        Layer2 = acDoc.Layers.Item(LayerName)
        If Layer2.Name = "" Then
            For Each Layer In acDoc.Layers
                If StrComp(Layer.Name, LayerName, CompareMethod.Text) <> 0 Then
                    Layer2 = acDoc.Layers.Add(LayerName)
                    Exit For
                End If
            Next
        End If
        acDoc.ActiveLayer = Layer2
        acApp.Update()
    End Sub
    Function RemoveLeftString(InputText As String, ParamArray txt() As String) As String
        Dim temptxt, text As String
        Dim pos As Integer
        Dim len As Long
        Dim i As Integer
        text = InputText
        pos = InStr(txt(i), InputText, CompareMethod.Text)
        If pos > 0 Then
            For i = LBound(txt, 1) To UBound(txt, 1)
                pos = InStr(txt(i), InputText, CompareMethod.Text)
                len = txt(i).Length
                temptxt = Mid(InputText, pos + len)
                text = temptxt
            Next
        End If
        Return text
    End Function
    Public Function SplitText(text As String, pos As Integer) As String
        Dim Arr1(0 To 30) As String
        Dim arr2(0 To 30) As String
        Dim j As Long
        Dim item As String
        Arr1 = Regex.Split(text, "[^0-9.]")
        j = 1
        For Each item In Arr1
            If IsNumeric(item) Then
                arr2(j) = item
                j = j + 1
            End If
        Next
        Return arr2(pos)
    End Function
    Sub FindtextInAcad()
        Dim inputstring As String
        Dim acEnt As AcadEntity
        Dim sset As AcadSelectionSet
        Dim p1, p2
        On Error Resume Next
        sset = acDoc.SelectionSets.Add("TimText")
        acDoc.Utility.Prompt("Chọn vùng cần tim")
        sset.SelectOnScreen()
        inputstring = acDoc.Utility.GetString(True, "Text cần tìm: ")
        For Each acEnt In sset
            If TypeOf acEnt Is AcadText Or TypeOf acEnt Is AcadMText Then
                If LCase(acEnt.textstring) = LCase(inputstring) Or InStr(LCase(acEnt.textstring), LCase(inputstring), CompareMethod.Text) > 0 Then
                    acDoc.HandleToObject(acEnt.Handle)
                    acEnt.GetBoundingBox(p1, p2)
                    acApp.ZoomWindow(p1, p2)
                    Exit For
                End If
            End If
        Next
        sset.Delete()
    End Sub
    ' CAC THU TUC GIAI PHONG BO NHO
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub
    Sub ReleaseExcelObj()
        On Error Resume Next
        releaseObject(xlWS)
        releaseObject(xlWB)
        releaseObject(xlApp)
    End Sub
    Sub ReleaseAcadObj()
        On Error Resume Next
        releaseObject(acDoc)
        releaseObject(acApp)
    End Sub
End Class
