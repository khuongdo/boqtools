Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports System.Text.RegularExpressions
Public Class BOQWallTools
    Public acApp As AcadApplication
    Public acDoc As AcadDocument
    Public xlApp As Application
    Public xlWB As Workbook
    Public xlWS As Worksheet
    Public objBOQOtherTools As New BOQOtherTools
    Public Const AppName = "BOQTools(c)DHK"
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
    Sub AddWallLV1()
        Dim NextRow As Long
        Dim temptxt As String
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        xlWS = xlWB.Sheets("dulieu_insert")
        temptxt = InputBox("Nhập ten LV1: ", AppName)
        If Not temptxt = "" Then
            xlWS.Range("A84").Value = temptxt
            xlWS.Range("A87").Value = temptxt
        Else : Exit Sub
        End If
        'Chen vao sheet Rein
        xlWS = xlApp.Sheets("Reinforcement")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("84:85").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        'Chen vao sheet Concrete
        xlWS = xlApp.Sheets("Concrete")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("87:88").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        'Chen vao sheet formwork
        xlWS = xlApp.Sheets("Formwork")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("87:88").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        xlWB.Sheets("Reinforcement").Activate()
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub
    Sub AddWallLV2(TextLV1 As String)
        Dim Cell, WorkRange As Range
        Dim StartRow, EndRow As Long
        Dim temptxt As String
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        temptxt = InputBox("Nhập tên LV2: ", AppName)
        If Not temptxt = "" Then
            xlWB.Sheets("dulieu_insert").Range("A90").Value = temptxt
            xlWB.Sheets("dulieu_insert").Range("A93").Value = temptxt
        Else : Exit Sub
        End If
        'lay du lieu dam LV1 hien tai tu sheet LIST------------------

        'Tim vi tri bat dau va ket thuc Beam LV1 sheet Reinforcement------------------
        xlWS = xlWB.Sheets("Reinforcement")
        xlWS.Activate()
        WorkRange = xlApp.Intersect(xlWS.Columns("T"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas) ' column chua lv1
        For Each Cell In WorkRange
            If Cell.Text <> "" Then
                If Cell.Text = "/" & TextLv1 Then
                    EndRow = Cell.Row
                    Exit For
                End If
            End If
        Next
        xlWB.Sheets("dulieu_insert").Rows("90:91").Copy() ' Chen du lieu vao sheet Rein
        xlWS.Range("A" & EndRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        'Tim vi tri bat dau va ket thuc Beam LV1 sheet Concrete----------
        xlWS = xlWB.Sheets("Concrete")
        xlWS.Activate()
        WorkRange = xlApp.Intersect(xlWS.Columns("Q"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas) ' column chua lv1
        For Each Cell In WorkRange
            Select Case Cell.Text
                Case TextLv1
                    StartRow = Cell.Row
                Case "/" & TextLv1
                    EndRow = Cell.Row
                    Exit For
            End Select
        Next Cell
        ' Chen LV2 vao sheet concrete
        xlWB.Sheets("dulieu_insert").Rows("93:94").Copy()
        xlWS.Range("A" & EndRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        '---------------------------
        'Tim vi tri bat dau va ket thuc Beam LV1 sheet Formwork----------
        xlWS = xlWB.Sheets("Formwork")
        xlWS.Activate()
        WorkRange = xlApp.Intersect(xlWS.Columns("Q"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas)

        For Each Cell In WorkRange
            Select Case Cell.Value
                Case TextLv1
                    StartRow = Cell.Row
                Case "/" & TextLv1
                    EndRow = Cell.Row
                    Exit For
            End Select
        Next Cell

        ' Chen LV2 vao sheet Formwork
        xlWB.Sheets("dulieu_insert").Rows("93:94").Copy()
        xlWS.Range("A" & EndRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        '---------------------------
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub
    ' THU TUC XU LY TRONG ACAD
    Sub AddWallLV3(str_SlabThk As String, Htang As Double, str_WallLV2 As String)
        Dim str_tenvach As String
        Dim sset1(4) As AcadSelectionSet
        Dim i As Integer
        Dim opt1, opt2 As String
        Dim acEnt As AcadEntity
        Dim actext As AcadEntity
        Dim dbl_wallthk As Double = 0
        Dim dbl_Lwall As Double = 0
        Dim xlWs_insert As Worksheet
        Call StartACAD()
        Call StartExcel()
        On Error Resume Next
        xlWs_insert = xlWB.Sheets("dulieu_insert")
        ' XOA HET DU LIEU TREN SHEET ADDBEAMLV3
        xlWS = xlWB.Sheets("addbeamlv3")
        xlWS.UsedRange.EntireRow.Delete(Shift:=XlDirection.xlUp)
        For i = 0 To 4
            sset1(i) = acDoc.SelectionSets.Add("Wall" & i)
        Next
        ' CHON LOAI VACH
        opt1 = "1"
        acDoc.Utility.InitializeUserInput(0, "1 2")
        opt1 = acDoc.Utility.GetKeyword("Chọn loại vách: 1<Vách chữ nhật> 2<Vách bất kì>")
        Select Case opt1
            Case "1"
                ' CHON TEN VACH
                acDoc.Utility.Prompt("Chọn tên vách: ")
                sset1(0).SelectOnScreen()
                actext = sset1(0).Item(0)
                str_tenvach = sset1(0).Item(0).textstring
                ' CHON CHIEU DAY VACH
                acDoc.Utility.Prompt("Chọn chiều dày vách: ")
                opt2 = "1"
                'acDoc.Utility.InitializeUserInput(0, "1 2")
                'opt2 = acDoc.Utility.GetKeyword("Chiều dày vách:1<Chọn DIM> 2<Nhập tay>")
                'Select Case opt2
                '    Case "1"
                acDoc.Utility.Prompt("Chọn dim: ")
                sset1(1).SelectOnScreen()
                acEnt = sset1(1).Item(0)
                If acEnt.TextOverride <> "" Then
                    dbl_wallthk = acEnt.TextOverride
                Else
                    dbl_wallthk = acEnt.Measurement
                End If
                'Case "2"
                '    dbl_wallthk = acDoc.Utility.GetReal("Nhập chiều dày(mm): ")
                '  End Select
                ' CHON CHIEU DAI VACH
                'opt2 = "1"
                'acDoc.Utility.Prompt("Chọn chiều dài vách: ")
                'acDoc.Utility.InitializeUserInput(0, "1 2")
                'opt2 = acDoc.Utility.GetKeyword("Chiều dài vách:1<Chọn DIM> 2<Vẽ line>")
                'Select Case opt2
                '    Case "1"
                acDoc.Utility.Prompt("Chọn dim: ")
                sset1(2).SelectOnScreen()
                dbl_Lwall = 0
                For Each acEnt In sset1(2)
                    If acEnt.TextOverride <> "" Then
                        dbl_Lwall = dbl_Lwall + acEnt.TextOverride
                    Else
                        dbl_Lwall = dbl_Lwall + acEnt.Measurement
                    End If
                Next
                '    Case "2"
                'End Select
        ' ADD TRUOC HANG DAU TIEN VAO SHEET ADDBEAMLV3
        xlWs_insert.Rows("106").Copy()
        xlWS = xlWB.Sheets("addbeamlv3")
        xlWS.Activate()
        xlWS.Range("A1").Insert(XlDirection.xlDown)
        xlApp.CutCopyMode = False
        ' CHON THEP CHU VACH
        Call NhapThepChuVach(Htang)
        'CHON THEP DAI VACH
        Call NhapThepDai(Htang)
        'CHON THEP C
        Call NhapThepC(Htang)
        ' NHAP HANG KET THUC
        Dim LastRow As Long
        xlWS = xlWB.Sheets("addbeamlv3")
        xlWs_insert.Rows(113).copy()
        LastRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlWS.Range("A" & LastRow).Insert()
        xlApp.CutCopyMode = False
        ' HOI NGUOI DUNG CO NHAP WALL HAY KO
        acDoc.Utility.InitializeUserInput(0, "Yes No")
        opt2 = "Yes"
        opt2 = acDoc.Utility.GetKeyword("Nhập vách này Yes[No]: ")
        If opt2 = "No" Then
            GoTo end_sub
        End If
        ' XUAT SANG SHEET WALL DATA
        xlWS = xlWB.Sheets("walldata")
        Dim nextrow As Long
        nextrow = xlWS.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        With xlWS
            .Range("A" & nextrow).Value = "WALL" & nextrow
            .Range("B" & nextrow).Value = str_tenvach
            .Range("D" & nextrow).Value = dbl_wallthk
            .Range("E" & nextrow).Value = dbl_Lwall
            .Range("F" & nextrow).Value = Htang
            .Range("G" & nextrow).Value = str_SlabThk
            .Range("H" & nextrow).FormulaR1C1 = "=VLOOKUP(RC[-1],SlabsThkTable,2,0)"
        End With
        End Select
        Dim p1, p2
        actext.GetBoundingBox(p1, p2)
        acDoc.ModelSpace.AddLine(p1, p2)
        Call ExportToExcel(str_WallLV2)
end_sub:
        For i = 0 To 4
            sset1(i).Delete()
        Next
        Call ReleaseAcadObj()
        Call ReleaseExcelObj()
    End Sub
    Sub NhapThepChuVach(Htang As Double)
        Dim sset(0 To 3) As AcadSelectionSet
        Dim i As Integer
        Dim str_kyhieuthep As String
        Dim str_loaithep As String
        Dim str_sohieuthep As String
        Dim fi As Integer = 0
        Dim Num As Integer = 0
        Dim L As Double = 0
        Dim NextRow As Long
        Dim opt As String
        Dim sl As Integer
        Dim acText As AcadText
        Try
            Do
                For i = 0 To 3
                    sset(i) = acDoc.SelectionSets.Add("ThepChuVach" & i)
                Next
                ' Chon loai thep
                With acDoc.Utility
                    .InitializeUserInput(0, "1 2 3")
                    opt = .GetKeyword("Chọn loại thép: 1<no-no> 2<td-no> 3<co-co> 4<ne-co>")
                    Select Case opt
                        Case 1
                            str_loaithep = "no-no"
                        Case 2
                            str_loaithep = "td-no"
                        Case 3
                            str_loaithep = "co-co"
                        Case 4
                            str_loaithep = "ne-co"
                    End Select
                End With
                ' CHON TEN THEP
                acDoc.Utility.Prompt("Chọn số hiệu thép chủ vách [ESC để bỏ qua]: ")
                sset(0).SelectOnScreen()
                If Not sset(0).Item(0) Is Nothing Then str_sohieuthep = sset(0).Item(0).TextString

                ' CHON KY HIEU THEP
                acDoc.Utility.Prompt("Chọn ký hiệu thép chủ vách [ESC để thoát] : ")
                sset(1).SelectOnScreen()
                acText = sset(1).Item(0)
                acText.color = ACAD_COLOR.acRed
                str_kyhieuthep = acText.TextString
                fi = CType(objBOQOtherTools.SplitText(str_kyhieuthep, 1), Integer)
                sl = CType(objBOQOtherTools.SplitText(str_kyhieuthep, 2), Integer)
                '   ADD VAO EXCEL
                xlWS = xlWB.Sheets("dulieu_insert")
                Dim R As Long
                Select Case str_loaithep
                    Case "no-no"
                        R = 107
                    Case "td-no"
                        R = 114
                    Case "co-co"
                        R = 109
                    Case "ne-co"
                        R = 110
                End Select
                xlWS.Range("A" & R).Value = str_sohieuthep
                xlWS.Range("F" & R).Value = objBOQOtherTools.SplitText(str_kyhieuthep, 2) ' Fi
                xlWS.Range("G" & R).Value = objBOQOtherTools.SplitText(str_kyhieuthep, 1) ' SL
                xlWS.Rows(R).Copy()
                xlWS = xlWB.Sheets("addbeamlv3")
                xlWS.Activate()
                NextRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                xlWS.Range("A" & NextRow).Insert()
                xlApp.CutCopyMode = False
                xlWS.Range("C" & NextRow).Value = "=B1"
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
    Sub NhapThepC(Htang As Double)
        Dim sset(2) As AcadSelectionSet
        Dim dbl_khoangRaiC As Double
        Dim str_thepc As String
        Dim int_fiC As Integer
        Dim int_SThepC As Integer
        Dim WthepC As Double
        Dim actext As AcadEntity
        Try
            Do
                ' NHAP TRONG ACAD
                For i As Integer = 0 To 2
                    sset(i) = acDoc.SelectionSets.Add("ThepC" & i)
                Next
                With acDoc.Utility
                    .Prompt("Chọn ký hiệu thép C: ")
                    sset(0).SelectOnScreen()
                    actext = sset(0).Item(0)
                    actext.color = ACAD_COLOR.acRed
                    str_thepc = sset(0).Item(0).TextString
                    int_fiC = CType(objBOQOtherTools.SplitText(str_thepc, 1), Integer)
                    int_SThepC = CType(objBOQOtherTools.SplitText(str_thepc, 2), Integer)
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
                End With
                ' XUAT SANG EXCEL
                xlWS = xlWB.Sheets("dulieu_insert")
                xlWS.Range("C112").FormulaR1C1 = "=" & WthepC & "-2*wa" ' chieu dai thep c
                xlWS.Range("F112").Value = int_fiC ' fi thep c
                xlWS.Range("G112").FormulaR1C1 = "=ROUND(" & Htang & "*" & dbl_khoangRaiC & "/(" & int_SThepC & "*" & int_SThepC & "),0)+1"
                xlWS.Rows("112").Copy()
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
    Sub NhapThepDai(H As Double)
        Dim sset(3) As AcadSelectionSet
        Dim str_sohieu As String
        Dim str_kyhieu As String
        Dim int_fiC As Integer
        Dim int_Sdai As Integer
        Dim dbl_b As Double
        Dim dbl_h As Double
        Dim acText As AcadEntity
        Try
            Do
                ' NHAP TRONG ACAD
                For i As Integer = 0 To 3
                    sset(i) = acDoc.SelectionSets.Add("ThepDaiVach" & i)
                Next
                With acDoc.Utility
                    .Prompt("Chọn số hiệu thép đai: ")
                    sset(3).SelectOnScreen()
                    str_sohieu = sset(3).Item(0).TextString
                    .Prompt("Chọn ký hiệu thép đai: ")
                    sset(0).SelectOnScreen()
                    acText = sset(0).Item(0)
                    acText.color = ACAD_COLOR.acRed
                    str_kyhieu = sset(0).Item(0).TextString
                    int_fiC = CType(objBOQOtherTools.SplitText(str_kyhieu, 1), Integer)
                    int_Sdai = CType(objBOQOtherTools.SplitText(str_kyhieu, 2), Integer)
                    .Prompt("Chọn DIM BxH thép đai: ")
                    sset(1).SelectOnScreen()
                    dbl_b = sset(1).Item(0).Measurement
                    dbl_h = sset(1).Item(1).Measurement
                End With
                ' XUAT SANG EXCEL
                xlWS = xlWB.Sheets("dulieu_insert")
                xlWS.Range("A111").Value = str_sohieu
                xlWS.Range("B111").FormulaR1C1 = "=" & dbl_b & "-2*wa" ' Rong thep dai
                xlWS.Range("C111").FormulaR1C1 = "=" & dbl_h & "-2*wa" ' dai thep dai
                xlWS.Range("F111").Value = int_fiC ' 
                xlWS.Range("G111").FormulaR1C1 = "=ROUND(" & H & "/" & int_Sdai & ",0)+1"
                xlWS.Rows("111").Copy()
                xlWS = xlWB.Sheets("addbeamlv3")
                Dim LastRow As Long
                LastRow = xlWS.Range("C" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
                xlWS.Range("A" & LastRow).Insert()
                xlApp.CutCopyMode = False
                'XOA SELECTION SET
                For i As Integer = 0 To 3
                    sset(i).Delete()
                Next
            Loop
        Catch ex As Exception
            For i As Integer = 0 To 3
                sset(i).Delete()
            Next
            Exit Try
        End Try
    End Sub
    Sub ExportToExcel(str_WallVL2 As String)
        Dim InsertRow, CurrRow As Long
        On Error Resume Next
        CurrRow = FindNextWall() - 1 ' Vi thu tuc nay goi sau khi da add du lieu vao sheet beamdata nen phai -1
        xlWS = xlWB.Sheets("addbeamlv3") : xlWS.Range("A1").Value = "WALL" & CurrRow
        With xlWB.Sheets("dulieu_insert") ' Ghi ten dam vao sheet dulieu_insert truoc khi insert 
            .Range("A121") = "WALL" & CurrRow
            .Range("A123") = "WALL" & CurrRow
        End With
        ' Chen dam LV3 vao sheet Concrete
        xlWS = xlWB.Sheets("Concrete")
        InsertRow = FindInsertRow("Concrete", "Q", str_WallVL2)
        xlWB.Sheets("dulieu_insert").Rows(121).Copy()
        xlWS.Activate()
        xlWS.Range("A" & InsertRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        ' Chen dam LV3 vao sheet Formwork
        xlWS = xlWB.Sheets("Formwork")
        InsertRow = FindInsertRow("Formwork", "Q", str_WallVL2)
        xlWB.Sheets("dulieu_insert").Rows(123).Copy()
        xlWS.Activate()
        xlWS.Range("A" & InsertRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        ' Chen dam LV3 vao sheet Reinforcement
        xlWB.Sheets("addbeamlv3").UsedRange.EntireRow.Copy()
        xlWS = xlWB.Sheets("Reinforcement")
        xlWS.Activate()
        InsertRow = FindInsertRow("Reinforcement", "T", str_WallVL2)
        xlWS.Range("A" & InsertRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
    End Sub
    ' CAC THU TUC HO TRO
    Function FindNextWall() As Long
        Dim NextRow As Long
        xlWS = xlWB.Sheets("walldata")
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
        CurrBeam = FindNextWall() - 1
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
        sheet_Outline.Range("G1").CurrentRegion.ClearContents()
        '   Xac dinh dau muc LV1 trong sheet Reinforcement va Ghi vao Sheet List
        WriteRange = sheet_Outline.Range("G1")
        WorkRange = xlApp.Intersect(xlWS.Columns("Q"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas) ' Column chua thong tin Beam LV1
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
            NextRow = sheet_Outline.Range("G" & sheet_Outline.Rows.Count).End(XlDirection.xlUp).Row + 1
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
        '   Ghi dau muc level 2 vao sheet list
        For j = 1 To countlv2
            NextRow = sheet_Outline.Range("H" & sheet_Outline.Rows.Count).End(XlDirection.xlUp).Row + 1
            sheet_Outline.Range("H" & NextRow).Value = Listlv2(j)
        Next j
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
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
