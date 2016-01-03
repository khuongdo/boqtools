Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports System.Text.RegularExpressions
Imports System.GC

Public Class BOQSlabTools
    Public xlApp As Application
    Public xlWB As Workbook
    Public xlWS As Worksheet
    Public acApp As AcadApplication
    Public acDoc As AcadDocument
    Public acEnt_kyhieuthep As AcadEntity
    Public Const PI As Double = Math.PI 'so PI
    Public SoDoanThepSan As Integer
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
    '   XU LY TRONG ACAD
    Sub AddLayer(LayerName As String)
        Dim Layer, Layer2 As AcadLayer
        On Error Resume Next
        Call StartACAD()
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
        Call ReleaseAcadObj()
    End Sub
    Function DrawRebar(InputColor As ACAD_COLOR, SuffixName As String, TextHeight As Double) As Double
        Dim SlabLine As AcadLine
        Dim Lrebar As Double = 0
        Dim pt1, pt2
        'Dim sodoan As Long = 0
        Dim Midpt(0 To 2) As Double
        Dim acText_SlabRebar As AcadText
        SoDoanThepSan = 0
        Call StartACAD()
        Try
            Do
                With acDoc.Utility
                    pt1 = CType(.GetPoint(, "Chọn điểm dầu thép sàn[ESC để vẽ khoảng rải]: "), Double())
                    pt2 = CType(.GetPoint(pt1, "Chọn điểm cuối thép sàn[ESC để vẽ khoảng rải]: "), Double())

                    SlabLine = acDoc.ModelSpace.AddLine(pt1, pt2)
                    SlabLine.Lineweight = ACAD_LWEIGHT.acLnWt060
                    SlabLine.color = InputColor
                    SlabLine.Update()
                    Lrebar = SlabLine.Length ' Gan chieu dai cay thep
                    ' Add text tren line
                    Midpt(0) = (SlabLine.EndPoint(0) - SlabLine.StartPoint(0)) / 2.0# + SlabLine.StartPoint(0)
                    Midpt(1) = (SlabLine.EndPoint(1) - SlabLine.StartPoint(1)) / 2.0# + SlabLine.StartPoint(1)
                    Midpt(2) = (SlabLine.EndPoint(2) - SlabLine.StartPoint(2)) / 2.0# + SlabLine.StartPoint(2)

                    xlWS = xlWB.Sheets("List")
                    Dim str As String = "SLABREBAR" & FindNextSlab() & SuffixName

                    acText_SlabRebar = acDoc.ModelSpace.AddText(str, Midpt, TextHeight)
                    acText_SlabRebar.color = ACAD_COLOR.acRed
                    Dim TextAngle As Double = SlabLine.Angle
                    'If TextAngle >= PI And Then TextAngle = TextAngle - PI
                    'acText_SlabRebar.Rotation = TextAngle
                    acText_SlabRebar.Update()
                    ' them handle text object vao excel
                    xlWS = xlWB.Sheets("slabdata")
                    xlWS.Range("XFD" & FindNextSlab()).Value = acText_SlabRebar.Handle
                End With
                SoDoanThepSan = SoDoanThepSan + 1
            Loop
        Catch ex As Exception
            If ex.Message Like "*0x80020009*" Then
                Return Lrebar
                Exit Try
            End If
        End Try
        Return Lrebar
        Call ReleaseAcadObj()
    End Function
    Function DrawSpace(InputColor As ACAD_COLOR) As Double
        Dim pt1, pt2
        Dim SpaceLine As AcadLine = Nothing
        Dim Lspace As Double = 0
        Call StartACAD()
        Do
            Try
                With acDoc.Utility
                    pt1 = CType(.GetPoint(, "Chọn điểm đầu khoảng rải [ECS để thoát]: "), Double())
                    pt2 = CType(.GetPoint(pt1, "Chọn điểm cuối khoảng rải [ECS để thoát]: "), Double())
                    SpaceLine = acDoc.ModelSpace.AddLine(pt1, pt2)
                    Lspace = Lspace + SpaceLine.Length
                    SpaceLine.Lineweight = ACAD_LWEIGHT.acLnWt060
                    SpaceLine.color = InputColor
                    SpaceLine.Update()
                End With
            Catch ex As Exception
                Exit Do
            End Try
        Loop
        Return Lspace
        Call ReleaseAcadObj()
    End Function
    Function SelectRebarType() As String
        Dim type As String = ""
        Call StartACAD()
        Try
            With acDoc.Utility
                Dim keywords As String = "1 2 3 4 5"
                .InitializeUserInput(0, keywords)
                Dim opt As String = .GetKeyword("Chọn loại thép" & vbCrLf & "1<mu-mu>/2<td-td>/3<mo-mo>/4<ne-ne>/5<ne-td>")
                Select Case opt
                    Case "1"
                        type = "mu-mu"
                    Case "2"
                        type = "td-td"
                    Case "3"
                        type = "mo-mo"
                    Case "4"
                        type = "ne-ne"
                    Case "5"
                        type = "ne-td"
                End Select
            End With
        Catch ex As Exception
            type = "mu-mu"
        End Try
        Return type
        Call ReleaseAcadObj()
    End Function
    Sub AddSlabLV3FromACAD_Acadfile1(TextListLV2 As String)
        On Error Resume Next
        Call StartACAD()
        Call StartExcel()
        Dim sset(0 To 2) As AcadSelectionSet
        Dim str_tenthep As String = ""
        Dim str_kyhieuthep As String
        Dim Lrebar As Double = 0
        Dim Lspace As Double = 0
        Dim rebartype As String
        For i As Integer = 0 To 2
            sset(i) = acDoc.SelectionSets.Add(i)
        Next
        ' cho nguoi dung nhap lieu tu acad
        With acDoc.Utility
            ' Nhap loai thep
            rebartype = CType(SelectRebarType(), String)
            ' Nhap ten thep
            .Prompt("Chọn tên thép [ESC để nhập tay]: ")
            sset(0).SelectOnScreen()
            str_tenthep = sset(0).Item(0).TextString
            If str_tenthep = "" Then
                str_tenthep = .GetString(False, "Nhập tên: ")
            End If
            ' Nhap ky hieu thep
            .Prompt("Chọn ký hiệu thép: ")
            sset(1).SelectOnScreen()
            acEnt_kyhieuthep = sset(1).Item(0)
            str_kyhieuthep = sset(1).Item(0).TextString
        End With
        Dim fi As Integer = CType(SplitText(str_kyhieuthep, 1), Integer)
        Dim space As Double = CType(SplitText(str_kyhieuthep, 2), Double)
        ' ADD LAYER
        Call AddLayer("BOQTools(c)DHK_ThepSan")
        ' Ve chieu dai thep
        Lrebar = DrawRebar(ACAD_COLOR.acMagenta, "", acEnt_kyhieuthep.Height)
        ' Ve khoang phan bo thep
        Lspace = DrawSpace(ACAD_COLOR.acBlue)
        If Lspace = 0 Or Lrebar = 0 Then
            MsgBox("Có lỗi xảy ra!Vui lòng nhập lại", vbCritical, "BOQ-Tools(c)DHK")
            For i = 0 To 2
                sset(i).Delete()
            Next
            GC.Collect()
            Exit Sub
        End If
        ' XUAT DU LIEU SANG EXCEL
        xlApp.ScreenUpdating = False
        Dim NextSlab As Long = FindNextSlab()
        xlWS = xlWB.Sheets("slabdata")
        With xlWS
            .Range("A" & NextSlab).Value = "SLABREBAR" & NextSlab ' Ten cay thep
            .Range("B" & NextSlab).Value = str_tenthep ' Ki hieu cay thep
            .Range("C" & NextSlab).Value = fi 'DK thep
            .Range("D" & NextSlab).Value = space ' @ thep
            .Range("E" & NextSlab).Value = Lrebar ' Chieu dai cay thep
            .Range("F" & NextSlab).Value = Lspace ' chieu dai doan rai thep
            .Range("G" & NextSlab).FormulaR1C1 = "=RC[4]*(ROUND(RC[-1]/RC[-3],0)+1)"
            .Range("H" & NextSlab).Value = rebartype
            .Range("K" & NextSlab).Value = SoDoanThepSan
        End With
        xlWS = xlWB.Sheets("dulieu_insert")
        Select Case rebartype
            Case "mu-mu"
                xlWS.Range("A64").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("64").Copy()
            Case "mo-mo"
                xlWS.Range("A65").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("65").Copy()
            Case "td-td"
                xlWS.Range("A66").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("66").Copy()
            Case "ne-ne"
                xlWS.Range("A67").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("67").Copy()
            Case "ne-td"
                xlWS.Range("A68").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("68").Copy()
        End Select
        xlWS = xlWB.Sheets("Reinforcement")
        xlWS.Activate()
        xlWS.Range("A" & FindInsertRowSlab("Reinforcement", TextListLV2)).Insert(XlDirection.xlDown)
        xlApp.CutCopyMode = False
        xlApp.ScreenUpdating = True
        ' XOA DATA SELECTION SET
        For i = 0 To 2
            sset(i).Delete()
        Next
        Call ReleaseAcadObj()
        Call ReleaseExcelObj()
    End Sub
    Sub AddSlabLV3FromACAD_Acadfile2(VitriThep As Integer, VitriTenThep As Integer, KytuTach As String, TextListLV2 As String)
        On Error Resume Next
        Call StartExcel()
        Call StartACAD()
        Dim sset(2) As AcadSelectionSet
        Dim str_kyhieuthep As String = ""
        Dim Lrebar As Double = 0
        Dim Lspace As Double = 0
        Dim rebartype As String
        Dim fi As Integer
        Dim space As Double
        Dim i As Integer
        For i = 0 To 2
            sset(i) = acDoc.SelectionSets.Add("SLABREBAR" & i)
        Next
        ' CHO NGUOI DUNG NHAP LIEU
        '' CHON LOAI THEP
        rebartype = CType(SelectRebarType(), String)
        ' NHAP KY HIEU THEP
        acDoc.Utility.Prompt("Chọn tên thép [ESC để nhập tay]: ")
        sset(0).SelectOnScreen()
        acEnt_kyhieuthep = sset(0).Item(0)
        str_kyhieuthep = acEnt_kyhieuthep.TextString
        ' ADD LAYER
        Call AddLayer("BOQTools(c)DHK_ThepSan")
        ' VE CHIEU DAI CAY THEP
        Lrebar = DrawRebar(ACAD_COLOR.acMagenta, "", acEnt_kyhieuthep.Height)
        ' VE KHOANG PHAN BO THEP
        Lspace = DrawSpace(ACAD_COLOR.acBlue)
        If Lspace = 0 Or Lrebar = 0 Then
            MsgBox("Có lỗi xảy ra!Vui lòng nhập lại", vbCritical, "BOQ-Tools(c)DHK")
            For i = 0 To 2
                sset(i).Delete()
            Next
            Call ReleaseAcadObj()
            Call ReleaseExcelObj()
            Exit Sub
        End If
        ' XUAT DU LIEU SANG EXCEL
        xlWS = xlWB.Sheets("Input")

        Dim RebarArray1 As Range = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("R")).SpecialCells(XlCellType.xlCellTypeConstants)
        Dim kyhieubanve(), KyHieuThep, TenThep, Thep As String
        kyhieubanve = Split(str_kyhieuthep, KytuTach)
        KyHieuThep = kyhieubanve(VitriThep - 1) ' Neu tren BV A-T50 --> thi KQ = A
        ' Tim kiem ky hieu tuong ung
        For Each cell As Range In RebarArray1 ' Tim ky hieu tuong ung trong excel voi ky hieu lay duoc tren acad
            If cell.Text = KyHieuThep Then
                Thep = cell.Offset(0, 1).Text
                Exit For
            End If
        Next
        TenThep = kyhieubanve(VitriTenThep - 1)
        fi = SplitText(Thep, 1)
        space = SplitText(Thep, 2)

        xlApp.ScreenUpdating = False
        Dim NextSlab As Long = FindNextSlab()
        xlWS = xlWB.Sheets("slabdata")
        With xlWS
            .Range("A" & NextSlab).Value = "SLABREBAR" & NextSlab ' Ten cay thep
            .Range("B" & NextSlab).Value = TenThep ' Ki hieu cay thep
            .Range("C" & NextSlab).Value = fi 'DK thep
            .Range("D" & NextSlab).Value = space ' @ thep
            .Range("E" & NextSlab).Value = Lrebar ' Chieu dai cay thep
            .Range("F" & NextSlab).Value = Lspace ' chieu dai doan rai thep
            .Range("G" & NextSlab).FormulaR1C1 = "=RC[4]*(ROUND(RC[-1]/RC[-3],0)+1)"
            .Range("H" & NextSlab).Value = rebartype
            .Range("K" & NextSlab).Value = SoDoanThepSan
        End With
        xlWS = xlWB.Sheets("dulieu_insert")
        Select Case rebartype
            Case "mu-mu"
                xlWS.Range("A64").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("64").Copy()
            Case "mo-mo"
                xlWS.Range("A65").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("65").Copy()
            Case "td-td"
                xlWS.Range("A66").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("66").Copy()
            Case "ne-ne"
                xlWS.Range("A67").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("67").Copy()
            Case "ne-td"
                xlWS.Range("A68").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("68").Copy()
        End Select
        xlWS = xlWB.Sheets("Reinforcement")
        xlWS.Activate()
        xlWS.Range("A" & FindInsertRowSlab("Reinforcement", TextListLV2)).Insert(XlDirection.xlDown)
        xlApp.CutCopyMode = False
        xlApp.ScreenUpdating = True
        ' XOA DATA SELECTION SET
        For i = 0 To 2
            sset(i).Delete()
        Next
        ' GIAI PHONG BO NHO
        Call ReleaseAcadObj()
        Call ReleaseExcelObj()
    End Sub
    Sub FindTextInAcad()
        Dim ent As AcadEntity
        Dim pt1, pt2
        Call StartACAD()
        Call StartExcel()
        Try
            pt1 = Nothing : pt2 = Nothing
            Dim handleID As String = Nothing
            Dim ActiveValue As String
            ActiveValue = xlApp.ActiveCell.Value
            xlWS = xlWB.Sheets("slabdata")
            Dim WorkRange As Excel.Range = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("A"))
            For Each Cell As Excel.Range In WorkRange
                If TypeOf Cell.Value Is String Then
                    If CType(Cell.Value, String) = ActiveValue Then
                        handleID = CType(xlWS.Range("XFD" & Cell.Row).Value, String)
                    End If
                End If
            Next
            ent = acDoc.HandleToObject(handleID)
            ent.GetBoundingBox(pt1, pt2)
            acApp.ZoomWindow(pt1, pt2)
        Catch ex As Exception
            MsgBox("Đối tượng không tồn tại hoặc bản vẽ hiện hành không đúng", vbCritical, "BOQ-Tool@DHK")
            Exit Try
        End Try
        Call ReleaseAcadObj()
        Call ReleaseExcelObj()
    End Sub
    Sub AddThepCauTao(kyhieu As String, TxtHeight As Double, TextListLV2 As String)
        Call StartACAD()
        Call StartExcel()
        Dim rebartype As String = ""
        Dim fi As Integer = 0
        Dim space As Double = 0
        Dim Lrebar As Double = 0
        Dim Lspace As Double = 0
        Call AddLayer("BOQTool(c)DHK_ThepCauTao")
        rebartype = CType(SelectRebarType(), String) ' Chon loai thep
        fi = SplitText(kyhieu, 1)
        space = SplitText(kyhieu, 2)
        Lrebar = DrawRebar(ACAD_COLOR.acYellow, "-CT", TxtHeight)
        Lspace = DrawSpace(ACAD_COLOR.acCyan)
        ' XUAT DU LIEU SANG EXCEL
        xlApp.ScreenUpdating = False

        Dim NextSlab As Long = FindNextSlab()
        xlWS = xlWB.Sheets("slabdata")
        With xlWS
            .Range("A" & NextSlab).Value = "SLABREBAR" & NextSlab ' Ten cay thep
            .Range("B" & NextSlab).Value = "Cấu tạo" ' Ki hieu cay thep
            .Range("C" & NextSlab).Value = fi 'DK thep
            .Range("D" & NextSlab).Value = space ' @ thep
            .Range("E" & NextSlab).Value = Lrebar ' Chieu dai cay thep
            .Range("F" & NextSlab).Value = Lspace ' chieu dai doan rai thep
            .Range("G" & NextSlab).FormulaR1C1 = "=RC[4]*(ROUND(RC[-1]/RC[-3],0)+1)"
            .Range("H" & NextSlab).Value = rebartype
            .Range("K" & NextSlab).Value = SoDoanThepSan
        End With
        xlWS = xlWB.Sheets("dulieu_insert")
        Select Case rebartype
            Case "mu-mu"
                xlWS.Range("A64").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("64").Copy()
            Case "mo-mo"
                xlWS.Range("A65").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("65").Copy()
            Case "td-td"
                xlWS.Range("A66").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("66").Copy()
            Case "ne-ne"
                xlWS.Range("A67").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("67").Copy()
            Case "ne-td"
                xlWS.Range("A68").Value = "SLABREBAR" & NextSlab
                xlWS.Rows("68").Copy()
        End Select
        xlWS = xlWB.Sheets("Reinforcement")
        xlWS.Activate()
        xlWS.Range("A" & FindInsertRowSlab("Reinforcement", TextListLV2)).Insert(XlDirection.xlDown)
        xlApp.CutCopyMode = False
        xlApp.ScreenUpdating = True
        ' Giai phong bo nho
        Call ReleaseAcadObj()
        Call ReleaseExcelObj()
    End Sub
    ' Xu ly tren excel
    Sub AcquireSlabOutline()
        Dim NextRow As Long
        Dim top, bot, j, i As Long
        Dim countlv1, countlv2 As Long
        Dim Cell, WorkRange As Range
        Dim RangeTop(0 To 100) As Range
        Dim RangeBot(0 To 100) As Range
        Dim Listlv1(0 To 100), Listlv2(0 To 1000) As String
        Dim WriteRange As Range
        Dim sheet_Outline As Worksheet
        Dim RangeLV2 As Range
        Dim LastRow As Long
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        xlWS = xlWB.Sheets("Reinforcement")
        '   Xoa du lieu cu
        sheet_Outline = xlWB.Sheets("List")
        sheet_Outline.Range("D1").CurrentRegion.ClearContents()
        '   Xac dinh va ghi dau muc level 1
        WriteRange = sheet_Outline.Range("D1")
        WorkRange = xlApp.Intersect(xlWS.Columns("R"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas) ' Column chua thong tin Beam LV1
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
            NextRow = sheet_Outline.Range("D" & sheet_Outline.Rows.Count).End(XlDirection.xlUp).Row + 1
            sheet_Outline.Cells(NextRow, WriteRange.Column) = Listlv1(j)
        Next j
        '   Xac dinh dau muc level 2
        countlv2 = 0
        For j = 1 To top
            RangeLV2 = xlWS.Range(RangeTop(j).Offset(0, 1), RangeBot(j).Offset(0, 1)).SpecialCells(XlCellType.xlCellTypeFormulas)
            For Each Cell In RangeLV2
                If Left(Cell.Text, 1) <> "/" And Cell.Text <> "" Then
                    countlv2 = countlv2 + 1
                    Listlv2(countlv2) = Listlv1(j) & "/" & Cell.Text
                End If
            Next Cell
        Next j
        '   Ghi dau muc level 2 vao sheet
        For j = 1 To countlv2
            NextRow = sheet_Outline.Range("E" & sheet_Outline.Rows.Count).End(XlDirection.xlUp).Row + 1
            sheet_Outline.Cells(NextRow, WriteRange.Column + 1) = Listlv2(j)
        Next j
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()

    End Sub
    Sub AddSlabLV1()
        Dim NextRow As Long
        Dim temptxt As String
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        xlWS = xlWB.Sheets("dulieu_insert")
        temptxt = InputBox("Nhap ten: ", "khuong.do")
        If Not temptxt = "" Then
            xlWS.Range("A56").Value = temptxt
            xlWS.Range("A72").Value = temptxt
        Else : Exit Sub
        End If
        'Chen vao sheet Rein
        xlWS = xlApp.Sheets("Reinforcement")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("56:57").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        'Chen vao sheet Concrete
        xlWS = xlApp.Sheets("Concrete")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("72:73").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        'Chen vao sheet formwork
        xlWS = xlApp.Sheets("Formwork")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("72:73").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub
    Sub AddSlabLV2()
        Dim TextLv1 As String
        Dim Cell, WorkRange As Range
        Dim StartRow, EndRow As Long
        Dim temptxt As String
        Dim xlws_dulieuInsert As Worksheet
        ' Khoi dong excel
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        temptxt = InputBox("Nhập tên LV2: ", "khuong.do")
        xlws_dulieuInsert = xlWB.Sheets("dulieu_insert")
        If Not temptxt = "" Then
            xlws_dulieuInsert.Range("A60").Value = temptxt
            xlws_dulieuInsert.Range("A74").Value = temptxt ' ghi du lieu thep san hien tai vao sheet insert
        Else : Exit Sub
        End If
        'lay du lieu dam LV1 hien tai ------------------
        xlWS = xlWB.Sheets("list")
        TextLv1 = CType(xlWS.Range("XFB2").Value, String)
        'Tim vi tri bat dau va ket thuc Beam LV1 sheet Reinforcement------------------
        xlWS = xlWB.Sheets("Reinforcement")
        xlWS.Activate()
        WorkRange = xlApp.Intersect(xlWS.Columns("R"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas)
        For Each Cell In WorkRange
            Select Case Cell.Value
                Case TextLv1
                    StartRow = Cell.Row
                Case "/" & TextLv1
                    EndRow = Cell.Row
                    Exit For
            End Select
        Next Cell
        'Chen du lieu vao sheet Reinforcement---------------
        xlws_dulieuInsert.Rows("60:61").Copy()
        xlWS.Range("A" & EndRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        'Tim vi tri bat dau va ket thuc Beam LV1 sheet Concrete----------
        xlWS = xlWB.Sheets("Concrete")
        xlWS.Activate()
        WorkRange = xlApp.Intersect(xlWS.Columns("O"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas)
        For Each Cell In WorkRange
            Select Case Cell.Value
                Case TextLv1
                    StartRow = Cell.Row
                Case "/" & TextLv1
                    EndRow = Cell.Row
                    Exit For
            End Select
        Next Cell
        ' Chen beam LV2 vao sheet Reinforcement
        xlws_dulieuInsert.Rows("74:75").Copy()
        xlWS.Range("A" & EndRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        '---------------------------
        'Tim vi tri bat dau va ket thuc Beam LV1 sheet Formwork----------
        xlWS = xlWB.Sheets("Formwork")
        xlWS.Activate()
        WorkRange = xlApp.Intersect(xlWS.Columns("O"), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeFormulas)

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
        xlws_dulieuInsert.Rows("74:75").Copy()
        xlWS.Range("A" & EndRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        '---------------------------
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub
    ' CAC THU TUC HO TRO
    Function FindInsertRowSlab(SheetName As String, TextListLV2 As String) As Long
        Dim TextLv1, TextLv2 As String
        Dim Cell, WorkRange As Range
        Dim StartCellLv1, EndCellLv1 As Range
        Dim CurrSlab As Long
        Dim InsertRow As Long
        Dim xlWS_List = xlWB.Sheets("List")
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        CurrSlab = FindNextSlab() - 1
        If CurrSlab <= 4 Then CurrSlab = 5
        '-----------------------------
        '   Lay Current Index cua beam LV1,LV2
        Dim arr1() As String
        ReDim arr1(2)

        arr1 = Split(TextListLV2, "/")
        TextLv1 = arr1(0)
        TextLv2 = arr1(1)

        xlWS = xlWB.Sheets(SheetName)
        ' Tim vung LV1
        Dim WR As Range = xlApp.Intersect(xlWS.Columns("R"), xlWS.UsedRange)
        WorkRange = WR.SpecialCells(XlCellType.xlCellTypeFormulas)

        For Each Cell In WorkRange
            Select Case Cell.Text
                Case TextLv1
                    StartCellLv1 = Cell
                Case "/" & TextLv1
                    EndCellLv1 = Cell
            End Select
        Next Cell
        ' Tim vung LV2
        WR = xlWS.Range(StartCellLv1.Offset(0, 1), EndCellLv1.Offset(0, 1))
        WorkRange = WR.SpecialCells(XlCellType.xlCellTypeFormulas)
        TextLv2 = "/" & TextLv2
        For Each Cell In WorkRange
            If Cell.Text = TextLv2 Then
                InsertRow = Cell.Row
                Exit For
            End If
        Next Cell
        Return InsertRow
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Function
    Function FindNextSlab() As Long
        Call StartExcel()
        Dim NextSlab As Long
        xlWS = xlWB.Sheets("slabdata")
        NextSlab = xlWS.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        If NextSlab <= 2 Then NextSlab = 3
        Return NextSlab
        Call ReleaseExcelObj()
    End Function
    Public Function SplitText(text As String, pos As Integer)
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
