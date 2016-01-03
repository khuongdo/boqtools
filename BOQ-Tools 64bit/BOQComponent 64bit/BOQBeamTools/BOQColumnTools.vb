Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports System.Text.RegularExpressions

Public Class BOQColumnTools
    Public acApp As AcadApplication
    Public acDoc As AcadDocument
    Public xlApp As Application
    Public xlWB As Workbook
    Public xlWS As Worksheet
    Public Const AppName = "BOQTools(c)DHK"
    'THU TUC GOI ACAD EXCEL 
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
    ' THU TUC XU LY TREN EXCEL 
    Sub AcquireColumnOutline()
        Dim NextRow As Long
        Dim top, bot, j, i As Long
        Dim countlv1, countlv2 As Long
        Dim Cell, WorkRange As Range
        Dim RangeTop(100) As Range
        Dim RangeBot(100) As Range
        Dim Listlv1(100), Listlv2(100) As String
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
        sheet_Outline.Range("H1").CurrentRegion.ClearContents()
        '   Xac dinh va ghi dau muc level 1
        WriteRange = sheet_Outline.Range("H1")
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
            NextRow = sheet_Outline.Range("H" & sheet_Outline.Rows.Count).End(XlDirection.xlUp).Row + 1
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
            NextRow = sheet_Outline.Range("I" & sheet_Outline.Rows.Count).End(XlDirection.xlUp).Row + 1
            sheet_Outline.Cells(NextRow, WriteRange.Column + 1) = Listlv2(j)
        Next j
        xlApp.ScreenUpdating = True
        Call StartExcel()
    End Sub
    Sub AddColumnLV1()
        Dim NextRow As Long
        Dim temptxt As String
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        xlWS = xlWB.Sheets("dulieu_insert")
        temptxt = InputBox("Nhap ten: ", "khuong.do")
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
        'Chen vao sheet Concrete
        xlWS = xlApp.Sheets("Concrete")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("87:88").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        'Chen vao sheet formwork
        xlWS = xlApp.Sheets("Formwork")
        xlWS.Activate()
        NextRow = xlApp.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        xlApp.Sheets("dulieu_insert").Rows("87:88").Copy()
        xlWS.Range("A" & NextRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub
    Sub AddColumnLV2()
        Dim TextLv1 As String
        Dim Cell, WorkRange As Range
        Dim StartRow, EndRow As Long
        Dim temptxt As String
        On Error Resume Next
        Call StartExcel()
        xlApp.ScreenUpdating = False
        temptxt = InputBox("Nhập tên LV2: ", AppName)
        xlWS = xlWB.Sheets("dulieu_insert")
        If Not temptxt = "" Then
            xlWS.Range("A84").Value = temptxt
            xlWS.Range("A87").Value = temptxt
        Else : Exit Sub
        End If
        'lay du lieu dam LV1 hien tai tu sheet LIST------------------
        xlWS = xlWB.Sheets("List")
        TextLv1 = xlWS.Range("XFB3").Text
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
        ' Chen beam LV2 vao sheet concrete
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

        ' Chen beam LV3 vao sheet Formwork
        xlWB.Sheets("dulieu_insert").Rows("93:94").Copy()
        xlWS.Range("A" & EndRow).EntireRow.Insert(Shift:=XlDirection.xlDown)
        xlApp.CutCopyMode = False
        '---------------------------
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub
    ' THU TUC GIAI PHONG BO NHO
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
