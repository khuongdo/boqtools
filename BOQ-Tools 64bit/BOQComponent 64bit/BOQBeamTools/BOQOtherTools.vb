Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

Public Class BOQOtherTools
    Public xlApp As Application
    Public xlWB As Workbook
    Public xlWS As Worksheet
    Public acApp As AcadApplication
    Public acDoc As AcadDocument
    ' THU TUC KHOI DONG 
    Sub StartExcel()
        Try
            xlApp = GetObject(, "Excel.Application")
            xlWB = xlApp.ActiveWorkbook
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
    End Sub
    Sub StartACAD()
        Try
            acApp = GetObject(, "Autocad.Application")
            acDoc = acApp.ActiveDocument
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    ' CAC THU TUC HO TRO TRONG ACAD
    Sub FindtextInAcad()
        Call StartACAD()
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
        ReleaseAcadObj()
    End Sub
    Sub CopyText()
        Dim sset As AcadSelectionSet
        Dim tempEnt As AcadEntity
        Dim acText() As AcadEntity
        Dim x(), y() As Double
        Dim dbl_temp As Double
        Dim AlignmentPoint1(0 To 2), AlignmentPoint2(0 To 2) As Double
        Dim i, j As Integer
        Dim opt As String

        Try
            Call StartACAD()
            Call StartExcel()
            sset = acDoc.SelectionSets.Add("CopyText")
            acDoc.Utility.Prompt("Chọn text cần copy: ")
            sset.SelectOnScreen()
            If sset.Count = 0 Then Throw New IO.IOException
            ReDim acText(0 To sset.Count)
            ReDim x(0 To sset.Count)
            ReDim y(0 To sset.Count)
            i = 0
            For Each ent As AcadEntity In sset
                If TypeOf ent Is AcadText Or TypeOf ent Is AcadMText Then
                    acText(i) = ent
                    x(i) = ent.TextAlignmentPoint(0)
                    y(i) = ent.TextAlignmentPoint(1)
                    i = i + 1
                End If
            Next
            '' TEXT TRONG ACAD NAM TREN 1 COT
            For i = 0 To sset.Count - 1
                For j = i + 1 To sset.Count
                    If y(i) < y(j) Then
                        dbl_temp = y(i)
                        y(i) = y(j)
                        y(j) = dbl_temp
                        tempEnt = acText(i)
                        acText(i) = acText(j)
                        acText(j) = tempEnt
                    End If
                Next
            Next
            acDoc.Utility.InitializeUserInput(0, "Row Column")
            opt = acDoc.Utility.GetKeyword("Row[Column]: ")
            Select Case opt
                Case "Row"
                    For i = 0 To sset.Count
                        xlApp.ActiveCell.Offset(0, i).Value = acText(i).TextString
                    Next
                Case "Column"
                    For i = 0 To sset.Count
                        xlApp.ActiveCell.Offset(i, 0).Value = acText(i).TextString
                    Next
            End Select
        Catch ex As IO.IOException
            Exit Try
        Catch ex As Exception
        End Try
        sset.Delete()
        Call ReleaseAcadObj()
        Call ReleaseExcelObj()
    End Sub
    ' CAC THU TUC HO TRO KHAC
    Sub AutoSubtotal(Text As String, FormulaColumn As String, CodeCol As String)
        Dim WorkRange As Range
        Dim Formula As String
        Dim EndSubt As String, StartSubt As String
        Dim Cell As Range
        Dim CellArr() As Range
        Dim FirstRow() As Long, LastRow() As Long
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Dim count As Long
        Call StartExcel()
        xlWS = xlWB.ActiveSheet
        xlApp.ScreenUpdating = False
        '   Xac dinh vung lam viec
        WorkRange = xlApp.Intersect(xlWS.Columns(CodeCol), xlWS.UsedRange).SpecialCells(XlCellType.xlCellTypeConstants)
        '   Xac dinh so luong phan tu mang CellArr
        i = 0
        j = 0
        StartSubt = Text
        EndSubt = "/" & Text
        For Each Cell In WorkRange
            If Cell.Text = StartSubt Then
                count = count + 1
            End If
        Next Cell
        If count = 0 Then Exit Sub
        ReDim CellArr(0 To count)
        ReDim FirstRow(0 To count)
        ReDim LastRow(0 To count)
        For Each Cell In WorkRange
            Select Case Cell.Text
                Case StartSubt
                    i = i + 1
                    CellArr(i) = xlWS.Range(FormulaColumn & Cell.Row)
                    FirstRow(i) = Cell.Row
                Case EndSubt
                    j = j + 1
                    LastRow(j) = Cell.Row
            End Select
        Next Cell
        '   Chen cong thuc subtotal
        For k = 1 To count
            Formula = "=SUBTOTAL(9,R[1]C:R[" & (LastRow(k) - FirstRow(k)) & "]C)"
            CellArr(k).FormulaR1C1 = Formula
        Next k
        xlApp.ScreenUpdating = True
    End Sub
    Sub ClearContent(SheetName As String, StartRow As Integer)
        On Error Resume Next
        Call StartExcel()
        xlWS = xlWB.Sheets(SheetName)
        xlWS.Activate()
        Dim WorkRange As Range
        Dim Del_Range As Range
        Dim Cell As Range
        Del_Range = xlWS.Range("A" & StartRow, "A" & xlWS.Rows.Count).EntireRow
        Del_Range.Delete(Shift:=XlDirection.xlUp)
        Call ReleaseExcelObj()
    End Sub
    Sub ClearShapes()
        On Error Resume Next
        xlWS = xlWB.Sheets("Reinforcement")
        For Each Sh As Shape In xlWS.Shapes
            Sh.Delete()
        Next
    End Sub
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
    Sub CountObjectAcad()
        Dim count As Long
        Dim ent As AcadEntity
        Dim sset As AcadSelectionSet
        Dim text As String
        Dim p1, p2
        Dim acline As AcadLine
        Call StartACAD()

        Try
            sset = acDoc.SelectionSets.Add("CountObject")
            acDoc.Utility.Prompt("Chọn vùng tìm kiếm")
            sset.SelectOnScreen()
            text = acDoc.Utility.GetString(True, "Ký tự cần đếm: ")
            count = 0
            For Each ent In sset
                If TypeOf ent Is AcadText Or TypeOf ent Is AcadMText Then
                    If StrComp(LCase(ent.textstring), LCase(text), CompareMethod.Text) = 0 Then
                        count = count + 1
                        ent.GetBoundingBox(p1, p2)
                        acline = acDoc.ModelSpace.AddLine(p1, p2)
                        acline.Lineweight = ACAD_LWEIGHT.acLnWt060
                        acline.color = ACAD_COLOR.acRed
                    End If
                End If
            Next
            acDoc.Utility.Prompt(text & "=" & count)
        Catch ex As Exception
            Exit Try
        End Try
        sset.Delete()
        Call ReleaseAcadObj()

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
