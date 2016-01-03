Imports Microsoft.Office.Interop.Excel
Imports System.GC
Public Class SlabThick
    Public xlApp As Application
    Public xlWB As Workbook
    Public xlWS As Worksheet
    Private Sub StartExcel()
        Try
            xlApp = GetObject(, "Excel.Application")
            xlWB = xlApp.ActiveWorkbook
        Catch ex As Exception
            MsgBox("Có lỗi xảy ra")
        End Try
    End Sub
    Private Sub SlabThick_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call StartExcel()
        On Error Resume Next
        xlWS = xlWB.Sheets("Input")
        Dim workrange As Range = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("O"))
        For Each cell As Range In workrange
            If Not cell.Value = "" And cell.Value <> "Slabs" Then
                ListBox_Slaplist.Items.Add(cell.Value)
            End If
        Next
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button_close.Click
        Close()
        Call uf_Menu_ReinConc.LoadSlabOutline()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button_addslabthk.Click
        Call StartExcel()
        On Error Resume Next
        Dim NextRow As Long
        xlWS = xlWB.Sheets("Input")
        NextRow = xlApp.Range("O" & xlWS.Rows.Count).End(XlDirection.xlUp).Row + 1
        If NextRow = 1 Then NextRow = 2
        If TextBox_addslabthk.Text = "" Or TextBox_thickness.Text = "" Then ' thoat neu nguoi dung nhap thieu
            MsgBox("Thiếu dự liệu đầu vào", vbCritical, "BOQTools(c)DHK")
            Exit Sub
        Else
            ListBox_Slaplist.Items.Add(TextBox_addslabthk.Text)
            ListBox_Slaplist.Update()
            xlWS.Range("O" & NextRow).Value = TextBox_addslabthk.Text ' Nhap ten slab vao sheet slabthk
            xlWS.Range("P" & NextRow).Value = TextBox_thickness.Text ' Nhap be day san tuong ung
            ' Dim Name As String = "SLABTHK_" & TextBox_addslabthk.Text ' Ten cua Name add vao
            'xlApp.Names.Item(Name).Delete()
            ' xlApp.Names.Add(Name, xlWS.Range("P" & NextRow), True)
        End If
        GC.Collect()
    End Sub
    Private Sub ListBox_Slaplist_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox_Slaplist.SelectedValueChanged
        Call StartExcel()
        On Error Resume Next
        xlWS = xlWB.Sheets("Input")
        Dim CurrSlab As String = CType(ListBox_Slaplist.SelectedItem, String) ' Ten cua slab hien tai
        Dim workrange As Range = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("O"))
        For Each cell As Range In workrange
            If cell.Value = CurrSlab Then
                TextBox_thickness.Text = cell.Offset(0, 1).Value
                TextBox_addslabthk.Text = CurrSlab
                Exit Sub
            End If
        Next
        GC.Collect()
    End Sub
    Private Sub Button_RenameSlabsTHK_Click(sender As Object, e As EventArgs) Handles Button_modify.Click
        Call StartExcel()
        xlWS = xlWB.Sheets("Input")
        Dim CurrItem As String = ListBox_Slaplist.SelectedItem
        Dim workrange As Range = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("O"))
        For Each cell As Range In workrange
            If cell.Value = CurrItem Then
                cell.Offset(0, 1).Value = TextBox_thickness.Text
                Exit Sub
            End If
        Next
        GC.Collect()
    End Sub
    Private Sub Button_RemoveSlabTHK_Click(sender As Object, e As EventArgs) Handles Button_RemoveSlabTHK.Click
        Call StartExcel()
        On Error Resume Next
        Dim CurrItem As String = ListBox_Slaplist.SelectedItem
        ListBox_Slaplist.Items.Remove(CurrItem)
        xlWS = xlWB.Sheets("Input")
        Dim workrange As Range = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("O"))
        For Each cell As Range In workrange
            If cell.Value = CurrItem Then
                xlWS.Range("O" & cell.Row, "P" & cell.Row).Delete(Shift:=XlDirection.xlUp)
                Exit Sub
            End If
        Next
        GC.Collect()
    End Sub
    Private Sub Button_rename_Click(sender As Object, e As EventArgs) Handles Button_rename.Click
        Call StartExcel()
        Dim NewItem As String = TextBox_addslabthk.Text
        Dim OldItem As String = ListBox_Slaplist.SelectedItem
        Dim CurrIndex As Integer = -1
        If StrComp(NewItem, OldItem, CompareMethod.Text) = 0 Then
            MsgBox("Tên mới phải khác tên cũ", vbOKOnly, "BOQTools(c)DHK")
            Exit Sub
        Else
            CurrIndex = ListBox_Slaplist.SelectedIndex
            ListBox_Slaplist.Items.RemoveAt(CurrIndex)
            ListBox_Slaplist.Items.Insert(CurrIndex, NewItem)
            Dim workrange As Range = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("O"))
            For Each cell As Range In workrange
                If cell.Value = OldItem Then
                    cell.Value = NewItem
                    Dim OldItemName As String = "SLABTHK_" & OldItem ' Xoa name ref cu
                    xlApp.Names.Item(OldItemName).Delete()
                    Dim NewItemName As String = "SLABTHK_" & NewItem ' Them name ref moi
                    xlApp.Names.Add(NewItemName, xlWS.Range("P" & cell.Row), True)
                    Exit For
                End If
            Next
        End If
        GC.Collect()
    End Sub
End Class