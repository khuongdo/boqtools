Imports Microsoft.Office.Interop.Excel

Public Class spec_beam
    Public xlApp As Application
    Public xlWB As Workbook
    Public xlWS As Worksheet
    Sub StartExcel()
        On Error Resume Next
        xlApp = GetObject(, "Excel.Application")
        xlWB = xlApp.ActiveWorkbook
        If Err.Number = 91 Then
            MsgBox("Bạn chưa mở file Excel BOQ Tools hoặc file không đúng", MsgBoxStyle.Critical, "Có lỗi xảy ra")
            Me.Close()
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged ' Thay doi khoang vuon goi
        Call StartExcel()
        xlWS = xlWB.Sheets("Input")
        xlWS.Range("V2").Value = TextBox1.Text
        Call ReleaseExcelObj()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub
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
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Call StartExcel()
        xlWS = xlWB.Sheets("Input")
        xlWS.Range("V3").Value = TextBox2.Text
        Call ReleaseExcelObj()
    End Sub
    Private Sub spec_beam_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call StartExcel()
        xlWS = xlWB.Sheets("Input")
        TextBox1.Text = xlWS.Range("V2").Value
        TextBox2.Text = xlWS.Range("V3").Value
        Call ReleaseExcelObj()
    End Sub
End Class