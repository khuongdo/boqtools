Imports Microsoft.Office.Interop.Excel
Imports BOQ_Components
Public Class uf_AddBeamLV3
    Public xlApp As Application
    Public xlWB As Workbook
    Public xlWS As Worksheet
    Private Sub cmb_cancle_Click(sender As Object, e As EventArgs) Handles cmb_cancle.Click
        Close()
    End Sub
    Private Sub uf_AddBeamLV3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost() = True
        tb_beamname.Focus()
    End Sub
    Private Sub cmb_addbeamlv3_Click(sender As Object, e As EventArgs) Handles cmb_addbeamlv3.Click
        Dim CurrRow As Long
        Dim obj1 As New BOQBeamTools
        Dim obj2 As New BOQOtherTools
        On Error Resume Next
        xlApp = GetObject(, "Excel.Application") : xlWB = xlApp.ActiveWorkbook : xlWS = xlWB.Sheets("beamdata")
        CurrRow = obj1.FindNextBeam
        'Nhap du lieu vao sheet beamdata
        With xlWS
            .Cells(CurrRow, 1) = "BEAM" & CurrRow
            .Cells(CurrRow, 2) = tb_beamname.Text
            .Cells(CurrRow, 3) = tb_b.Text
            .Cells(CurrRow, 4) = tb_h.Text
            .Cells(CurrRow, 5) = tb_l.Text
            .Range("F" & CurrRow).FormulaR1C1 = "=RC[-1]-RC[3]" ' Tinh L be tong
            .Cells(CurrRow, 7) = tb_sck.Text
            .Cells(CurrRow, 8) = tb_sonhip.Text
            .Cells(CurrRow, 9) = tb_col.Text
            ' Nhap thep chu tren
            .Cells(CurrRow, 10) = obj2.SplitText(tb_ttren.Text, 1)
            .Cells(CurrRow, 11) = obj2.SplitText(tb_ttren.Text, 2)
            ' Nhap thep chu duoi
            .Cells(CurrRow, 12) = obj2.SplitText(tb_tduoi.Text, 1)
            .Cells(CurrRow, 13) = obj2.SplitText(tb_tduoi.Text, 2)
            ' Nhap thep dai
            Dim Formula As String
            Dim s1, s2 As Double
            s1 = obj2.SplitText(tb_tdai.Text, 2)
            s2 = obj2.SplitText(tb_tdai.Text, 3)
            Formula = "=ROUND(RC[-8]/AVERAGE(" & s1 & "," & s2 & "),0)+1"
            .Range("N" & CurrRow).FormulaR1C1 = Formula
            .Cells(CurrRow, 15) = obj2.SplitText(tb_tdai.Text, 1)
            ' Nhap thep tang cuong goi
            .Cells(CurrRow, 16) = obj2.SplitText(tb_tctren.Text, 1)
            .Cells(CurrRow, 17) = obj2.SplitText(tb_tctren.Text, 2)
            ' Nhap thep tang cuong nhip
            .Cells(CurrRow, 18) = obj2.SplitText(tb_tcduoi.Text, 1)
            .Cells(CurrRow, 19) = obj2.SplitText(tb_tcduoi.Text, 2)
            ' Nhap thep gia
            .Cells(CurrRow, 20) = obj2.SplitText(tb_tgia.Text, 1)
            .Cells(CurrRow, 21) = obj2.SplitText(tb_tgia.Text, 2)
            ' Nhap dai C
            s1 = obj2.SplitText(tb_tc.Text, 2)
            Formula = "=ROUND(RC[-16]/" & s1 & ",0)+1"
            .Range("V" & CurrRow).FormulaR1C1 = Formula
            .Cells(CurrRow, 23) = obj2.SplitText(tb_tgia.Text, 1)
            ' Nhap Chieu day san
            .Cells(CurrRow, 24) = uf_Menu_ReinConc.ComboBox_SlabThkList.SelectedItem
            .Range("Y" & CurrRow).FormulaR1C1 = "=VLOOKUP(RC[-1],SlabsThkTable,2,0)"
        End With
        obj1.CalculateBeamLV3_1(CType(uf_Menu_ReinConc.combobox_beamlv2.Text, String)) ' Goi thu tuc add beam lv3
        ' Xoa du lieu trong textbox
        tb_beamname.Text = Nothing
        tb_b.Text = Nothing
        tb_h.Text = Nothing
        tb_sck.Text = Nothing
        tb_sonhip.Text = Nothing
        tb_l.Text = Nothing
        tb_col.Text = Nothing
        tb_ttren.Text = Nothing
        tb_tduoi.Text = Nothing
        tb_tdai.Text = Nothing
        tb_tctren.Text = Nothing
        tb_tcduoi.Text = Nothing
        tb_tgia.Text = Nothing
        tb_tc.Text = Nothing
        xlWB.Sheets("Reinforcement").activate()

        tb_beamname.Focus()
        xlApp.ScreenUpdating = True
        GC.Collect()
    End Sub
    
End Class