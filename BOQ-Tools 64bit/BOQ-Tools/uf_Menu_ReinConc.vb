Imports Microsoft.Office.Interop.Excel
Imports BOQ_Components
Imports System.Runtime.InteropServices

Public Class uf_Menu_ReinConc
    'WINDOWS API
    Public Const MOD_ALT As Integer = &H1 'Alt key
    Public Const MOD_SHIFT As Integer = &H4 ' SHIFT Key
    Public Const WM_HOTKEY As Integer = &H312
    Public ObjSlabTool As New BOQSlabTools
    Public objBeamTool As New BOQBeamTools
    <DllImport("User32.dll")> _
    Public Shared Function RegisterHotKey(ByVal hwnd As IntPtr, _
                        ByVal id As Integer, ByVal fsModifiers As Integer, _
                        ByVal vk As Integer) As Integer
    End Function
    <DllImport("User32.dll")> _
    Public Shared Function UnregisterHotKey(ByVal hwnd As IntPtr, _
                        ByVal id As Integer) As Integer
    End Function
    ' Khai bao bien tong the
    Public xlApp As Application
    Public xlWB As Workbook
    Public xlWS As Worksheet
    Public Const AppName As String = "BOQTool(c)DHK"
    'KHAI BAO BIEN CHO KIEU NHAP SLAB ACADFILE 2
    Public tHeight As Double
    Public ky_hieu As String
    ' KHOI DONG EXCEL,AUTOCAD
    Sub StartExcel()
        On Error Resume Next
        xlApp = GetObject(, "Excel.Application")
        xlWB = xlApp.ActiveWorkbook
        If Err.Number = 91 Then
            MsgBox("Bạn chưa mở file Excel BOQ Tools hoặc file không đúng", MsgBoxStyle.Critical, "Có lỗi xảy ra")
            Me.Close()
        End If
    End Sub
    ' XU LY SU KIEN OPEN FORM
    Private Sub CheckBox_useHotkey_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_useHotkey.CheckedChanged
        If CheckBox_useHotkey.Checked Then
            RegisterHotKey(Me.Handle, 1, MOD_ALT, Keys.S)
            RegisterHotKey(Me.Handle, 2, MOD_SHIFT, Keys.S)
            RegisterHotKey(Me.Handle, 3, MOD_SHIFT, Keys.D)
        Else
            UnregisterHotKey(Me.Handle, 1)
            UnregisterHotKey(Me.Handle, 2)
            UnregisterHotKey(Me.Handle, 3)
        End If
    End Sub
    Private Sub uf_Menu_Load(sender As Object, e As EventArgs) Handles MyBase.Load ' load cac gia tri trong file excel vao phan mem
        On Error Resume Next
        '' KHAI BAO SU KIEN KEY PRESS
        CheckBox_useHotkey.Checked = True
        RegisterHotKey(Me.Handle, 1, MOD_ALT, Keys.S)
        RegisterHotKey(Me.Handle, 2, MOD_SHIFT, Keys.S)
        RegisterHotKey(Me.Handle, 3, MOD_SHIFT, Keys.D)
        RegisterHotKey(Me.Handle, 4, MOD_SHIFT, Keys.F)
        RegisterHotKey(Me.Handle, 5, MOD_SHIFT, Keys.W)
        '-----
        Call StartExcel() ' Kiem tra file excel
        'GAN GIA TRI BIEN
        tHeight = xlWB.Sheets("Input").Range("U5").Value
        ky_hieu = xlWB.Sheets("Input").Range("U6").Value
        '----
        Call UpdateSpecItem()
        Call LoadBeamOutline()
        Call LoadSlabOutline()
        Call LoadWallOutline()
        Call UpdateSlabTHK()
        Call UpdateSheetInput()
        With ComboBox_inputtypeBeam.Items ' them du lieu vao combobox tab beam
            .Add("UserForm")
            .Add("Acadfile 1")
            .Add("Acadfile 2")
            .Add("Tính bê tông")
            .Add("PDF Converted")
        End With
        With ComboBox_InputTypeSlab.Items ' tham du lieu vao combobox tab slab
            .Add("UserForm")
            .Add("Ký hiệu: A@B")
            .Add("Ký hiếu: X-Y")
            .Add("PDF Converted")
            .Add("Thép cấu tạo")
        End With
        xlWS = xlWB.Sheets("List")
        combobox_beamlv1.SelectedIndex = xlWS.Range("XEZ1").Value ' Lay index hien tai beamlv1
        combobox_beamlv2.SelectedIndex = xlWS.Range("XFA1").Value ' Lay index hien tai beamlv2
        '   KHOA TEXTBOX DUOI DAY
        TextBox_RebarNamePos.Enabled = False
        TextBox_symbolpos.Enabled = False
        TextBox_kytuphantach.Enabled = False
        xlApp.ScreenUpdating = True
        GC.Collect()
    End Sub
    Private Sub UpdateSlabTHK()
        Call StartExcel()
        On Error Resume Next
        ' Nhap du lieu vao combo box slab list
        ComboBox_SlabThkList.Items.Clear()
        xlWS = xlWB.Sheets("Input")
        Dim cell As Range
        Dim workrange As Range = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("O")).SpecialCells(XlCellType.xlCellTypeConstants)
        For Each cell In workrange
            If cell.Value <> "" And cell.Value <> "Slabs" Then
                ComboBox_SlabThkList.Items.Add(cell.Value)
            End If
        Next
        ComboBox_SlabThkList.SelectedIndex = 0
        GC.Collect()
    End Sub
    Sub UpdateSheetInput()
        On Error Resume Next
        Call StartExcel()
        xlWS = xlWB.Sheets("Input")
        'TextBox_symbolpos.Text = xlWS.Range("U2").Value
        'TextBox_RebarNamePos.Text = xlWS.Range("U3").Value
        'TextBox_kytuphantach.Text = xlWS.Range("U4").Value
       
        GC.Collect()
    End Sub
    ' SU KIEN HOTKEY
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = WM_HOTKEY Then
            Dim id As IntPtr = m.WParam
            Select Case (id.ToString)
                Case "1" ' Nhan to hop phim ALT+S de click Addslab LV3
                    Button_slablv3.PerformClick()
                Case "2" ' Nhan to hop phim Shift+S de click Addslab LV3
                    Button_slablv3.PerformClick()
                Case "3" ' Nhan to hop phim Shift+D de click beam lv3
                    cmb_beamlv3.PerformClick()
                Case "4" ' Nhan to hop phim Shift+F de click beam lv3
                    Button_findtextinAcad.PerformClick()
                Case "5" ' Nhan to hop phim Shift+F de click beam lv3
                    Button_wallLV3.PerformClick()
            End Select
        End If
        MyBase.WndProc(m)
    End Sub
    '  XU LY TAB BEAM
    Function FindNextBeam()
        Call StartExcel()
        xlWS = xlWB.Sheets("beamdata")
        Return xlWS.Range("A" & xlWS.Rows.Count).End(XlDirection.xlUp).Row
        GC.Collect()
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles cmb_beamlv1.Click 'Nut addbeam LV1
        objBeamTool.AddBeamLV1()
    End Sub ' click button add beam lv1
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles cmb_beamlv2.Click ' Nut addbeam LV2
        Dim obj As New BOQBeamTools
        obj.AddBeamLV2()
        GC.Collect()
    End Sub 'Click button add beam lv2
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles cmb_beamlv3.Click 'Nut addbeam lV3
        Select Case ComboBox_inputtypeBeam.SelectedIndex
            Case 0
                uf_AddBeamLV3.Show()
            Case 1
                objBeamTool.AddBeamLV3FromACAD_1(CType(combobox_beamlv2.SelectedItem, String), CType(ComboBox_SlabThkList.SelectedItem, String)) ' Chay thu tuc add beam tren acad
            Case 2
                objBeamTool.AddBeamLV3FromACAD_2(CType(combobox_beamlv2.SelectedItem, String), CType(ComboBox_SlabThkList.SelectedItem, String))
            Case 3
                objBeamTool.AddBeamLV3FromACAD_3(CType(combobox_beamlv2.SelectedItem, String), CType(ComboBox_SlabThkList.SelectedItem, String))
            Case Else
                MsgBox("Chưa chọn cách nhập", vbCritical, "BOQTools(c)DHK")
                Exit Sub
        End Select
        Call ReleaseExcelObj()
    End Sub ' Click button add beam lv3
    Private Sub combobox_beamlv1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles combobox_beamlv1.SelectedIndexChanged
        Call StartExcel()
        xlWS = xlWB.Sheets("List")
        ' ghi Gia tri nguoi dung chon vao sheet List
        With xlWS
            .Range("XEZ1").Value = combobox_beamlv1.SelectedIndex
            .Range("XFB1").Value = combobox_beamlv1.SelectedItem
        End With
        Call ReleaseExcelObj()
    End Sub
    Private Sub combobox_beamlv2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles combobox_beamlv2.SelectedIndexChanged
        Call StartExcel()
        xlWS = xlWB.Sheets("List")
        ' ghi Gia tri nguoi dung chon vao sheet List
        With xlWS
            .Range("XFA1").Value = combobox_beamlv2.SelectedIndex
            .Range("XFC1").Value = combobox_beamlv2.SelectedItem
        End With
        Call ReleaseExcelObj()
    End Sub
    Private Sub cmb_updatelist_Click(sender As Object, e As EventArgs) Handles cmb_updatelist.Click
        On Error Resume Next
        Dim OBJ As New BOQBeamTools
        OBJ.AcquireOutline() ' Lay outline tu sheet Rein export sang sheet list
        Call LoadBeamOutline() ' Load Outline tu sheet List
        MsgBox("Cập nhật thành công", vbOKOnly, "Thông báo")
        xlApp.ScreenUpdating = True
    End Sub ' Click Update list tab Beam
    Sub LoadBeamOutline()
        ' Update item trong 2 combobox beam lv1,lv2
        Dim BeamLV1, BeamLV2, Cell As Range
        Call StartExcel()
        xlWS = xlWB.Sheets("List")
        BeamLV1 = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("A"))
        BeamLV2 = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("B"))
        '   Nhap item vao combobox beamlv1 /tab beam
        If Not BeamLV1 Is Nothing Then
            combobox_beamlv1.Items.Clear()
            For Each Cell In BeamLV1
                If Cell.Text <> "" Then combobox_beamlv1.Items.Add(Cell.Text)
            Next
        Else
            Exit Sub
        End If
        '   Nhap item vao combobox beamlv2/ tab beam
        If Not BeamLV2 Is Nothing Then
            combobox_beamlv2.Items.Clear()
            For Each Cell In BeamLV2
                If Cell.Text <> "" Then combobox_beamlv2.Items.Add(Cell.Text)
            Next
        Else
            Exit Sub
        End If
        Call ReleaseExcelObj()
    End Sub
    '  XU LY TAB OTHER TOOL
    Private Sub cmb_exit_Click(sender As Object, e As EventArgs) Handles cmb_exit.Click 'Nut thoat chuong trinh
        Close()
    End Sub ' Click exit
    Private Sub cmb_autosubt_Click(sender As Object, e As EventArgs) Handles cmb_autosubt.Click
        Dim obj As New BOQOtherTools
        Call StartExcel()
        On Error Resume Next
        xlWS = xlApp.ActiveWorkbook.ActiveSheet
        Select Case xlWS.Name
            Case "Reinforcement"
                obj.AutoSubtotal("subtotal1", "N", "O")
                obj.AutoSubtotal("subtotal2", "N", "O")
                obj.AutoSubtotal("subtotal3", "N", "O")
            Case Else
                obj.AutoSubtotal("subtotal1", "K", "L")
                obj.AutoSubtotal("subtotal2", "K", "L")
        End Select
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub ' click autosubtotal
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles cmb_clearall.Click
        On Error Resume Next
        Dim Obj As New BOQOtherTools
        Call StartExcel()
        xlApp.ScreenUpdating = False
        If MsgBox("Tất cả dữ liệu sẽ bị xóa! Bạn có chắc chắn?", vbOKCancel, "Thông báo") = vbOK Then
            Obj.ClearContent("beamdata", 5)
            Obj.ClearContent("Formwork", 9)
            Obj.ClearContent("Concrete", 9)
            Obj.ClearContent("Reinforcement", 9)
            Obj.ClearContent("slabdata", 3)
            Obj.ClearShapes()
        Else
            Exit Sub
        End If
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
        Call releaseObject(Obj)
    End Sub ' click button xoa tat ca du lieu
    Private Sub Button_CountObjAcad_Click(sender As Object, e As EventArgs) Handles Button_CountObjAcad.Click
        Dim obj As New BOQOtherTools
        obj.CountObjectAcad()
        Call releaseObject(obj)
    End Sub
    ' XU LY TAB SPECS
    Private Sub Button_apply_Click(sender As Object, e As EventArgs) Handles Button_apply.Click
        Call StartExcel()
        xlApp.Range("fo").Value = TextBox_footing.Text
        xlApp.Range("col").Value = TextBox_col.Text
        xlApp.Range("grbe").Value = TextBox_grbeam.Text
        xlApp.Range("sl").Value = TextBox_slab.Text
        xlApp.Range("be").Value = TextBox_beam.Text
        xlApp.Range("wa").Value = TextBox_wall.Text
        xlApp.Range("anchorage").Value = TextBox_anchorage.Text
        xlApp.Range("lap").Value = TextBox_laps.Text
        xlApp.Range("hook").Value = TextBox_hook.Text
        MsgBox("Đổi thành công", vbOKOnly, "Thông báo")
        GC.Collect()
    End Sub ' ghi cac gia tri vao sheet specs
    Private Sub UpdateSpecItem()
        Call StartExcel()
        ' Load cai gia tri sheet specs vao textbox trong tab specs
        TextBox_footing.Text = xlApp.Range("fo").Value
        TextBox_col.Text = xlApp.Range("col").Value
        TextBox_grbeam.Text = xlApp.Range("grbe").Value
        TextBox_slab.Text = xlApp.Range("sl").Value
        TextBox_beam.Text = xlApp.Range("be").Value
        TextBox_wall.Text = xlApp.Range("wa").Value
        TextBox_anchorage.Text = xlApp.Range("anchorage").Value
        TextBox_laps.Text = xlApp.Range("laps").Value
        TextBox_hook.Text = xlApp.Range("hook").Value
        GC.Collect()
    End Sub ' Update tab specs
    Private Sub CheckBox_AlwaysOnTop_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_AlwaysOnTop.CheckedChanged
        Me.TopMost = False
        If CheckBox_AlwaysOnTop.Checked Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If
    End Sub ' checkbox always on top
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Dim obj As New BOQBeamTools
        Call StartExcel()
        obj.ZoomToObject()
        GC.Collect()
    End Sub ' Find text in acad
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        SlabThick.Show()
    End Sub
    ' XU LY TAB SLAB
    Sub LoadSlabOutline()
        Call StartExcel()
        On Error Resume Next
        ' Nhap du lieu vao combo box slab level
        ' Update item trong 2 combobox beam lv1,lv2
        Dim SlabLV1, SlabLV2 As Range
        xlWS = xlWB.Sheets("List")
        SlabLV1 = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("D")).SpecialCells(XlCellType.xlCellTypeConstants)
        SlabLV2 = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("E")).SpecialCells(XlCellType.xlCellTypeConstants)
        '   Nhap item vao combobox slab lv1 vao tab slab
        If Not SlabLV1 Is Nothing Then
            ComboBox_slablv1.Items.Clear()
            For Each Cell In SlabLV1
                If Cell.Text <> "" Then ComboBox_slablv1.Items.Add(Cell.Text)
            Next Cell
        Else
            Exit Sub
        End If
        '   Nhap item vao combobox beamlv2/ tab beam
        If Not SlabLV2 Is Nothing Then
            ComboBox_slablv2.Items.Clear()
            For Each Cell In SlabLV2
                If Cell.Text <> "" Then ComboBox_slablv2.Items.Add(Cell.Text)
            Next Cell
        Else
            Exit Sub
        End If
        Call ReleaseExcelObj()
    End Sub
    Private Sub Button_slablv1_Click(sender As Object, e As EventArgs) Handles Button_slablv1.Click
        Dim obj As New BOQSlabTools
        obj.AddSlabLV1()
        Call releaseObject(obj)
    End Sub
    Private Sub Button_slablv2_Click(sender As Object, e As EventArgs) Handles Button_slablv2.Click
        Dim obj As New BOQSlabTools
        obj.AddSlabLV2()
        releaseObject(obj)
    End Sub
    Private Sub Button_slablv3_Click(sender As Object, e As EventArgs) Handles Button_slablv3.Click
        On Error Resume Next
        Call StartExcel()
        Select Case ComboBox_InputTypeSlab.SelectedIndex
            Case 1
                ObjSlabTool.AddSlabLV3FromACAD_Acadfile1(CType(ComboBox_slablv2.SelectedItem, String))
                xlWS = xlWB.Sheets("slabdata")
                With xlWS
                    .Range("I" & ObjSlabTool.FindNextSlab - 1).Value = ComboBox_SlabThkList.SelectedItem
                    .Range("J" & ObjSlabTool.FindNextSlab - 1).FormulaR1C1 = "=VLOOKUP(RC[-1],SlabsThkTable,2,0)"
                End With
            Case 2
                ObjSlabTool.AddSlabLV3FromACAD_Acadfile2(CType(TextBox_symbolpos.Text, Integer), _
                                                         CType(TextBox_RebarNamePos.Text, Integer), _
                                                         TextBox_kytuphantach.Text, CType(ComboBox_slablv2.SelectedItem, String))

                xlWS = xlWB.Sheets("slabdata")
                With xlWS
                    .Range("I" & ObjSlabTool.FindNextSlab - 1).Value = ComboBox_SlabThkList.SelectedItem
                    .Range("J" & ObjSlabTool.FindNextSlab - 1).FormulaR1C1 = "=VLOOKUP(RC[-1],SlabsThkTable,2,0)"
                End With
            Case 4
                ObjSlabTool.AddThepCauTao(CType(TextBox_ThepCT.Text, String), _
                                          CType(TextBox_TextHeightThepCT.Text, Double), _
                                          CType(ComboBox_slablv2.SelectedItem, String))
                xlWS = xlWB.Sheets("slabdata")
                With xlWS
                    .Range("I" & ObjSlabTool.FindNextSlab - 1).Value = ComboBox_SlabThkList.SelectedItem
                    .Range("J" & ObjSlabTool.FindNextSlab - 1).FormulaR1C1 = "=VLOOKUP(RC[-1],SlabsThkTable,2,0)"
                End With
            Case Else
                MsgBox("Not Available")
        End Select
        Call releaseObject(ObjSlabTool)
        Call ReleaseExcelObj()
    End Sub
    Private Sub Button_UpdateListSlab_Click(sender As Object, e As EventArgs) Handles Button_UpdateListSlab.Click
        Dim obj As New BOQSlabTools : obj.AcquireSlabOutline() 'Doc ouline tu sheet reinforcement roi xuat ra sheet list

        Call LoadSlabOutline() ' Load Outline tu sheet List
        MsgBox("Cập nhật thành công", vbOKOnly, AppName)
        Call releaseObject(obj)
    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        On Error Resume Next
        Dim obj As New BOQSlabTools
        obj.FindTextInAcad()
        Call releaseObject(obj)
    End Sub ' NUT TIM KIEM DOI TUONG BEN ACAD
    Private Sub ComboBox_slablv1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_slablv1.SelectedIndexChanged
        Call StartExcel()
        xlWS = xlWB.Sheets("List")
        ' ghi Gia tri nguoi dung chon vao sheet List
        With xlWS
            .Range("XEZ2").Value = ComboBox_slablv1.SelectedIndex
            .Range("XFB2").Value = ComboBox_slablv1.SelectedItem
        End With
        Call ReleaseExcelObj()
    End Sub
    Private Sub ComboBox_slablv2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_slablv2.SelectedIndexChanged
        Call StartExcel()
        xlWS = xlWB.Sheets("List")
        ' ghi Gia tri nguoi dung chon vao sheet List
        With xlWS
            .Range("XFA2").Value = ComboBox_slablv2.SelectedIndex
            .Range("XFC2").Value = ComboBox_slablv2.SelectedItem
        End With
        Call ReleaseExcelObj()
    End Sub
    Private Sub ComboBox_InputTypeSlab_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_InputTypeSlab.SelectedIndexChanged
        ' CHO PHEP NHAP NHUNG GIA TRI DUOI DAY KHI KIEU CHON LA ACAD FILE 2
        If ComboBox_InputTypeSlab.SelectedIndex = 2 Then
            TextBox_RebarNamePos.Enabled = True
            TextBox_symbolpos.Enabled = True
            TextBox_kytuphantach.Enabled = True
        Else
            TextBox_RebarNamePos.Enabled = False
            TextBox_symbolpos.Enabled = False
            TextBox_kytuphantach.Enabled = False
        End If
    End Sub
    ' XU LY TAB WALL
    Sub LoadWallOutline()
        Call StartExcel()
        On Error Resume Next
        ' Nhap du lieu vao combo box slab level
        ' Update item trong 2 combobox beam lv1,lv2
        Dim WallLV1, WallLV2 As Range
        xlWS = xlWB.Sheets("List")
        WallLV1 = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("G"))
        WallLV2 = xlApp.Intersect(xlWS.UsedRange, xlWS.Columns("H"))
        '   Nhap item vao combobox slab lv1 vao tab slab
        If Not WallLV1 Is Nothing Then
            ComboBox_wallLV1.Items.Clear()
            For Each Cell In WallLV1
                If Cell.Text <> "" Then ComboBox_wallLV1.Items.Add(Cell.Text)
            Next Cell
        Else
            Exit Sub
        End If
        '   Nhap item vao combobox beamlv2/ tab beam
        If Not WallLV2 Is Nothing Then
            ComboBox_wallLV2.Items.Clear()
            For Each Cell In WallLV2
                If Cell.Text <> "" Then ComboBox_wallLV2.Items.Add(Cell.Text)
            Next Cell
        Else
            Exit Sub
        End If
        Call ReleaseExcelObj()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button_updatewalllist.Click
        Dim obj As New BOQWallTools
        obj.AcquireOutline()
        Call LoadWallOutline()
        MsgBox("Cập nhật thành công", vbOKOnly, AppName)
        Call releaseObject(obj)
    End Sub
    Private Sub ComboBox_wallLV1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_wallLV1.SelectedIndexChanged
        Call StartExcel()
        xlWS = xlWB.Sheets("List")
        xlWS.Range("XEZ3").Value = ComboBox_wallLV1.SelectedIndex
        xlWS.Range("XFB3").Value = ComboBox_wallLV1.SelectedItem
        Call ReleaseExcelObj()
    End Sub
    Private Sub ComboBox_wallLV2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_wallLV2.SelectedIndexChanged
        Call StartExcel()
        xlWS = xlWB.Sheets("List")
        xlWS.Range("XFA3").Value = ComboBox_wallLV2.SelectedIndex
        xlWS.Range("XEC3").Value = ComboBox_wallLV2.SelectedItem
        Call ReleaseExcelObj()
    End Sub
    Private Sub Button_wallLV1_Click(sender As Object, e As EventArgs) Handles Button_wallLV1.Click
        Dim obj As New BOQWallTools
        obj.AddWallLV1()
        Call releaseObject(obj)
    End Sub
    Private Sub Button_wallLV2_Click(sender As Object, e As EventArgs) Handles Button_wallLV2.Click
        Dim obj As New BOQWallTools
        obj.AddWallLV2(ComboBox_wallLV1.Text)
        Call releaseObject(obj)
    End Sub
    Private Sub Button_wallLV3_Click(sender As Object, e As EventArgs) Handles Button_wallLV3.Click
        Dim obj As New BOQWallTools
        On Error Resume Next
        obj.AddWallLV3(ComboBox_SlabThkList.SelectedItem, CType(TextBox_FLHeight.Text, Double), ComboBox_wallLV2.SelectedItem)
        Call releaseObject(obj)
    End Sub
    ' SU KIEN CLOSE FORM
    Private Sub uf_Menu_ReinConc_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        On Error Resume Next
        UnregisterHotKey(Me.Handle, 1)
        UnregisterHotKey(Me.Handle, 2)
        UnregisterHotKey(Me.Handle, 3)
        UnregisterHotKey(Me.Handle, 4)
        UnregisterHotKey(Me.Handle, 5)
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
    Private Sub ComboBox_inputtypeBeam_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_inputtypeBeam.SelectedIndexChanged
        If ComboBox_InputTypeSlab.SelectedItem = "Acad file 1" Then
            spec_beam.Show()
        End If
    End Sub
    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button_findtextinAcad.Click
        Dim obj As New BOQOtherTools
        obj.FindtextInAcad()
        releaseObject(obj)
    End Sub
    Private Sub Button4_Click_2(sender As Object, e As EventArgs) Handles Button4.Click
        Dim obj As New BOQOtherTools
        obj.CopyText()
    End Sub

   
 
    Private Sub Button_AnSheet_Click(sender As Object, e As EventArgs) Handles Button_AnSheet.Click
        On Error Resume Next

        Dim sh As Worksheet
        Call StartExcel()
        xlApp.ScreenUpdating = False
        sh = xlWB.Sheets("list")
        sh.Visible = XlSheetVisibility.xlSheetVeryHidden
        sh = xlWB.Sheets("addbeamlv3")
        sh.Visible = XlSheetVisibility.xlSheetVeryHidden
        sh = xlWB.Sheets("dulieu_insert")
        sh.Visible = XlSheetVisibility.xlSheetVeryHidden
        xlApp.ScreenUpdating = True
        Call ReleaseExcelObj()
    End Sub

    Private Sub Button_hienSheet_Click(sender As Object, e As EventArgs) Handles Button_hienSheet.Click
        On Error Resume Next
        Dim sh As Worksheet
        Call StartExcel()
        sh.Visible = XlSheetVisibility.xlSheetVisible
        sh = xlWB.Sheets("addbeamlv3")
        sh.Visible = XlSheetVisibility.xlSheetVisible
        sh = xlWB.Sheets("dulieu_insert")
        sh.Visible = XlSheetVisibility.xlSheetVisible
        sh = xlWB.Sheets("list")
        sh.Visible = XlSheetVisibility.xlSheetVisible
        Call ReleaseExcelObj()
    End Sub
End Class
