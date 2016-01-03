Public Class Menu
    Sub ExitBOQtools()
        If MsgBox("Bạn có muốn thoát khỏi chương trình", vbOKCancel, "BOQ-Tools(C)DHK") = vbOK Then
            Me.Close()
        End If
    End Sub
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Call ExitBOQtools()
    End Sub
    Private Sub ReinforcedConcreteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReinforcedConcreteToolStripMenuItem.Click
        uf_Menu_ReinConc.Show()
    End Sub

    Private Sub AboutAuthorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutAuthorToolStripMenuItem.Click

    End Sub
End Class