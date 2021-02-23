Public Class Frm

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        生徒データマスタ.ShowDialog()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        科目データマスタ.ShowDialog()
        科目データマスタ.Dispose()
    End Sub
End Class