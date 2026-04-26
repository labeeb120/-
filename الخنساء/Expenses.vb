Public Class Expenses
    Private Sub Expenses_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Conn.Open()
            MsgBox("تم الاتصال بقاعدة البيانات بنجاح ✅")
            Conn.Close()
        Catch ex As Exception
            MsgBox("فشل الاتصال ❌" & vbCrLf & ex.Message)
        End Try
    End Sub
End Class