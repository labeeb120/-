Public Class MINE
    ' فتح واجهة الطلاب من القائمة
    Private Sub واجههالطلابToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles واجههالطلابToolStripMenuItem.Click
        studnt.Show()
    End Sub

    ' فتح واجهة الطلاب من الزر الجانبي
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        studnt.Show()
    End Sub

    ' فتح واجهة العاملين (الموظفين)
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        emp.Show()
    End Sub

    ' فتح واجهة المصروفات
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Expenses.Show()
    End Sub

    ' فتح واجهة الدفع المالي (mane)
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        mane.Show()
    End Sub

    ' فتح واجهة الدعم (الداعمين)
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Support.Show()
    End Sub

    ' خروج من البرنامج
    Private Sub خروجToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles خروجToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub قبضToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles قبضToolStripMenuItem.Click
        payment.Show()
    End Sub

    Private Sub صرفToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles صرفToolStripMenuItem.Click
        Receipt.Show()
    End Sub

End Class
