Imports System.Data.OleDb

Public Class payment
    ' زر حفظ
    Private Sub حفظToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles حفظToolStripMenuItem.Click
        If String.IsNullOrWhiteSpace(TextBox2.Text) Or String.IsNullOrWhiteSpace(TextBox3.Text) Then
            MsgBox("يرجى إدخال اسم المستلم والمبلغ")
            Return
        End If

        Try
            Dim query As String = "INSERT INTO payment (payer_Name, paid_Amount, PaymentMethod, PaymentDate, Notes) VALUES (?, ?, ?, ?, ?)"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", TextBox2.Text)) ' اسم الدافعة
            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(TextBox3.Text))) ' المبلغ
            parameters.Add(New OleDbParameter("?", ComboBox4.Text)) ' مقابل ماذا
            parameters.Add(New OleDbParameter("?", DateTimePicker2.Value.Date))
            parameters.Add(New OleDbParameter("?", "سند قبض"))

            DataAccessHelper.ExecuteNonQuery(query, parameters)
            MsgBox("تم حفظ السند بنجاح")
        Catch ex As Exception
            MsgBox("خطأ في الحفظ: " & ex.Message)
        End Try
    End Sub

    ' زر حذف
    Private Sub حذفToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles حذفToolStripMenuItem.Click
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            MsgBox("يرجى إدخال رقم السند للحذف")
            Return
        End If

        If MsgBox("هل أنت متأكد من حذف هذا السند؟", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Return

        Try
            Dim query As String = "DELETE FROM payment WHERE payment_ID = ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", Convert.ToInt32(TextBox1.Text)))

            DataAccessHelper.ExecuteNonQuery(query, parameters)
            MsgBox("تم الحذف بنجاح")
            ClearFields()
        Catch ex As Exception
            MsgBox("خطأ في الحذف: " & ex.Message)
        End Try
    End Sub

    ' زر إضافة (جديد)
    Private Sub إضافةToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles إضافةToolStripMenuItem.Click
        ClearFields()
    End Sub

    ' زر خروج
    Private Sub خروجToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles خروجToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ClearFields()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        ComboBox4.SelectedIndex = -1
    End Sub

End Class
