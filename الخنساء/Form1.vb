
Imports System.Data.OleDb

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    ' زر الدخول
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If String.IsNullOrWhiteSpace(TextBox1.Text) Or String.IsNullOrWhiteSpace(TextBox2.Text) Then
            MsgBox("يرجى إدخال اسم المستخدم وكلمة المرور", MsgBoxStyle.Exclamation, "تنبيه")
            Return
        End If

        Try
            Dim query As String = "SELECT COUNT(*) FROM Users WHERE Username = ? AND [Password] = ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", TextBox1.Text))
            parameters.Add(New OleDbParameter("?", TextBox2.Text))

            Dim result As Integer = Convert.ToInt32(DataAccessHelper.ExecuteScalar(query, parameters))

            If result > 0 Then
                MsgBox("تم تسجيل الدخول بنجاح", MsgBoxStyle.Information, "مرحباً")
                MINE.Show()
                Me.Hide()
            Else
                MsgBox("اسم المستخدم أو كلمة المرور غير صحيحة", MsgBoxStyle.Critical, "فشل الدخول")
            End If
        Catch ex As Exception
            MsgBox("خطأ أثناء محاولة تسجيل الدخول: " & ex.Message, MsgBoxStyle.Critical, "خطأ")
        End Try
    End Sub

    ' زر الخروج
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Application.Exit()
    End Sub

    ' زر نسيت كلمة السر
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("يرجى التواصل مع مسؤول النظام لاستعادة كلمة المرور", MsgBoxStyle.Information, "نسيت كلمة السر")
    End Sub

    ' زر تغيير كلمة السر
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' يمكن إضافة واجهة لتغيير كلمة السر هنا لاحقاً
        MsgBox("هذه الخاصية ستتوفر قريباً", MsgBoxStyle.Information, "تغيير كلمة السر")
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs)

    End Sub
End Class
