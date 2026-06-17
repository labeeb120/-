Imports System.Data.OleDb

Public Class Support
    Private Sub Support_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim query As String = "SELECT * FROM Support"
            Dim dt As DataTable = DataAccessHelper.GetData(query)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MsgBox("خطأ في تحميل بيانات الداعمين: " & ex.Message)
        End Try
    End Sub

    ' زر حفظ
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            MsgBox("يرجى إدخال اسم الداعم")
            Return
        End If

        Try
            Dim query As String = "INSERT INTO Support (SupporterName, phone, email, adress, Amount, SupportDate, PaymentMethod, [Beneficiary Entity]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", TextBox1.Text))

            ' معالجة رقم الهاتف
            Dim phoneVal As Object = DBNull.Value
            If IsNumeric(TextBox3.Text) Then
                Try
                    phoneVal = Convert.ToInt64(TextBox3.Text)
                Catch ex As OverflowException
                    phoneVal = 0
                End Try
            End If
            parameters.Add(New OleDbParameter("?", phoneVal))

            parameters.Add(New OleDbParameter("?", TextBox5.Text))
            parameters.Add(New OleDbParameter("?", TextBox7.Text))
            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(If(String.IsNullOrEmpty(TextBox4.Text), "0", TextBox4.Text))))
            parameters.Add(New OleDbParameter("?", DateTimePicker1.Value.Date))
            parameters.Add(New OleDbParameter("?", ComboBox1.Text))
            parameters.Add(New OleDbParameter("?", ComboBox2.Text))

            DataAccessHelper.ExecuteNonQuery(query, parameters)
            MsgBox("تم الحفظ بنجاح")
            LoadData()
            ClearFields()
        Catch ex As Exception
            MsgBox("خطأ في الحفظ: " & ex.Message)
        End Try
    End Sub

    ' زر تعديل
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView2.SelectedRows.Count = 0 Then
            MsgBox("يرجى اختيار سجل من الجدول للتعديل")
            Return
        End If

        Try
            Dim id As Integer = DataGridView2.SelectedRows(0).Cells("SupporterID").Value
            Dim query As String = "UPDATE Support SET SupporterName = ?, phone = ?, email = ?, adress = ?, Amount = ?, SupportDate = ?, PaymentMethod = ?, [Beneficiary Entity] = ? WHERE SupporterID = ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", TextBox1.Text))

            Dim phoneVal As Object = DBNull.Value
            If IsNumeric(TextBox3.Text) Then
                Try
                    phoneVal = Convert.ToInt64(TextBox3.Text)
                Catch ex As OverflowException
                    phoneVal = 0
                End Try
            End If
            parameters.Add(New OleDbParameter("?", phoneVal))

            parameters.Add(New OleDbParameter("?", TextBox5.Text))
            parameters.Add(New OleDbParameter("?", TextBox7.Text))
            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(If(String.IsNullOrEmpty(TextBox4.Text), "0", TextBox4.Text))))
            parameters.Add(New OleDbParameter("?", DateTimePicker1.Value.Date))
            parameters.Add(New OleDbParameter("?", ComboBox1.Text))
            parameters.Add(New OleDbParameter("?", ComboBox2.Text))
            parameters.Add(New OleDbParameter("?", id))

            DataAccessHelper.ExecuteNonQuery(query, parameters)
            MsgBox("تم التعديل بنجاح")
            LoadData()
        Catch ex As Exception
            MsgBox("خطأ في التعديل: " & ex.Message)
        End Try
    End Sub

    ' زر حذف
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView2.SelectedRows.Count = 0 Then
            MsgBox("يرجى اختيار سجل للحذف")
            Return
        End If

        If MsgBox("هل أنت متأكد من الحذف؟", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Return

        Try
            Dim id As Integer = DataGridView2.SelectedRows(0).Cells("SupporterID").Value
            Dim query As String = "DELETE FROM Support WHERE SupporterID = ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", id))

            DataAccessHelper.ExecuteNonQuery(query, parameters)
            MsgBox("تم الحذف بنجاح")
            LoadData()
        Catch ex As Exception
            MsgBox("خطأ في الحذف: " & ex.Message)
        End Try
    End Sub

    ' زر بحث
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Dim query As String = "SELECT * FROM Support WHERE SupporterName LIKE ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", "%" & TextBox1.Text & "%"))

            Dim dt As DataTable = DataAccessHelper.GetData(query, parameters)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MsgBox("خطأ في البحث: " & ex.Message)
        End Try
    End Sub

    ' زر عرض (تحديث)
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        LoadData()
    End Sub

    ' زر إلغاء
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ClearFields()
    End Sub

    Private Sub ClearFields()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
    End Sub

    Private Sub DataGridView2_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView2.SelectionChanged
        If DataGridView2.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView2.SelectedRows(0)
            TextBox1.Text = If(IsDBNull(row.Cells("SupporterName").Value), "", row.Cells("SupporterName").Value.ToString())
            TextBox2.Text = If(IsDBNull(row.Cells("SupporterID").Value), "", row.Cells("SupporterID").Value.ToString())
            TextBox3.Text = If(IsDBNull(row.Cells("phone").Value), "", row.Cells("phone").Value.ToString())
            TextBox5.Text = If(IsDBNull(row.Cells("email").Value), "", row.Cells("email").Value.ToString())
            TextBox7.Text = If(IsDBNull(row.Cells("adress").Value), "", row.Cells("adress").Value.ToString())
            TextBox4.Text = If(IsDBNull(row.Cells("Amount").Value), "0", row.Cells("Amount").Value.ToString())
            DateTimePicker1.Value = If(IsDBNull(row.Cells("SupportDate").Value), DateTime.Now, row.Cells("SupportDate").Value)
            ComboBox1.Text = If(IsDBNull(row.Cells("PaymentMethod").Value), "", row.Cells("PaymentMethod").Value.ToString())
            ComboBox2.Text = If(IsDBNull(row.Cells("Beneficiary Entity").Value), "", row.Cells("Beneficiary Entity").Value.ToString())
        End If
    End Sub
End Class
