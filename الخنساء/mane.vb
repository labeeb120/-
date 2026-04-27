Imports System.Data.OleDb

Public Class mane
    Private Sub mane_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub

    ' جلب البيانات وعرضها في الجدول
    Private Sub LoadData()
        Try
            Dim query As String = "SELECT * FROM Revenue"
            Dim dt As DataTable = DataAccessHelper.GetData(query)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MsgBox("خطأ في تحميل البيانات: " & ex.Message)
        End Try
    End Sub

    ' زر حفظ
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If String.IsNullOrWhiteSpace(TextBox1.Text) Or String.IsNullOrWhiteSpace(TextBox4.Text) Then
            MsgBox("يرجى ملء الحقول الأساسية (الاسم والمبلغ)")
            Return
        End If

        Try
            Dim query As String = "INSERT INTO Revenue ([Revenue Name], [Revenue phone], Amount, RevenueDate) VALUES (?, ?, ?, ?)"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", TextBox1.Text)) ' اسم الدافع

            Dim phoneVal As Object = DBNull.Value
            If IsNumeric(TextBox3.Text) Then
                Try
                    phoneVal = Convert.ToInt64(TextBox3.Text)
                Catch ex As OverflowException
                    phoneVal = 0
                End Try
            End If
            parameters.Add(New OleDbParameter("?", phoneVal))

            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(TextBox4.Text))) ' المبلغ
            parameters.Add(New OleDbParameter("?", DateTimePicker1.Value.Date)) ' تاريخ الدفع

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
            Dim id As Integer = DataGridView2.SelectedRows(0).Cells("RevenueID").Value
            Dim query As String = "UPDATE Revenue SET [Revenue Name] = ?, [Revenue phone] = ?, Amount = ?, RevenueDate = ? WHERE RevenueID = ?"
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

            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(TextBox4.Text)))
            parameters.Add(New OleDbParameter("?", DateTimePicker1.Value.Date))
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
            Dim id As Integer = DataGridView2.SelectedRows(0).Cells("RevenueID").Value
            Dim query As String = "DELETE FROM Revenue WHERE RevenueID = ?"
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
            Dim query As String = "SELECT * FROM Revenue WHERE [Revenue Name] LIKE ?"
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

    ' زر إلغاء (تفريغ الحقول)
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ClearFields()
    End Sub

    Private Sub ClearFields()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox8.Clear()
    End Sub

    ' عند اختيار سجل من الجدول
    Private Sub DataGridView2_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView2.SelectionChanged
        If DataGridView2.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView2.SelectedRows(0)
            TextBox1.Text = If(IsDBNull(row.Cells("Revenue Name").Value), "", row.Cells("Revenue Name").Value.ToString())
            TextBox3.Text = If(IsDBNull(row.Cells("Revenue phone").Value), "", row.Cells("Revenue phone").Value.ToString())
            TextBox4.Text = If(IsDBNull(row.Cells("Amount").Value), "", row.Cells("Amount").Value.ToString())
            DateTimePicker1.Value = If(IsDBNull(row.Cells("RevenueDate").Value), DateTime.Now, row.Cells("RevenueDate").Value)
        End If
    End Sub
End Class
