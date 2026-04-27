Imports System.Data.OleDb

Public Class Expenses
    Private Sub Expenses_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim query As String = "SELECT * FROM Expenses"
            Dim dt As DataTable = DataAccessHelper.GetData(query)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MsgBox("خطأ في تحميل بيانات المصروفات: " & ex.Message)
        End Try
    End Sub

    ' زر حفظ
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        If String.IsNullOrWhiteSpace(TextBox5.Text) Then
            MsgBox("يرجى إدخال مبلغ الصرف")
            Return
        End If

        Try
            Dim query As String = "INSERT INTO Expenses (Expense_type, Amount, Expense_Date, notes) VALUES (?, ?, ?, ?)"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", ComboBox4.Text))
            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(TextBox5.Text)))
            parameters.Add(New OleDbParameter("?", DateTimePicker2.Value.Date))
            parameters.Add(New OleDbParameter("?", TextBox8.Text))

            DataAccessHelper.ExecuteNonQuery(query, parameters)
            MsgBox("تم حفظ المصروف بنجاح")
            LoadData()
            ClearFields()
        Catch ex As Exception
            MsgBox("خطأ في الحفظ: " & ex.Message)
        End Try
    End Sub

    ' زر تعديل
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        If DataGridView2.SelectedRows.Count = 0 Then
            MsgBox("يرجى اختيار سجل من الجدول للتعديل")
            Return
        End If

        Try
            Dim id As Integer = DataGridView2.SelectedRows(0).Cells("Expense_ID").Value
            Dim query As String = "UPDATE Expenses SET Expense_type = ?, Amount = ?, Expense_Date = ?, notes = ? WHERE Expense_ID = ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", ComboBox4.Text))
            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(TextBox5.Text)))
            parameters.Add(New OleDbParameter("?", DateTimePicker2.Value.Date))
            parameters.Add(New OleDbParameter("?", TextBox8.Text))
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
            Dim id As Integer = DataGridView2.SelectedRows(0).Cells("Expense_ID").Value
            Dim query As String = "DELETE FROM Expenses WHERE Expense_ID = ?"
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
            Dim query As String = "SELECT * FROM Expenses WHERE Expense_type LIKE ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", "%" & ComboBox4.Text & "%"))

            Dim dt As DataTable = DataAccessHelper.GetData(query, parameters)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MsgBox("خطأ في البحث: " & ex.Message)
        End Try
    End Sub

    ' زر عرض (تحديث)
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        LoadData()
    End Sub

    ' زر إلغاء (تفريغ الحقول)
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ClearFields()
    End Sub

    ' زر إضافة (اختياري لتفريغ الحقول)
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ClearFields()
    End Sub

    Private Sub ClearFields()
        TextBox1.Clear()
        TextBox5.Clear()
        TextBox8.Clear()
        ComboBox4.SelectedIndex = -1
    End Sub

    Private Sub DataGridView2_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView2.SelectionChanged
        If DataGridView2.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView2.SelectedRows(0)
            TextBox1.Text = If(IsDBNull(row.Cells("Expense_ID").Value), "", row.Cells("Expense_ID").Value.ToString())
            TextBox5.Text = If(IsDBNull(row.Cells("Amount").Value), "0", row.Cells("Amount").Value.ToString())
            TextBox8.Text = If(IsDBNull(row.Cells("notes").Value), "", row.Cells("notes").Value.ToString())
            ComboBox4.Text = If(IsDBNull(row.Cells("Expense_type").Value), "", row.Cells("Expense_type").Value.ToString())
            DateTimePicker2.Value = If(IsDBNull(row.Cells("Expense_Date").Value), DateTime.Now, row.Cells("Expense_Date").Value)
        End If
    End Sub

End Class
