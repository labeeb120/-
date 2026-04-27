Imports System.Data.OleDb

﻿Public Class emp
    Private Sub emp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim query As String = "SELECT * FROM Emp"
            Dim dt As DataTable = DataAccessHelper.GetData(query)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MsgBox("خطأ في تحميل بيانات الموظفين: " & ex.Message)
        End Try
    End Sub

    ' زر حفظ
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            MsgBox("يرجى إدخال اسم الموظف")
            Return
        End If

        Try
            Dim query As String = "INSERT INTO Emp (Employee_Name, Gendar, contract_fees, absence_days, Hours_count, contract_Type, Final_Amount) VALUES (?, ?, ?, ?, ?, ?, ?)"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", TextBox1.Text))
            parameters.Add(New OleDbParameter("?", ComboBox3.Text))
            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(If(String.IsNullOrEmpty(TextBox7.Text), "0", TextBox7.Text))))
            parameters.Add(New OleDbParameter("?", Convert.ToInt32(If(String.IsNullOrEmpty(TextBox5.Text), "0", TextBox5.Text))))
            parameters.Add(New OleDbParameter("?", Convert.ToInt32(If(String.IsNullOrEmpty(TextBox10.Text), "0", TextBox10.Text))))
            parameters.Add(New OleDbParameter("?", ComboBox4.Text))
            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(If(String.IsNullOrEmpty(TextBox4.Text), "0", TextBox4.Text))))

            DataAccessHelper.ExecuteNonQuery(query, parameters)
            MsgBox("تم الحفظ بنجاح")
            LoadData()
            ClearFields()
        Catch ex As Exception
            MsgBox("خطأ في الحفظ: " & ex.Message)
        End Try
    End Sub

    ' زر تعديل
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("يرجى اختيار سجل للتعديل")
            Return
        End If

        Try
            Dim id As Integer = DataGridView1.SelectedRows(0).Cells("emp_id").Value
            Dim query As String = "UPDATE Emp SET Employee_Name = ?, Gendar = ?, contract_fees = ?, absence_days = ?, Hours_count = ?, contract_Type = ?, Final_Amount = ? WHERE emp_id = ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", TextBox1.Text))
            parameters.Add(New OleDbParameter("?", ComboBox3.Text))
            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(If(String.IsNullOrEmpty(TextBox7.Text), "0", TextBox7.Text))))
            parameters.Add(New OleDbParameter("?", Convert.ToInt32(If(String.IsNullOrEmpty(TextBox5.Text), "0", TextBox5.Text))))
            parameters.Add(New OleDbParameter("?", Convert.ToInt32(If(String.IsNullOrEmpty(TextBox10.Text), "0", TextBox10.Text))))
            parameters.Add(New OleDbParameter("?", ComboBox4.Text))
            parameters.Add(New OleDbParameter("?", Convert.ToDecimal(If(String.IsNullOrEmpty(TextBox4.Text), "0", TextBox4.Text))))
            parameters.Add(New OleDbParameter("?", id))

            DataAccessHelper.ExecuteNonQuery(query, parameters)
            MsgBox("تم التعديل بنجاح")
            LoadData()
        Catch ex As Exception
            MsgBox("خطأ في التعديل: " & ex.Message)
        End Try
    End Sub

    ' زر حذف
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("يرجى اختيار موظف للحذف")
            Return
        End If

        If MsgBox("هل أنت متأكد من الحذف؟", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Return

        Try
            Dim id As Integer = DataGridView1.SelectedRows(0).Cells("emp_id").Value
            Dim query As String = "DELETE FROM Emp WHERE emp_id = ?"
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
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            Dim query As String = "SELECT * FROM Emp WHERE Employee_Name LIKE ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", "%" & TextBox1.Text & "%"))

            Dim dt As DataTable = DataAccessHelper.GetData(query, parameters)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MsgBox("خطأ في البحث: " & ex.Message)
        End Try
    End Sub

    ' زر عرض (تحديث)
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
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
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            TextBox1.Text = row.Cells("Employee_Name").Value.ToString()
            TextBox2.Text = row.Cells("emp_id").Value.ToString()
            ComboBox3.Text = row.Cells("Gendar").Value.ToString()
            TextBox7.Text = row.Cells("contract_fees").Value.ToString()
            TextBox5.Text = row.Cells("absence_days").Value.ToString()
            TextBox10.Text = row.Cells("Hours_count").Value.ToString()
            ComboBox4.Text = row.Cells("contract_Type").Value.ToString()
            TextBox4.Text = row.Cells("Final_Amount").Value.ToString()
        End If
    End Sub
End Class