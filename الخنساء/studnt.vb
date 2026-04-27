Imports System.Data.OleDb

﻿Public Class studnt
    Private Sub studnt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub

    Private Sub LoadData()
        Try
            Dim query As String = "SELECT * FROM Students"
            Dim dt As DataTable = DataAccessHelper.GetData(query)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MsgBox("خطأ في تحميل بيانات الطلاب: " & ex.Message)
        End Try
    End Sub

    ' زر حفظ
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            MsgBox("يرجى إدخال اسم الطالبة")
            Return
        End If

        Try
            Dim query As String = "INSERT INTO Students (StudentName, Phone, Notes, Student_Status, EnrollmentDate) VALUES (?, ?, ?, ?, ?)"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", TextBox1.Text))
            parameters.Add(New OleDbParameter("?", TextBox3.Text))
            parameters.Add(New OleDbParameter("?", TextBox8.Text))
            parameters.Add(New OleDbParameter("?", ComboBox4.Text))
            parameters.Add(New OleDbParameter("?", DateTimePicker2.Value.Date))

            DataAccessHelper.ExecuteNonQuery(query, parameters)
            MsgBox("تم حفظ بيانات الطالبة بنجاح")
            LoadData()
            ClearFields()
        Catch ex As Exception
            MsgBox("خطأ في الحفظ: " & ex.Message)
        End Try
    End Sub

    ' زر تعديل
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If DataGridView2.SelectedRows.Count = 0 Then
            MsgBox("يرجى اختيار طالبة من الجدول للتعديل")
            Return
        End If

        Try
            Dim id As Integer = DataGridView2.SelectedRows(0).Cells("StudentID").Value
            Dim query As String = "UPDATE Students SET StudentName = ?, Phone = ?, Notes = ?, Student_Status = ?, EnrollmentDate = ? WHERE StudentID = ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", TextBox1.Text))
            parameters.Add(New OleDbParameter("?", TextBox3.Text))
            parameters.Add(New OleDbParameter("?", TextBox8.Text))
            parameters.Add(New OleDbParameter("?", ComboBox4.Text))
            parameters.Add(New OleDbParameter("?", DateTimePicker2.Value.Date))
            parameters.Add(New OleDbParameter("?", id))

            DataAccessHelper.ExecuteNonQuery(query, parameters)
            MsgBox("تم تعديل بيانات الطالبة بنجاح")
            LoadData()
        Catch ex As Exception
            MsgBox("خطأ في التعديل: " & ex.Message)
        End Try
    End Sub

    ' زر حذف
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If DataGridView2.SelectedRows.Count = 0 Then
            MsgBox("يرجى اختيار طالبة للحذف")
            Return
        End If

        If MsgBox("هل أنت متأكد من حذف بيانات هذه الطالبة؟", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Return

        Try
            Dim id As Integer = DataGridView2.SelectedRows(0).Cells("StudentID").Value
            Dim query As String = "DELETE FROM Students WHERE StudentID = ?"
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
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Try
            Dim query As String = "SELECT * FROM Students WHERE StudentName LIKE ?"
            Dim parameters As New List(Of OleDbParameter)
            parameters.Add(New OleDbParameter("?", "%" & TextBox1.Text & "%"))

            Dim dt As DataTable = DataAccessHelper.GetData(query, parameters)
            DataGridView2.DataSource = dt
        Catch ex As Exception
            MsgBox("خطأ في البحث: " & ex.Message)
        End Try
    End Sub

    ' زر عرض (تحديث)
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        LoadData()
    End Sub

    ' زر خروج
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Me.Close()
    End Sub

    ' زر إضافة (تفريغ الحقول)
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ClearFields()
    End Sub

    Private Sub ClearFields()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox8.Clear()
        TextBox10.Clear()
        TextBox7.Clear()
        TextBox6.Clear()
    End Sub

    Private Sub DataGridView2_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView2.SelectionChanged
        If DataGridView2.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView2.SelectedRows(0)
            TextBox1.Text = row.Cells("StudentName").Value.ToString()
            TextBox3.Text = row.Cells("Phone").Value.ToString()
            TextBox8.Text = row.Cells("Notes").Value.ToString()
            ComboBox4.Text = row.Cells("Student_Status").Value.ToString()
            DateTimePicker2.Value = If(IsDBNull(row.Cells("EnrollmentDate").Value), DateTime.Now, row.Cells("EnrollmentDate").Value)
            TextBox2.Text = row.Cells("StudentID").Value.ToString()
        End If
    End Sub
End Class