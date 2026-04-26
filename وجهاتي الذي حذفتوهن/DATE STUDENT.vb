Imports System.Data.OleDb
Imports System.Data

Public Class DATE_STUDENT

    Dim SelectedID As Integer ' ✔️ مرة واحدة فقط

    Private Sub DATE_STUDENT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
    End Sub

    Private Sub LoadData()
        Dim query As String = "SELECT * FROM Students"
        DataGridView2.DataSource = DataAccessHelper.GetData(query)
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then
            Dim row = DataGridView2.Rows(e.RowIndex)

            SelectedID = row.Cells("StudentID").Value

            TextBox2.Text = row.Cells("StudentName").Value.ToString()
            TextBox6.Text = row.Cells("Phone").Value.ToString()
        End If
    End Sub

End Class