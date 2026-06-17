Imports System.Data.OleDb
Imports System.Data

Public Class DataAccessHelper

    Public Shared Function ExecuteNonQuery(query As String, Optional parameters As List(Of OleDbParameter) = Nothing) As Integer
        Using conn As New OleDbConnection(GlobalVariables.ConnectionString)
            Using cmd As New OleDbCommand(query, conn)
                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters.ToArray())
                End If

                conn.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Public Shared Function GetData(query As String, Optional parameters As List(Of OleDbParameter) = Nothing) As DataTable
        Dim dt As New DataTable()

        Using conn As New OleDbConnection(GlobalVariables.ConnectionString)
            Using cmd As New OleDbCommand(query, conn)

                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters.ToArray())
                End If

                Using da As New OleDbDataAdapter(cmd)
                    da.Fill(dt)
                End Using

            End Using
        End Using

        Return dt
    End Function

    Public Shared Function ExecuteScalar(query As String, Optional parameters As List(Of OleDbParameter) = Nothing) As Object
        Using conn As New OleDbConnection(GlobalVariables.ConnectionString)
            Using cmd As New OleDbCommand(query, conn)

                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters.ToArray())
                End If

                conn.Open()
                Return cmd.ExecuteScalar()

            End Using
        End Using
    End Function

End Class
