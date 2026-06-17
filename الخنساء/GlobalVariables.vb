Imports System.Data.OleDb
Imports System.Data

Module GlobalVariables

    Public AccessDBPath As String = Application.StartupPath & "\users.accdb"

    Public ConnectionString As String =
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDBPath & ";Persist Security Info=False;"

    Public Conn As New OleDbConnection(ConnectionString)

End Module
