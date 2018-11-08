Imports System.Diagnostics
Imports MySql.Data.MySqlClient
Public Class clsDatabase
    Dim conn As MySqlConnection
    Public Sub Make_connection(server As String, user As String,
                                    database As String, port As Integer)
        conn = New MySqlConnection
        conn.ConnectionString = "server=" & server & ";user=" & user &
                                 ";database=" & database & ";port=" & port '&
        '";password=" & password

        conn.Open()

    End Sub

    Sub New(server As String, user As String,
                                    database As String, port As Integer)

        Call Make_connection(server, user, database, port)

    End Sub

    Public Function SelectData(Optional what As String = "*", Optional where As String = "",
                               Optional table As String = ThisAddIn.defaultTable) As String
        Dim cmd As New MySqlCommand
        cmd.Connection = conn

        cmd.CommandText = "Select " & what & " from " & table
        If where <> "" Then
            cmd.CommandText = cmd.CommandText & " where " & where
        End If

        SelectData = ""

        Dim reader As MySqlDataReader
        reader = cmd.ExecuteReader

        While reader.Read
            Debug.WriteLine(reader)
        End While
    End Function

End Class
