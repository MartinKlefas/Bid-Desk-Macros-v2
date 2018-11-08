Imports System.Data.Common
Imports System.Diagnostics
Imports MySql.Data.MySqlClient
Imports String_Extensions.StringExtensions

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

        Dim j As Integer


        While reader.Read
            If SelectData <> "" Then SelectData.Append(vbCrLf)
            For j = 0 To reader.FieldCount - 1
                If j > 0 Then
                    SelectData.Append(", ")
                End If
                SelectData.Append(reader.GetString(j))

            Next

        End While

        reader.Close()
    End Function

    Public Function SelectData_List(Optional what As String = "*", Optional where As String = "",
                               Optional table As String = ThisAddIn.defaultTable) As List(Of List(Of String))
        Dim cmd As New MySqlCommand
        cmd.Connection = conn
        SelectData_List = New List(Of List(Of String))

        cmd.CommandText = "Select " & what & " from " & table
        If where <> "" Then
            cmd.CommandText.Append(" where " & where)
        End If

        Dim reader As MySqlDataReader
        reader = cmd.ExecuteReader

        Dim j As Integer
        Dim tmp As List(Of String)

        While reader.Read
            tmp = New List(Of String)

            For j = 0 To reader.FieldCount - 1
                tmp.Add(reader.GetString(j))
            Next
            SelectData_List.Add(tmp)

        End While

    End Function
End Class
