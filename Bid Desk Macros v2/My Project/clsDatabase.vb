Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports String_Extensions.StringExtensions

Public Class ClsDatabase
    Dim conn As SqlConnection
    Public Sub Make_connection(server As String, user As String,
                                    database As String, password As String)
        Dim builder As New SqlConnectionStringBuilder
        With builder

            .DataSource = server
            .UserID = user
            .Password = password

            .InitialCatalog = database

        End With

        conn = New SqlConnection(builder.ConnectionString)

        conn.Open()

    End Sub

    Sub New(server As String, user As String,
                                    database As String, password As String)

        Call Make_connection(server, user, database, password)

    End Sub

    Public Function SelectData(Optional what As String = "*", Optional where As String = "",
                               Optional table As String = ThisAddIn.defaultTable) As String
        Dim cmd As New SqlCommand With {
            .Connection = conn,
            .CommandText = "Select " & what & " from " & table
        }

        If where <> "" Then
            cmd.CommandText = cmd.CommandText & " where " & where
        End If

        SelectData = ""

        Dim reader As SqlDataReader
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
        Dim cmd As New SqlCommand With {
            .Connection = conn,
            .CommandText = "Select " & what & " from " & table
        }


        SelectData_List = New List(Of List(Of String))


        If where <> "" Then
            cmd.CommandText.Append(" where " & where)
        End If

        Dim reader As SqlDataReader
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
