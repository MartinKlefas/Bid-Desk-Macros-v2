Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports String_Extensions.StringExtensions

Public Class ClsDatabase
    Dim conn As SqlConnection
    Public Sub Make_connection(server As String, user As String,
                                    database As String, password As String)
        Try
            Dim builder As New SqlConnectionStringBuilder
            With builder

                .DataSource = server
                .UserID = user
                .Password = password

                .InitialCatalog = database

            End With

            conn = New SqlConnection(builder.ConnectionString)

            conn.Open()
        Catch

            Dim connectionString As String


            connectionString = "Data Source=GBMNCDT12830\SQLEXPRESS;Initial Catalog=bids;Integrated Security=SSPI"
            conn = New SqlConnection(connectionString)
            conn.Open()
        End Try

    End Sub

    Sub New(server As String, user As String,
                                    database As String, password As String)

        Call Make_connection(server, user, database, password)

    End Sub

    Public Function ValueExists(value As String, Optional column As String = "DealID", Optional table As String = ThisAddIn.defaultTable) As Boolean


        Dim cmd As New SqlCommand With {
            .Connection = conn,
            .CommandText = "SELECT TOP 1 " & column & " FROM " & table & " WHERE" & column & " = '" & value & "'"
        }

        Dim reader As SqlDataReader
        reader = cmd.ExecuteReader
        Return reader.HasRows

    End Function

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

    Public Function SelectData_Date(what As String, where As String,
                               Optional table As String = ThisAddIn.defaultTable) As Date
        Dim cmd As New SqlCommand With {
            .Connection = conn,
            .CommandText = "Select " & what & " from " & table
        }

        If where <> "" Then
            cmd.CommandText = cmd.CommandText & " where " & where
        End If

        SelectData_Date = Nothing

        Dim reader As SqlDataReader
        reader = cmd.ExecuteReader


        While reader.Read
            If reader.FieldCount > 1 Then
                Exit While
            End If

            SelectData_Date = reader.GetDateTime(0)

        End While

        reader.Close()
    End Function

    Public Function SelectData_Dict(Optional what As String = "*", Optional where As String = "",
                               Optional table As String = ThisAddIn.defaultTable) As List(Of Dictionary(Of String, String))
        Dim cmd As New SqlCommand With {
            .Connection = conn,
            .CommandText = "Select " & what & " from " & table
        }


        SelectData_Dict = New List(Of Dictionary(Of String, String))


        If where <> "" Then
            cmd.CommandText.Append(" where " & where)
        End If

        Dim reader As SqlDataReader
        reader = cmd.ExecuteReader

        Dim j As Integer
        Dim tmp As Dictionary(Of String, String)

        While reader.Read
            tmp = New Dictionary(Of String, String)

            For j = 0 To reader.FieldCount - 1
                tmp.Add(reader.GetName(j), reader.GetString(j))
            Next
            SelectData_Dict.Add(tmp)

        End While

    End Function

    Public Function Add_Data(what As Dictionary(Of String, String), Optional table As String = ThisAddIn.defaultTable) As Boolean

        Dim cmd As New SqlCommand With {
           .Connection = conn
        }

        Dim columns As String, values As String

        columns = "("
        values = "("
        For Each kvp As KeyValuePair(Of String, String) In what
            columns &= "[" & kvp.Key & "], "
            values &= "N'" & MS_SQL_Escape(kvp.Value) & "', "
        Next

        columns = Left(columns, columns.Length - 2) & ")"
        values = Left(values, values.Length - 2) & ")"

        cmd.CommandText = "INSERT INTO " & table & columns & " VALUES " & values

        Add_Data = (cmd.ExecuteNonQuery = 1)

    End Function

    Public Function Update_Data(what As String, Optional where As String = "",
                               Optional table As String = ThisAddIn.defaultTable) As Integer
        Dim cmd As New SqlCommand With {
           .Connection = conn
        }

        cmd.CommandText = "UPDATE " & table & " SET " & what

        If where <> "" Then
            cmd.CommandText &= " WHERE " & where
        End If

        Update_Data = cmd.ExecuteNonQuery
    End Function


    Function MS_SQL_Escape(rawStr As String) As String
        MS_SQL_Escape = Replace(rawStr, "'", "''")
    End Function
End Class
