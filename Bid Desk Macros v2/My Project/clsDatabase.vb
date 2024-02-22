Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports String_Extensions.StringExtensions

Public Class ClsDatabase

    Private ReadOnly connectionString As String

    Sub New(server As String, database As String)

        connectionString = "Data Source=" & server & ";Initial Catalog=" & database & ";Integrated Security=SSPI"
    End Sub

    Public Function ValueExists(value As String, Optional column As String = "DealID", Optional table As String = ThisAddIn.defaultTable) As Boolean
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Dim result As Boolean
            Using cmd As New SqlCommand With {
            .Connection = conn,
            .CommandText = "SELECT TOP 1 " & column & " FROM " & table & " WHERE " & column & " = '" & value & "'"
        }

                Try
                    Using reader As SqlDataReader = cmd.ExecuteReader
                        result = reader.HasRows
                    End Using
                Catch
                    'some kind of overlap in sql reader usage!
                    Debug.WriteLine("SQL Error in ValueExists")
                    Return False
                End Try
            End Using

            Return result
        End Using



    End Function

    Public Function SelectData(Optional what As String = "*", Optional where As String = "",
                               Optional table As String = ThisAddIn.defaultTable) As String
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand With {
            .Connection = conn,
            .CommandText = "Select " & what & " from " & table
        }

                If where <> "" Then
                    cmd.CommandText = cmd.CommandText & " where " & where
                End If

                SelectData = ""
                Try
                    Using reader As SqlDataReader = cmd.ExecuteReader

                        Dim j As Integer


                        While reader.Read
                            If SelectData <> "" Then SelectData.Append(vbCrLf)
                            For j = 0 To reader.FieldCount - 1
                                If j > 0 Then
                                    SelectData.Append(", ")
                                End If

                                Dim value As String
                                Select Case reader.GetDataTypeName(j)
                                    Case "varchar", "text"
                                        value = reader.GetString(j)
                                    Case "int"
                                        value = CStr(reader.GetSqlInt32(j))
                                    Case "bit"
                                        value = CStr(reader.GetSqlBoolean(j))
                                    Case "datetime"
                                        value = reader.GetDateTime(j).ToString
                                    Case Else
                                        Debug.WriteLine(reader.GetDataTypeName(j))
                                        value = ""
                                End Select

                                SelectData.Append(value)

                            Next

                        End While
                    End Using
                Catch
                    Debug.WriteLine("SQL Error in SelectData")

                End Try
            End Using
        End Using

    End Function

    Friend Function FindColumns() As List(Of String)
        FindColumns = New List(Of String)
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand With {
            .Connection = conn,
            .CommandText = "exec sp_columns all_bids"
        }
                Try
                    Using reader As SqlDataReader = cmd.ExecuteReader

                        While reader.Read
                            For j = 0 To reader.FieldCount - 1
                                If reader.GetName(j) = "COLUMN_NAME" Then
                                    FindColumns.Add(reader.GetString(j))
                                    Exit For
                                End If
                            Next
                        End While
                    End Using
                Catch
                    Debug.WriteLine("SQL Error in FindColumns")
                End Try
            End Using
        End Using

    End Function

    Public Function SelectData_Date(what As String, where As String,
                               Optional table As String = ThisAddIn.defaultTable) As Date
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand With {
            .Connection = conn,
            .CommandText = "Select " & what & " from " & table
        }

                If where <> "" Then
                    cmd.CommandText = cmd.CommandText & " where " & where
                End If

                SelectData_Date = Nothing

                Try
                    Using reader As SqlDataReader = cmd.ExecuteReader



                        While reader.Read
                            If reader.FieldCount > 1 Then
                                Exit While
                            End If

                            SelectData_Date = reader.GetDateTime(0)

                        End While

                    End Using
                Catch
                    Debug.WriteLine("SQL Error in SelectData_Date")
                End Try
            End Using
        End Using
    End Function

    Public Function SelectData_Dict(Optional what As String = "*", Optional where As String = "",
                               Optional table As String = ThisAddIn.defaultTable) As List(Of Dictionary(Of String, String))
        Using conn As New SqlConnection(connectionString)
            conn.Open()

            Using cmd As New SqlCommand With {
            .Connection = conn,
            .CommandText = "Select " & what & " from " & table
        }


                SelectData_Dict = New List(Of Dictionary(Of String, String))


                If where <> "" Then
                    cmd.CommandText.Append(" where " & where)
                End If
                Try
                    Using reader As SqlDataReader = cmd.ExecuteReader

                        Dim j As Integer
                        Dim tmp As Dictionary(Of String, String)
                        Dim key As String, value As String

                        While reader.Read
                            tmp = New Dictionary(Of String, String)

                            For j = 0 To reader.FieldCount - 1
                                key = reader.GetName(j)
                                Try
                                    Select Case reader.GetDataTypeName(j)
                                        Case "varchar", "text"
                                            value = reader.GetString(j)
                                        Case "int"
                                            value = CStr(reader.GetSqlInt32(j))
                                        Case "bit"
                                            value = CStr(reader.GetSqlBoolean(j))
                                        Case "datetime"
                                            value = reader.GetDateTime(j).ToString
                                        Case Else
                                            Debug.WriteLine(reader.GetDataTypeName(j))
                                            value = ""
                                    End Select
                                Catch
                                    Debug.WriteLine(reader.GetDataTypeName(j))
                                    value = ""
                                End Try

                                tmp.Add(key, value)


                            Next
                            SelectData_Dict.Add(tmp)

                        End While

                    End Using
                Catch
                    Debug.WriteLine("SQL Error in SelectData_Dict")
                End Try
            End Using
        End Using

    End Function

    Public Function Add_Data(what As Dictionary(Of String, String), Optional table As String = ThisAddIn.defaultTable) As Integer
        Using conn As New SqlConnection(connectionString)
            conn.Open()

            Using cmd As New SqlCommand
                cmd.Connection = conn


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

                Try
                    Add_Data = cmd.ExecuteNonQuery

                Catch
                    Add_Data = 0
                    Debug.WriteLine("Error in adding lines to the database")

                    Debug.WriteLine(cmd.CommandText)
                End Try
            End Using
        End Using
    End Function

    Public Function Update_Data(what As String, Optional where As String = "",
                               Optional table As String = ThisAddIn.defaultTable) As Integer
        Using conn As New SqlConnection(connectionString)
            conn.Open()
            Using cmd As New SqlCommand
                Try

                    cmd.Connection = conn


                    cmd.CommandText = "UPDATE " & table & " SET " & what

                    If where <> "" Then
                        cmd.CommandText &= " WHERE " & where
                    End If

                    Update_Data = cmd.ExecuteNonQuery

                Catch ex As Exception
                    Update_Data = 0
                    Debug.WriteLine("SQL Error updating data:" & vbCrLf & ex.Message)
                    Debug.WriteLine(cmd.CommandText)
                End Try
            End Using
        End Using

    End Function


    Function MS_SQL_Escape(rawStr As String) As String
        MS_SQL_Escape = Replace(rawStr, "'", "''")
    End Function
End Class
