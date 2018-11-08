Imports System.Diagnostics
Imports mysql.data.mysqlclient

Public Class ThisAddIn
    Public Const server As String = "172.27.41.59"
    Public Const user As String = "root"
    Public Const database As String = "Bids"
    Public Const defaultTable As String = "All_Bids"
    Public Const port As Integer = 3306

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
