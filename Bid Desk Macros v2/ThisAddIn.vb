Imports System.Diagnostics
Imports Microsoft.Office.Interop.Outlook
Imports MySql.Data.MySqlClient

Public Class ThisAddIn
    Public Const server As String = "mklefass-sql2.database.windows.net"
    Public Const user As String = "mklefass"
    Public Const password As String = "nuNDCb4MqmU66j58"
    Public Const database As String = "Bids"
    Public Const defaultTable As String = "all_bids"
    Public Const port As Integer = 1433
    Public Const searchType As StringComparison = vbTextCompare




    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Sub MoveBasedOnDealID(Optional suppressWarnings As Boolean = False)

        Dim obj As Object, success As Boolean
        Dim msg As Outlook.MailItem

        '  Dim olApp As New Outlook.Application 'new throws security error
        Dim DealID As String, targetFolder As String
        Dim olCurrExplorer As Outlook.Explorer
        Dim olCurrSelection As Outlook.Selection


        '  Set olNameSpace = olApp.GetNamespace("MAPI")
        olCurrExplorer = Application.ActiveExplorer
        olCurrSelection = olCurrExplorer.Selection

        For m = 1 To olCurrSelection.Count
            obj = olCurrSelection.Item(m)
            If TypeName(obj) = "MailItem" Then
                msg = obj
                DealID = FindDealID(msg.Subject, msg.Body)
                If DealID = "" Then Exit Sub
                targetFolder = GetFolderbyDeal(DealID, suppressWarnings)

                success = MoveToFolder(targetFolder, msg)
            End If
        Next m
    End Sub
    Sub ReplyToBidRequest()



        Dim obj As Object
        Dim msg As Outlook.MailItem, myGreeting As String, success As Boolean
        Dim msgReply As Outlook.MailItem
        Dim Result As Object, rFName As Object, msgTxt As String

        Dim olCurrExplorer As Outlook.Explorer
        Dim olCurrSelection As Outlook.Selection

        olCurrExplorer = Application.ActiveExplorer
        olCurrSelection = olCurrExplorer.Selection

        If olCurrSelection.Count > 1 Then
            MsgBox("This can only be used with one bid request at a time", vbCritical)
            Exit Sub
        End If

        obj = GetCurrentItem()
        If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
            msg = obj
            msgReply = msg.ReplyAll

            Debug.Write(recordWaitTime(msg.ReceivedTime, Now(), "Me"))
            Result = CreateDealRecord(msgReply)



            rFName = Split(Result(2))

            myGreeting = writeGreeting(Now(), CStr(rFName(0)))

            msgTxt = myGreeting & "<br>&nbsp;I've created the below for you with " & Result(3) & " (ref: " _
                & Result(4) _
                & ").<br>&nbsp;Please check that everything is correct and let me know asap if there are any " _
                & "errors.<br> Regards, Martin."

            With msgReply
                .HTMLBody = msgTxt & drLogLink & .HTMLBody
                .Subject = .Subject & " - " & Result(4)
                .Display() ' or .Send
            End With
            success = MoveToFolder(Trim(Result(2)), msg)
        End If

    End Sub

    Sub ExpiryMessages(Optional passedMsg As Outlook.MailItem = Nothing)

        If passedMsg Is Nothing Then
            Dim obj As Object
            Dim msg As Outlook.MailItem


            Dim olCurrExplorer As Outlook.Explorer
            Dim olCurrSelection As Outlook.Selection

            olCurrExplorer = Application.ActiveExplorer
            olCurrSelection = olCurrExplorer.Selection

            For Each obj In olCurrSelection
                If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                    msg = obj
                    doOneExpiry(msg)

                End If


            Next
        Else
            doOneExpiry(passedMsg)
        End If




    End Sub


End Class
