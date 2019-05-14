Imports System.Diagnostics
Imports Microsoft.Office.Interop.Outlook
Imports MySql.Data.MySqlClient
Imports String_Extensions

Public Class ThisAddIn
    Public Const server As String = "GBMNCDT12830\SQLEXPRESS"
    Public Const user As String = "mklefass"
    Public Const password As String = "nuNDCb4MqmU66j58"
    Public Const database As String = "bids"
    Public Const defaultTable As String = "all_bids"
    Public Const port As Integer = 1433
    Public Const searchType As StringComparison = vbTextCompare
    Public Const timingFile As String = "\\insight.com\root\shared\Sales\public sector\Martin Klefas\Data\NextDesk Metrics\internaltimingfile.csv"



    Sub MoveBasedOnDealID(Optional passedMessage As Outlook.MailItem = Nothing, Optional suppressWarnings As Boolean = False, Optional CompleteAutonomy As Boolean = False)
        Dim MessagesList As New List(Of Outlook.MailItem)

        If passedMessage IsNot Nothing Then
            Dim obj As Object
            Dim msg As Outlook.MailItem

            '  Dim olApp As New Outlook.Application 'new throws security error



            For m = 1 To GetSelection().Count
                obj = GetSelection().Item(m)
                If TypeName(obj) = "MailItem" Then
                    msg = obj
                    MessagesList.Add(msg)
                End If
            Next m
        Else
            MessagesList.Add(passedMessage)
        End If

        Dim DealIDForm As New DealIdent(MessagesList, "Move", CompleteAutonomy)
        DealIDForm.Show()
    End Sub

    Friend Sub FwdHPResponse(Optional passedMessage As Outlook.MailItem = Nothing, Optional SuppressWarnings As Boolean = True)
        Dim MessagesList As New List(Of Outlook.MailItem)
        Dim Autonomy As Boolean

        If passedMessage Is Nothing Then
            Dim obj As Object
            Dim msg As Outlook.MailItem



            For Each obj In GetSelection()
                If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                    msg = obj
                    MessagesList.Add(msg)
                End If

            Next
            Autonomy = False
        Else
            MessagesList.Add(passedMessage)
            Autonomy = True
        End If

        Dim DealIDForm As New DealIdent(MessagesList, "FwdHP", Autonomy)
        DealIDForm.Show()
    End Sub

    Friend Sub MarkedWon()
        Dim obj As Object
        Dim msg As Outlook.MailItem
        Dim MessagesList As New List(Of Outlook.MailItem)

        For Each obj In GetSelection()
            If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                msg = obj
                MessagesList.Add(msg)
            End If

        Next

        Dim DealIDForm As New DealIdent(MessagesList, "MarkedWon")
        DealIDForm.Show()
    End Sub


    Friend Sub MarkDead()
        Dim obj As Object
        Dim msg As Outlook.MailItem
        Dim MessagesList As New List(Of Outlook.MailItem)

        For Each obj In GetSelection()
            If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                msg = obj
                MessagesList.Add(msg)
            End If

        Next
        Dim DealIDForm As New DealIdent(MessagesList, "MarkedWon")
        DealIDForm.Show()
    End Sub


    Friend Sub ExtensionMessage()
        Dim obj As Object
        Dim msg As Outlook.MailItem
        Dim MessagesList As New List(Of Outlook.MailItem)

        For Each obj In GetSelection()
            If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                msg = obj
                MessagesList.Add(msg)
            End If

        Next
        Dim DealIDForm As New DealIdent(MessagesList, "ExtensionMessage")
        DealIDForm.Show()


    End Sub

    Sub FwdPricing(Optional passedMessage As Outlook.MailItem = Nothing, Optional SuppressWarnings As Boolean = True, Optional CompleteAutonomy As Boolean = False)

        Dim MessagesList As New List(Of Outlook.MailItem)
        Dim Autonomy As Boolean

        If passedMessage Is Nothing Then
            Dim obj As Object
            Dim msg As Outlook.MailItem

            For Each obj In GetSelection()
                If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                    msg = obj
                    MessagesList.Add(msg)
                End If

            Next
            Autonomy = False
        Else
            MessagesList.Add(passedMessage)
            Autonomy = True
        End If

        Dim DealIDForm As New DealIdent(MessagesList, "ForwardPricing", Autonomy)
        DealIDForm.Show()

    End Sub



    Sub FwdDRDecision(Optional passedMessage As Outlook.MailItem = Nothing, Optional SuppressWarnings As Boolean = True, Optional CompleteAutonomy As Boolean = False)
        Dim MessagesList As New List(Of Outlook.MailItem)
        Dim Autonomy As Boolean

        If passedMessage Is Nothing Then
            Dim obj As Object
            Dim msg As Outlook.MailItem

            For Each obj In GetSelection()
                If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                    msg = obj
                    MessagesList.Add(msg)
                End If

            Next
            Autonomy = False
        Else
            MessagesList.Add(passedMessage)
            Autonomy = True
        End If

        Dim DealIDForm As New DealIdent(MessagesList, "DRDecision", Autonomy)
        DealIDForm.Show()
    End Sub

    Sub ReplyToBidRequest()



        Dim obj As Object
        Dim msg As Outlook.MailItem, myGreeting As String, success As Boolean
        Dim msgReply As Outlook.MailItem
        Dim Result As Dictionary(Of String, String), rFName As String()



        If GetSelection().Count > 1 Then
            ShoutError("This can only be used with one bid request at a time", False)
            Exit Sub
        End If

        obj = GetCurrentItem()
        If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
            msg = obj
            msgReply = msg.ReplyAll

            Debug.Write(RecordWaitTime(msg.ReceivedTime, Now(), "Me"))
            Result = CreateDealRecord(msg)

            If Result("Result").Equals("Success", searchType) Then

                rFName = Split(Result("AM"))

                myGreeting = WriteGreeting(Now(), CStr(rFName(0)))



                With msgReply
                    .HTMLBody = myGreeting & WriteSubmitMessage(Result) & .HTMLBody
                    .Subject = .Subject & " - " & Result("DealID")
                    .Display() ' or .Send
                End With
                success = MoveToFolder(TrimExtended(Result("AM")), msg)
            End If
        End If
    End Sub

    Sub ExpiryMessages(Optional passedMsg As Outlook.MailItem = Nothing, Optional SuppressWarnings As Boolean = True)

        Dim MessagesList As New List(Of Outlook.MailItem)
        Dim Autonomy As Boolean

        If passedMsg Is Nothing Then
            Dim obj As Object
            Dim msg As Outlook.MailItem

            For Each obj In GetSelection()
                If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                    msg = obj
                    MessagesList.Add(msg)
                End If

            Next
            Autonomy = True
        Else
            MessagesList.Add(passedMsg)
            Autonomy = True
        End If

        Dim DealIDForm As New DealIdent(MessagesList, "Expiry", Autonomy)
        DealIDForm.Show()
    End Sub

    Private Sub Application_NewMailEx(EntryIDCollection As String) Handles Application.NewMailEx
        If Globals.Ribbons.Ribbon1.AutoInbound Then
            Dim frm As New NewMailForm(EntryIDCollection)
            frm.Show()
        Else
            Debug.WriteLine("New Mail - Ignoring")
        End If
    End Sub
End Class
