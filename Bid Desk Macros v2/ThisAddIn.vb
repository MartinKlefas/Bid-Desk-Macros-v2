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


        For m = 1 To GetSelection().Count
            obj = GetSelection().Item(m)
            If TypeName(obj) = "MailItem" Then
                msg = obj
                DealID = FindDealID(msg.Subject, msg.Body)
                If DealID = "" Then Exit Sub
                targetFolder = GetFolderbyDeal(DealID, suppressWarnings)

                success = MoveToFolder(targetFolder, msg)
            End If
        Next m
    End Sub

    Friend Sub FwdHPResponse(Optional passedMessage As Outlook.MailItem = Nothing, Optional SuppressWarnings As Boolean = True)
        If passedMessage Is Nothing Then
            Dim obj As Object
            Dim msg As Outlook.MailItem



            For Each obj In GetSelection()
                If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                    msg = obj
                    DoOneDistiReminder(msg, SuppressWarnings)
                    DoOneFwd(msg, HPPublishMessage, SuppressWarnings)
                End If

            Next
        Else
            DoOneDistiReminder(passedMessage, SuppressWarnings)
            DoOneFwd(passedMessage, HPPublishMessage, SuppressWarnings)
        End If

    End Sub

    Friend Sub ExtensionMessage()
        Dim obj As Object
        Dim msg As Outlook.MailItem, DealID As String
        Dim msgReply As Outlook.MailItem

        Dim replyText As String, myGreeting As String, AM As String



        For Each obj In GetSelection()
            If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                msg = obj
                msgReply = msg.ReplyAll

                DealID = FindDealID(msg.Subject, msg.Body)

                AM = GetAMbyDeal(DealID)

                If GetVendor(DealID).ToLower.Contains("hp") Then
                    replyText = HPExtensionSubmitted
                Else
                    replyText = DellExtensionSubmitted

                End If

                Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket With {
                    .TicketNumber = GetNDTbyDeal(DealID)
                }

                ndt.AttachMail(msg, "Request to extend the DR")
                ndt.CloseTicket("As requested, an extension has been requested on the vendor portal.")

                myGreeting = WriteGreeting(Now(), AM.Split(" ")(0))

                msgReply.HTMLBody = myGreeting & replyText & msgReply.HTMLBody

                msgReply.Send()


            End If
        Next

    End Sub

    Sub FwdPricing(Optional passedMessage As Outlook.MailItem = Nothing, Optional SuppressWarnings As Boolean = True)

        If passedMessage Is Nothing Then
            Dim obj As Object
            Dim msg As Outlook.MailItem




            For Each obj In GetSelection()
                If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                    msg = obj
                    If msg.Subject.ToLower.Contains("opg") Then
                        DoOneFwd(msg, opgFwdMessage, SuppressWarnings)
                    Else
                        DoOneFwd(msg, sqFwdMessage, SuppressWarnings)
                    End If


                End If


            Next
        Else
            If passedMessage.Subject.ToLower.Contains("opg") Then
                DoOneFwd(passedMessage, opgFwdMessage, SuppressWarnings)
            Else
                DoOneFwd(passedMessage, sqFwdMessage, SuppressWarnings)
            End If


        End If

    End Sub



    Sub FwdDRDecision(Optional passedMessage As Outlook.MailItem = Nothing, Optional SuppressWarnings As Boolean = True)
        If passedMessage Is Nothing Then
            Dim obj As Object
            Dim msg As Outlook.MailItem




            For Each obj In GetSelection()
                If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                    msg = obj
                    DoOneFwd(msg, drDecision, SuppressWarnings)

                End If


            Next
        Else
            DoOneFwd(passedMessage, drDecision, SuppressWarnings)


        End If
    End Sub

    Sub ReplyToBidRequest()



        Dim obj As Object
        Dim msg As Outlook.MailItem, myGreeting As String, success As Boolean
        Dim msgReply As Outlook.MailItem
        Dim Result As Object, rFName As Object, msgTxt As String



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



            rFName = Split(Result(2))

            myGreeting = WriteGreeting(Now(), CStr(rFName(0)))

            msgTxt = myGreeting & "<br>&nbsp;I've created the below for you with " & Result(3) & " (ref: " _
                & Result(4) _
                & ").<br>&nbsp;Please check that everything is correct and let me know asap if there are any " _
                & "errors.<br> Regards, Martin."

            With msgReply
                .HTMLBody = msgTxt & drloglink & .HTMLBody
                .Subject = .Subject & " - " & Result(4)
                .Display() ' or .Send
            End With
            success = MoveToFolder(Trim(Result(2)), msg)
        End If

    End Sub

    Sub ExpiryMessages(Optional passedMsg As Outlook.MailItem = Nothing, Optional SuppressWarnings As Boolean = True)

        If passedMsg Is Nothing Then
            Dim obj As Object
            Dim msg As Outlook.MailItem


            For Each obj In GetSelection()
                If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                    msg = obj
                    If Not DoOneExpiry(msg, SuppressWarnings) Then
                        ShoutError("There was an error processing this expiration", SuppressWarnings)

                    End If
                End If


            Next
        Else
            If Not DoOneExpiry(passedMsg, SuppressWarnings) Then
                ShoutError("There was an error processing this expiration")
            End If
        End If




    End Sub


End Class
