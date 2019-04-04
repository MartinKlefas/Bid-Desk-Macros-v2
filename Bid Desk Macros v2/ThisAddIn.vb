Imports System.Diagnostics
Imports Microsoft.Office.Interop.Outlook
Imports MySql.Data.MySqlClient

Public Class ThisAddIn
    Public Const server As String = "GBMNCDT12830\SQLEXPRESS"
    Public Const user As String = "mklefass"
    Public Const password As String = "nuNDCb4MqmU66j58"
    Public Const database As String = "bids"
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

    Friend Sub MarkedWon()
        Dim obj As Object
        Dim msg As Outlook.MailItem

        Dim DealID As String, targetFolder As String


        For Each obj In GetSelection()
            If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                msg = obj

                DealID = FindDealID(msg.Subject, msg.Body)
                If DealID = "" Then Exit Sub
                targetFolder = GetFolderbyDeal(DealID, False)


                Dim msgFwdOne As Outlook.MailItem = msg.Forward


                With msgFwdOne
                    .To = MyResolveName(targetFolder).Address
                    .CC = GetCCbyDeal(DealID)
                    .HTMLBody = WriteGreeting(Now(), CStr(Split(targetFolder)(0))) & WonMessage & drloglink & .HTMLBody
                    .Send()
                End With

                MoveToFolder(targetFolder, msg)


            End If

        Next
    End Sub


    Friend Sub MarkDead()
        Dim obj As Object
        Dim msg As Outlook.MailItem

        Dim DealID As String, targetFolder As String


        For Each obj In GetSelection()
            If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                msg = obj

                DealID = FindDealID(msg.Subject, msg.Body)
                If DealID = "" Then Exit Sub
                targetFolder = GetFolderbyDeal(DealID, False)


                Dim msgReplyOne As Outlook.MailItem = msg.ReplyAll


                With msgReplyOne

                    .CC = GetCCbyDeal(DealID)
                    .HTMLBody = WriteGreeting(Now(), CStr(Split(targetFolder)(0))) & DeadMessage & drloglink & .HTMLBody
                    .Send()
                End With

                MoveToFolder(targetFolder, msg)


            End If

        Next
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
        Dim Result As Dictionary(Of String, String), rFName As Object, msgTxt As String



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
                success = MoveToFolder(Trim(Result("AM")), msg)
            End If
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
