﻿Imports System.Diagnostics
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions

Public Class ThisAddIn
    Public Const server As String = "GBMNCDT12889\SQLEXPRESS"
    Public Const user As String = "mklefass"
    Public Const password As String = "nuNDCb4MqmU66j58"
    Public Const database As String = "bids"
    Public Const defaultTable As String = "all_bids"
    Public Const port As Integer = 1433
    Public Const searchType As StringComparison = vbTextCompare
    Public Const timingFile As String = "\\insight.com\root\shared\Sales\public sector\Martin Klefas\Data\NextDesk Metrics\internaltimingfile.csv"


    Sub MoveBasedOnDealID(Optional attach As Boolean = False, Optional passedMessage As Outlook.MailItem = Nothing, Optional CompleteAutonomy As Boolean = False)
        Dim MessagesList As New List(Of Outlook.MailItem)

        If passedMessage Is Nothing Then
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
        Dim Mode As String

        If attach Then
            Mode = "MoveAttach"
        Else
            Mode = "Move"

        End If

        Dim DealIDForm As New DealIdent(MessagesList, Mode, CompleteAutonomy)
        DealIDForm.Show()

    End Sub

    Friend Sub FwdHPResponse(Optional passedMessage As Outlook.MailItem = Nothing)
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
        Dim DealIDForm As New DealIdent(MessagesList, "MarkedDead")
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

    Sub FwdPricing(Optional passedMessage As Outlook.MailItem = Nothing, Optional CompleteAutonomy As Boolean = False)

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
            Autonomy = CompleteAutonomy
        Else
            MessagesList.Add(passedMessage)
            Autonomy = True
        End If

        Dim DealIDForm As New DealIdent(MessagesList, "ForwardPricing", Autonomy)
        DealIDForm.Show()

    End Sub

    Sub FwdVendorUpdate(Optional passedMessage As Outlook.MailItem = Nothing, Optional CompleteAutonomy As Boolean = False)

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
            Autonomy = CompleteAutonomy
        Else
            MessagesList.Add(passedMessage)
            Autonomy = True
        End If

        Dim DealIDForm As New DealIdent(MessagesList, "Forward Vendor Update", Autonomy)
        DealIDForm.Show()

    End Sub

    Sub ReqMoreInfo(Optional passedMessage As Outlook.MailItem = Nothing, Optional CompleteAutonomy As Boolean = False)

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
            Autonomy = CompleteAutonomy
        Else
            MessagesList.Add(passedMessage)
            Autonomy = True
        End If

        Dim DealIDForm As New DealIdent(MessagesList, "ReqMoreInfo", Autonomy)
        DealIDForm.Show()

    End Sub


    Sub FwdDRDecision(Optional passedMessage As Outlook.MailItem = Nothing, Optional CompleteAutonomy As Boolean = False)
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
            Autonomy = CompleteAutonomy
        Else
            MessagesList.Add(passedMessage)
            Autonomy = True
        End If

        Dim DealIDForm As New DealIdent(MessagesList, "DRDecision", Autonomy)
        DealIDForm.Show()
    End Sub

    Public Sub FindOPG()
        Dim obj As Object
        Dim msg As Outlook.MailItem
        Dim MessagesList As New List(Of Outlook.MailItem)


        For Each obj In GetSelection()
            If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
                msg = obj
                MessagesList.Add(msg)
            End If

        Next
        Dim DealIDForm As New DealIdent(MessagesList, "FindOPG", True)
        DealIDForm.Show()

    End Sub


    Sub ReplyToBidRequest()

        Dim obj As Object
        Dim msg As Outlook.MailItem

        If GetSelection().Count > 1 Then
            ShoutError("This can only be used with one bid request at a time", False)
            Exit Sub
        End If

        obj = GetCurrentItem()
        If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
            msg = obj

            RecordWaitTime(msg.ReceivedTime, Now(), "Me")

            CreateDealRecord(msg)

        End If
    End Sub
    Public Sub CloneLater()
        Dim msg As Outlook.MailItem, obj As Object

        If GetSelection().Count > 1 Then
            ShoutError("This can only be used with one bid request at a time", False)
            Exit Sub
        End If

        obj = GetCurrentItem()
        If obj IsNot Nothing AndAlso TypeName(obj) = "MailItem" Then
            msg = obj
            Dim frmRemind As New CloneLater(ReadDate(msg), Msg)
            frmRemind.Show()

        End If
    End Sub
    Public Function ReadDate(msg As Outlook.MailItem) As Date
        Dim targetDate As Date = Today().AddDays(1)

        'find the date within the email
        Dim bodyArr As String() = msg.Body.Split(vbCrLf)

        For Each line As String In bodyArr
            If line.ToLower.Contains("deal expiration date:") Then
                Dim tDateString As String = Trim(line.Substring(InStr(line.ToLower, "deal expiration date:") + 20))

                Dim format() = {"MM/dd/yyyy", "M/d/yyyy"}
                Dim ParsedDate As Date
                If Date.TryParseExact(tDateString, format,
                    System.Globalization.DateTimeFormatInfo.InvariantInfo,
                    Globalization.DateTimeStyles.None, ParsedDate) Then
                    Return ParsedDate
                Else
                    Return targetDate
                End If

            End If
        Next

        Return targetDate
    End Function

    Sub ExpiryMessages(Optional passedMsg As Outlook.MailItem = Nothing, Optional CreateTicket As Boolean = True)

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

        If CreateTicket Then
            Dim DealIDForm As New DealIdent(MessagesList, "Expiry", Autonomy)
            DealIDForm.Show()
        Else
            Dim DealIDForm As New DealIdent(MessagesList, "ExpiryQuote", Autonomy)
            DealIDForm.Show()
        End If


    End Sub

    Private Sub Application_NewMailEx(EntryIDCollection As String) Handles Application.NewMailEx
        If Globals.Ribbons.Ribbon1.AutoInbound Then
            If My.Settings.entryIDCollection = "" Then
                Dim frm As New NewMailForm(EntryIDCollection)
                frm.Show()
            Else
                My.Settings.entryIDCollection = My.Settings.entryIDCollection & "," & EntryIDCollection
                My.Settings.Save()
            End If

        Else
            'Debug.WriteLine("New Mail - Ignoring")
        End If
    End Sub

    Sub DoCiscoDownload(msg As MailItem)
        Dim frm As New BrowserController("DownloadQuote", CiscoQuoteNumber(msg.Subject), msg)
        frm.RunCode()
        frm.Dispose()
    End Sub

    Sub AddAMDetails(msg As MailItem)
        Dim frm As New BrowserController("FindCiscoAM", CiscoQuoteNumberTKT(msg.Body), msg, Globals.ThisAddIn.TicketNumberFromSubject(msg.Subject))
        frm.RunCode()
        frm.Dispose()
    End Sub

    Sub OfferAMDetails(msg As MailItem)
        Dim ndtNumber As String = Globals.ThisAddIn.TicketNumberFromSubject(msg.Subject)

        If ndtNumber <> "" Then
            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket With {
                                .TicketNumber = ndtNumber,
                                .VisibleBrowser = False,
                                .TimeOperations = True,
                                .TimingOutputFile = ThisAddIn.timingFile
                            }

            ndt.UpdateNextDesk(FindAMMessage)

        End If
    End Sub

    Function CiscoQuoteNumber(MessageSubject As String) As String
        Try
            CiscoQuoteNumber = CInt(Strings.Left(MessageSubject.Split(" ")(2), 8))
        Catch
            Try
                CiscoQuoteNumber = CInt(Strings.Left(MessageSubject.Split(" ")(6), 8))
            Catch
                CiscoQuoteNumber = 0
            End Try
        End Try



    End Function
    Function CiscoQuoteNumberTKT(MessageBody As String) As String
        Try
            If MessageBody.ToLower.Contains("by:	martin klefas") And MessageBody.ToLower.Contains("deal id") And
                MessageBody.ToLower.Contains("submitted") Then
                Return TrimExtended(Mid(MessageBody.ToLower, Strings.InStr(MessageBody.ToLower, "deal id") + 8, 10))
            End If
            Return ""
        Catch
            Return ""
        End Try
    End Function


End Class
