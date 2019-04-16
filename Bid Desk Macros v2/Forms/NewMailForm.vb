﻿Imports System.Diagnostics
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions.StringExtensions

Public Class NewMailForm
    Private entryIDCollection As String
    Private searchType As StringComparison = ThisAddIn.searchType
    Private NumberOfEmails As Integer

    Public Sub New(entryIDCollection As String)
        InitializeComponent()
        Me.entryIDCollection = entryIDCollection
        Me.NumberOfEmails = entryIDCollection.CountCharacter(",") + 1
        Me.Label1.Text = "Determining the appropriate action for " & Me.NumberOfEmails & " new emails."

        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Public Sub New(EmailMessages As List(Of Outlook.MailItem))
        InitializeComponent()
        Me.entryIDCollection = ""
        For Each email As Outlook.MailItem In EmailMessages
            If Me.entryIDCollection <> "" Then Me.entryIDCollection.Append(",")
            Me.entryIDCollection.Append(email.EntryID)
        Next
        Me.NumberOfEmails = entryIDCollection.CountCharacter(",") + 1
        Me.Label1.Text = "Determining the appropriate action for " & Me.NumberOfEmails + 1 & " new emails."

        BackgroundWorker1.RunWorkerAsync()
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim msg As Outlook.MailItem

        For Each itemID In Split(entryIDCollection, ",")
            Try
                Dim item = Globals.ThisAddIn.Application.Session.GetItemFromID(itemID)
                If TypeName(item) = "MailItem" Then
                    msg = item
                    If IsExpiryNotice(msg) Then
                        Globals.ThisAddIn.ExpiryMessages(msg, True)
                    End If
                    If IsDRDecision(msg) Then
                        Globals.ThisAddIn.FwdDRDecision(msg, SuppressWarnings:=True, CompleteAutonomy:=True)
                    End If
                    If IsPricing(msg) Then
                        Globals.ThisAddIn.FwdPricing(msg, SuppressWarnings:=True, CompleteAutonomy:=True)
                    End If
                End If
            Catch
                Debug.WriteLine("Could not find item for some reason")
            End Try
            Me.NumberOfEmails -= 1
            Call UpdateLabel(Me.NumberOfEmails)
        Next

        Call CloseMe()
    End Sub

    Private Function IsPricing(msg As MailItem) As Boolean
        If msg.SenderEmailAddress.Equals("smart.quotes@techdata.com", searchType) And msg.Subject.StartsWith("QUOTE Deal", searchType) Then
            Return True
        ElseIf msg.SenderEmailAddress.Equals("Neil.Large@westcoast.co.uk", searchType) And msg.Subject.StartsWith("Deal", searchType) And msg.Subject.ToLower.Contains("for reseller insight direct") Then
            Return True

        Else
            Return False
        End If

    End Function

    Private Function IsDRDecision(msg As MailItem) As Boolean
        If msg.SenderEmailAddress.Equals("no_reply@dell.com", searchType) And msg.Subject.StartsWith("Opportunity", searchType) Then
            Return True
        ElseIf msg.Subject.StartsWith("Deal Registration REGE", searchType) And msg.Subject.EndsWith("review complete", searchType) Then
            Return True
        ElseIf msg.Subject.StartsWith("Deal Registration REGI", searchType) AndAlso msg.Body.ToLower.Contains("the review for the deal registration") Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsExpiryNotice(newMail As Outlook.MailItem) As Boolean

        If newMail.Subject.StartsWith("Deal Registration", searchType) And newMail.Subject.EndsWith("Expiring", searchType) Then
            Return True
        ElseIf newMail.Subject.Equals("A Reminder that your Approved Deal is about to Expire", searchType) Then
            Return True
        ElseIf newMail.Subject.ToLower.Contains("your quote expiration reminder mail") Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub CloseMe()

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New CloseMeCallback(AddressOf CloseMe)
            Me.Invoke(d, New Object() {})
        Else

            Me.Close()

        End If
    End Sub
    Delegate Sub CloseMeCallback()
    Private Sub UpdateLabel(ByVal [MailsRemaining] As Integer)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New UpdateLabelCallback(AddressOf UpdateLabel)
            Me.Invoke(d, New Object() {[MailsRemaining]})
        Else

            Me.Label1.Text = "Determining the appropriate action for " & Me.NumberOfEmails & " new emails."

        End If
    End Sub
    Delegate Sub UpdateLabelCallback(ByVal [MailsRemaining] As Integer)
End Class