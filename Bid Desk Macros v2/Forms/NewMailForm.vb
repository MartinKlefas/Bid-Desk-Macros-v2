Imports System.Diagnostics
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions.StringExtensions

Public Class NewMailForm
    Private entryIDCollection As String
    Private ReadOnly searchType As StringComparison = ThisAddIn.searchType
    Private NumberOfEmails As Integer
    Public myContinue As Boolean

    Public Sub New()
        myContinue = True
        ' This call is required by the designer.
        InitializeComponent()
        Me.entryIDCollection = My.Settings.entryIDCollection
        My.Settings.entryIDCollection = ""
        My.Settings.Save()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(entryIDCollection As String)
        myContinue = True
        InitializeComponent()
        Me.entryIDCollection = entryIDCollection
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Public Sub New(EmailMessages As List(Of Outlook.MailItem))
        myContinue = True
        InitializeComponent()
        Me.entryIDCollection = ""
        For Each email As Outlook.MailItem In EmailMessages
            If Me.entryIDCollection <> "" Then Me.entryIDCollection.Append(",")
            Me.entryIDCollection.Append(email.EntryID)
        Next
        BackgroundWorker1.RunWorkerAsync()
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim msg As Outlook.MailItem
startOver:
        Me.NumberOfEmails = entryIDCollection.CountCharacter(",") + 1
        Call UpdateLabel(Me.NumberOfEmails)
        For Each itemID In Split(entryIDCollection, ",")
            If myContinue Then
                Try
                    Dim item = Globals.ThisAddIn.Application.Session.GetItemFromID(itemID)
                    If TypeName(item) = "MailItem" Then
                        msg = item
                        Select Case FindMessageType(msg)
                            Case "Expiry"
                                Globals.ThisAddIn.ExpiryMessages(msg, True)
                            Case "ExpiryQuote"
                                Globals.ThisAddIn.ExpiryMessages(msg, False)
                            Case "Decision"
                                Globals.ThisAddIn.FwdDRDecision(msg, CompleteAutonomy:=True)
                            Case "Pricing"
                                Globals.ThisAddIn.FwdPricing(msg, CompleteAutonomy:=True)
                            Case "Submission"
                                Globals.ThisAddIn.MoveBasedOnDealID(False, msg, CompleteAutonomy:=True)
                            Case "DBADD"
                                Globals.ThisAddIn.RemoteDBAddition(msg)
                            Case "MoreInfo"
                                Globals.ThisAddIn.ReqMoreInfo(msg, CompleteAutonomy:=True)
                            Case "CiscoApproved"
                                Globals.ThisAddIn.DoCiscoDownload(msg)
                            Case "Forward Request"
                                Globals.ThisAddIn.FwdVendorEmail(msg, CompleteAutonomy:=True)
                            Case "Cisco Submitted"
                                Globals.ThisAddIn.AddAmDetails(msg)
                        End Select

                    End If

                Catch
                    Debug.WriteLine("Could not find item for some reason")
                End Try
                Me.NumberOfEmails -= 1
                Call UpdateLabel(Me.NumberOfEmails)
            End If
        Next
        If My.Settings.entryIDCollection = "" Then

            Call CloseMe()
        Else
            Me.entryIDCollection = My.Settings.entryIDCollection
            My.Settings.entryIDCollection = ""
            My.Settings.Save()
            GoTo startover
        End If
    End Sub

    Private Function FindMessageType(msg As MailItem) As String

        If IsCiscoApproval(msg) Then
            Return "CiscoApproved"
        End If


        If IsDatabaseAdd(msg) Then
            Return "DBADD"
        End If


        If IsExpiryNotice(msg) Then
            Return "Expiry"
        End If
        If IsQuoteExpiryNotice(msg) Then
            Return "ExpiryQuote"
        End If

        If IsDRDecision(msg) Then
            Return "Decision"
        End If
        If IsPricing(msg) Then
            Return "Pricing"
        End If
        If IsDRSubmission(msg) OrElse IsEscalation(msg) OrElse IsPricingApproval(msg) Then
            Return "Submission"
        End If

        If IsMoreInfo(msg) Then
            Return "MoreInfo"
        End If

        If IsForwardRequest(msg) Then
            Return "Forward Request"
        End If

        If isCiscoSubmittedTicket(msg) Then
            Return "Cisco Submitted"
        End If

        Return "Nothing"
    End Function

    Private Function IsCiscoApproval(msg As MailItem) As Boolean
        If msg.Subject.ToLower.StartsWith("deal id:") AndAlso msg.Subject.ToUpper.Contains("INSIGHT NETWORKING SOLUTIONS LIMITED HAS BEEN PROCESSED") Then
            Return True
        ElseIf msg.Subject.ToLower.StartsWith("cisco approved quote, deal id") AndAlso msg.Subject.ToUpper.Contains("INSIGHT NETWORKING SOLUTIONS LIMITED HAS BEEN COMPLETELY APPROVED") Then
            Return True
        Else
            Return False

        End If
    End Function

    Private Function IsDatabaseAdd(msg As MailItem) As Boolean
        Return msg.Subject.ToLower.StartsWith("[dbaddition]")
    End Function
    Private Function IsMoreInfo(msg As MailItem) As Boolean
        Return msg.Subject.ToLower.StartsWith("request incomplete")
    End Function
    Private Function IsPricing(msg As MailItem) As Boolean
        Dim tSubj As String = msg.Subject.ReplaceSpaces()
        If msg.SenderEmailAddress.Equals("smart.quotes@techdata.com", searchType) And tSubj.StartsWith("QUOTE Deal", searchType) Then
            Return True
        ElseIf msg.SenderEmailAddress.Equals("Neil.Large@westcoast.co.uk", searchType) And tSubj.StartsWith("Deal", searchType) And tSubj.ToLower.Contains("for reseller insight direct") Then
            Return True
        ElseIf msg.SenderEmailAddress.Equals("Neil.Large@westcoast.co.uk", searchType) And tSubj.StartsWith("OPG", searchType) And tSubj.ToLower.Contains("for reseller insight direct") Then
            Return True

        Else
            Return False
        End If

    End Function

    Private Function IsDRDecision(msg As MailItem) As Boolean
        Dim tSubj As String = msg.Subject.ReplaceSpaces()
        If msg.SenderEmailAddress.Equals("no_reply@dell.com", searchType) And tSubj.StartsWith("Opportunity", searchType) AndAlso Not tSubj.StartsWith("Opportunity Submitted", searchType) Then
            Return True
        ElseIf tSubj.StartsWith("Deal Registration REGE", searchType) And tSubj.EndsWith("review complete", searchType) Then
            Return True
        ElseIf tSubj.StartsWith("Deal Registration REGI", searchType) AndAlso msg.Body.ToLower.Contains("the review for the deal registration") Then
            Return True
        ElseIf tSubj.StartsWith("Case Processed", searchType) AndAlso msg.SenderEmailAddress.tolower.Contains("hbd-int@microsoft.com") Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsDRSubmission(msg As MailItem) As Boolean
        Dim tSubj As String = msg.Subject.ReplaceSpaces()
        If msg.SenderEmailAddress.Equals("no_reply@dell.com", searchType) And tSubj.StartsWith("Opportunity Submitted", searchType) Then
            Return True
        ElseIf tSubj.StartsWith("Deal Registration REGE", searchType) And tSubj.EndsWith("submitted", searchType) Then
            Return True
        ElseIf tSubj.StartsWith("Deal Registration REGI", searchType) And tSubj.tolower.contains("submitted") Then
            Return True

        ElseIf msg.SenderEmailAddress.Equals("hbd-int@microsoft.com", searchType) And tSubj.StartsWith("Submission Confirmation", searchType) Then

            Return True
        Else
            Return False
        End If



    End Function

    Private Function IsExpiryNotice(newMail As Outlook.MailItem) As Boolean

        If newMail.Subject.StartsWith("Deal Registration", searchType) And newMail.Subject.ToLower.Contains("expiring") Then
            Return True
        ElseIf newMail.Subject.Equals("A Reminder that your Approved Deal is about to Expire", searchType) Then
            Return True

        Else
            Return False
        End If

    End Function
    Private Function IsQuoteExpiryNotice(newMail As Outlook.MailItem) As Boolean
        Dim tmpresult As Boolean

        tmpresult = newMail.SenderEmailAddress.ToLower.Equals("sfdc.support@hpe.com") AndAlso newMail.Subject.StartsWith("your action required", searchType)

        If newMail.Subject.ToLower.StartsWith("your quote expiration reminder mail") Then Return True
        If Not tmpresult Then
            tmpresult = newMail.SenderEmailAddress.ToLower.Equals("donotreply@cisco.com") AndAlso newMail.Body.Contains("will expire in") AndAlso newMail.Body.Contains("days unless action is taken")
        End If
        Return tmpresult
    End Function


    Private Function IsEscalation(newMail As Outlook.MailItem) As Boolean
        If newMail.SenderEmailAddress.ToLower.Contains("dynamics.hpisupport@hp.com") AndAlso newMail.Subject.ToLower.Contains("escalation") Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsPricingApproval(newmail As MailItem) As Boolean
        If (newmail.SenderEmailAddress.ToLower.Contains("noreply.hpintegratedquoting@hp.com") Or
              newmail.SenderEmailAddress.ToLower.Contains("noreply.hpeintegratedquoting@hpe.com")) AndAlso
              newmail.Body.ToLower.Contains("quote request is now ready for viewing") Then

            Return True
        Else
            Return False
        End If
    End Function


    Private Function IsForwardRequest(newmail As MailItem) As Boolean
        Return newmail.Subject.ToLower.StartsWith("[autoforward")
    End Function


    Private Function isCiscoSubmittedTicket(newmail As MailItem) As Boolean

        If newmail.Subject.ToLower.StartsWith("[nextdesk]") Then
            Dim MessageBody As String = newmail.Body
            If MessageBody.ToLower.Contains("by:	martin klefas") And MessageBody.ToLower.Contains("deal id") And
                MessageBody.ToLower.Contains("submitted") Then
                Return True
            End If

        End If

        Return False

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
            Try
                Me.Label1.Text = "Determining the appropriate action for " & Me.NumberOfEmails & " new emails."
            Catch
            End Try

        End If
    End Sub

    Private Sub NewMailForm_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        myContinue = False
    End Sub

    Delegate Sub UpdateLabelCallback(ByVal [MailsRemaining] As Integer)


End Class