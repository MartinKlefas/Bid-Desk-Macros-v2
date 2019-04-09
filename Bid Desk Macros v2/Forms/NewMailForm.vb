Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions.StringExtensions

Public Class NewMailForm
    Private entryIDCollection As String
    Private searchType As StringComparison = ThisAddIn.searchType

    Public Sub New(entryIDCollection As String)
        Me.entryIDCollection = entryIDCollection
        Me.Label1.Text = "Determining the appropriate action for " & entryIDCollection.CountCharacter(",") & " new emails."
        InitializeComponent()
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim msg As Outlook.MailItem

        For Each itemID In Split(entryIDCollection, ",")
            Dim item = Globals.ThisAddIn.Application.Session.GetItemFromID(itemID)
            If TypeName(item) = "MailItem" Then
                msg = item
                If isExpiryNotice(msg) Then
                    Globals.ThisAddIn.ExpiryMessages(msg, True)
                End If
                If isDRDecision(msg) Then
                    ' Globals.ThisAddIn.FwdDRDecision(msg)
                End If
                If isPricing(msg) Then
                    ' Globals.ThisAddIn.FwdPricing(msg, True)
                End If
            End If
        Next

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
        ElseIf newMail.Subject.tolower.Contains("your quote expiration reminder mail") Then
            Return True
        Else
            Return False
        End If

    End Function
End Class