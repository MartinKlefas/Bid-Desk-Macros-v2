Imports System.Diagnostics
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions.StringExtensions

Public Class BackFromHolsReplyFrm
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
                        If Not msg.Subject.ToLower.StartsWith("automatic reply:") Then
                            Dim reply As MailItem = msg.ReplyAll

                            reply.HTMLBody = htmlMsgStart & Globals.ThisAddIn.WriteGreeting(Now()) & BackFromHolidayMessage & reply.HTMLBody

                            reply.Display()
                        End If

                    End If
                Catch
                    Debug.WriteLine("Could Not find item for some reason")
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
            GoTo startOver
        End If
    End Sub
End Class