Imports System.Diagnostics
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions.StringExtensions

Public Class SavePricingForm
    Private entryIDCollection As String
    Private ReadOnly searchType As StringComparison = ThisAddIn.searchType
    Private NumberOfEmails As Integer
    Private NumberOfFolders As Integer
    Public myContinue As Boolean
    Private startFolder As Outlook.Folder
    Public Sub New()
        myContinue = True
        ' This call is required by the designer.
        InitializeComponent()
        Me.entryIDCollection = My.Settings.entryIDCollection

        My.Settings.entryIDCollection = ""
        My.Settings.Save()
        startFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
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
            Try
                Me.entryIDCollection.Append(email.EntryID)
            Catch
                If Me.entryIDCollection <> "" Then Me.entryIDCollection = Me.entryIDCollection.Substring(0, Me.entryIDCollection.Length - 1)
                Debug.WriteLine("problem adding to entry list")
            End Try
        Next
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Public Sub New(startFolder As Outlook.Folder)
        myContinue = True
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.startFolder = startFolder
        BackgroundWorker2.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim msg As Outlook.MailItem
        Dim newMailFormInstance As New NewMailForm()
startOver:
        Me.NumberOfEmails = entryIDCollection.CountCharacter(",") + 1
        Call UpdateLabel(Me.NumberOfEmails)
        For Each itemID In Split(entryIDCollection, ",")
            If myContinue Then
                Try
                    Dim item = Globals.ThisAddIn.Application.Session.GetItemFromID(itemID)
                    If TypeName(item) = "MailItem" Then
                        msg = item
                        If newMailFormInstance.IsPricing(msg) Then
                            Globals.ThisAddIn.DoOneSharePointUpload(msg)
                        End If
                    End If

                Catch ex2 As System.Exception
                    Debug.WriteLine("Error: " & ex2.ToString())
                    Debug.WriteLine("Error in save pricing routine")
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

    Private Sub UpdateLabel(ByVal [MailsRemaining] As Integer)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Me.Invoke(Sub() UpdateLabel(MailsRemaining))
        Else
            Try
                Me.Label1.Text = "Determining the appropriate action for " & Me.NumberOfEmails & " new emails."
            Catch e As System.Exception
                Debug.Write(e)
                Debug.WriteLine("could not change label")
            End Try

        End If
    End Sub

    Private Sub UpdateLabelTwo(ByVal [MailsRemaining] As Integer)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Me.Invoke(Sub() UpdateLabelTwo(MailsRemaining))
        Else
            Try
                Me.Label2.Text = "Scanned " & Me.NumberOfFolders & " folders."
            Catch e As System.Exception
                Debug.Write(e)
                Debug.WriteLine("could not change label")
            End Try

        End If
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
    Private Sub AddMailItemsFromFolder(ByVal folder As Outlook.Folder, ByRef MessageList As List(Of Outlook.MailItem))
        ' Obtain a Table object that represents the contents of the specified folder.
        Dim table As Outlook.Table = folder.GetTable()

        ' Iterate through each row in the table.
        While Not table.EndOfTable
            Dim row As Outlook.Row = table.GetNextRow()
            Dim item As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(row("EntryID"))
            If TypeOf item Is Outlook.MailItem Then
                MessageList.Add(item)
                Me.NumberOfEmails += 1
                Call UpdateLabel(Me.NumberOfEmails)
            End If

        End While
        Me.NumberOfFolders += 1
        Call UpdateLabelTwo(Me.NumberOfFolders)
        ' Recursively process each subfolder.
        For Each subFolder In folder.Folders
            AddMailItemsFromFolder(subFolder, MessageList)
        Next
    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim MessageList As New List(Of Outlook.MailItem)
        AddMailItemsFromFolder(Me.startFolder, MessageList)
        Me.entryIDCollection = ""

        For Each email As Outlook.MailItem In MessageList
            If Me.entryIDCollection <> "" Then Me.entryIDCollection.Append(",")
            Try
                Me.entryIDCollection.Append(email.EntryID)
            Catch
                If Me.entryIDCollection <> "" Then Me.entryIDCollection = Me.entryIDCollection.Substring(0, Me.entryIDCollection.Length - 1)
                Debug.WriteLine("problem adding to entry list")
            End Try
        Next

        BackgroundWorker1.RunWorkerAsync()
    End Sub
    Private Sub NewMailForm_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        myContinue = False
    End Sub
End Class