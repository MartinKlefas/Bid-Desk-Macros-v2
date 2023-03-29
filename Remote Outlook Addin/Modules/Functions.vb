Imports System.Diagnostics
Imports System.Threading
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions

Partial Class ThisAddIn
    Public Shared Property SearchType As StringComparison = StringComparison.CurrentCultureIgnoreCase

    Public Function MoveToFolder(folderName As String, thisMailItem As MailItem, Optional suppressWarnings As Boolean = False) As Boolean
        Dim mailboxNameString As String
        mailboxNameString = "Martin.Klefas@insight.com"

        Dim olApp As New Outlook.Application
        Dim olNameSpace As Outlook.NameSpace

        Dim olDestFolder As Outlook.MAPIFolder

        If folderName = "" Then folderName = "Not Defined"

        olNameSpace = olApp.GetNamespace("MAPI")

        'OLD: Set olDestFolder = olNameSpace.Folders(mailboxNameString).Folders("Martins Emails").Folders(folderName)
        Try
            olDestFolder = olNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders("Bids").Folders(folderName)
            thisMailItem.Move(olDestFolder)
            MoveToFolder = True
        Catch
            Try
                olDestFolder = olNameSpace.Folders(mailboxNameString).Folders("Inbox").Folders("Bids").Folders.Add(folderName)
                thisMailItem.Move(olDestFolder)
                MoveToFolder = True
            Catch

                ShoutError("Couldn't move the mail for some reason...", suppressWarnings)

                MoveToFolder = False
            End Try
        End Try
    End Function


    Sub ShoutError(errorText As String, Optional SuppressWarnings As Boolean = True)
        Debug.WriteLine(errorText)

        If Not SuppressWarnings Then MsgBox(errorText)


    End Sub

    Public Function GetSelection() As Outlook.Selection
        Dim olCurrExplorer As Outlook.Explorer


        olCurrExplorer = Application.ActiveExplorer
        GetSelection = olCurrExplorer.Selection

    End Function




    Function GetCurrentItem() As Object
        Select Case True
            Case IsExplorer(Application.ActiveWindow)
                GetCurrentItem = Application.ActiveExplorer.Selection.Item(1)
            Case IsInspector(Application.ActiveWindow)
                GetCurrentItem = Application.ActiveInspector.CurrentItem
            Case Else
                GetCurrentItem = Nothing
        End Select
    End Function
    Function IsExplorer(itm As Object) As Boolean
        IsExplorer = (TypeName(itm) = "Explorer")
    End Function
    Function IsInspector(itm As Object) As Boolean
        IsInspector = (TypeName(itm) = "Inspector")
    End Function

    Public Function WriteGreeting(myTime As Date, Optional toName As String = "") As String
        Dim currenthour As Integer

        currenthour = Microsoft.VisualBasic.DateAndTime.Hour(myTime)

        WriteGreeting = "Good "

        If currenthour < 12 Then
            WriteGreeting &= "Morning"
        ElseIf currenthour >= 17 Then
            WriteGreeting &= "Evening"
        Else
            WriteGreeting &= "Afternoon"
        End If

        If toName = "" Or toName.ToLower = "insight" Then
            WriteGreeting &= ","
        Else
            WriteGreeting &= " " & toName & ","
        End If


    End Function

    Function MyResolveName(lookupName As String) As Outlook.ExchangeUser
        Dim oNS As Outlook.NameSpace, newLookupName As String

        newLookupName = lookupName

        Dim noChangeNamesStr As String
        noChangeNamesStr = "lgn.enquiries;lorrae.tomlinson@insight.com;nhssupport;teamwhite.uk;insight acpo tam team;insight met police team;mel wardle;insight capgemini team;ipt@insight.com;_iuk-72-2-brianboys@insight.com"

        Dim nameArry = Split(lookupName, " ")
        If UBound(nameArry) > 0 Then
            newLookupName = nameArry(1) & ", " & nameArry(0)
        End If
        oNS = Application.GetNamespace("MAPI")

        If lookupName.ToLower = "not defined" Or lookupName = "TP2 Enquiries" Then newLookupName = "Klefas, Martin"
        If lookupName.ToLower = "andy walsh" Then newLookupName = "Walsh, Andrew"
        If lookupName.ToLower = "nhs solutions" Then newLookupName = "NHSSolutions@Insight.com"
        If lookupName.ToLower = "insight police team" Then newLookupName = "ipt@insight.com"

        If noChangeNamesStr.ToLower.Contains(lookupName.ToLower) Then newLookupName = lookupName

        Return Globals.ThisAddIn.Application.Session.CreateRecipient(newLookupName).AddressEntry.GetExchangeUser

    End Function

    Public Function BooltoByte(ByVal tBoolean As Boolean) As Byte
        If tBoolean Then
            BooltoByte = 1
        Else
            BooltoByte = 0
        End If
    End Function



End Class



