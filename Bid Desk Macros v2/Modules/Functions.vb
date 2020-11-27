Imports System.Diagnostics
Imports System.Threading
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions

Partial Class ThisAddIn
    Public sqlInterface As New ClsDatabase(ThisAddIn.server, ThisAddIn.database)


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

            Dim ribbon As MainRibbon = Globals.Ribbons.Ribbon1
            ribbon.UpdateNotDefinedButton()

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






    Private Function RecordWaitTime(receivedTime As Date, completedTime As Date, person As String) As String


        Dim tmpDict As New Dictionary(Of String, String) From {
            {"WaitingFor", person},
            {"Received", receivedTime.ToString("yyyyMMdd HH:mm:ss")},
            {"Completed", completedTime.ToString("yyyyMMdd HH:mm:ss")}
        }


        If sqlInterface.Add_Data(tmpDict, "wait_times") > 0 Then
            Return PrettyString(completedTime - receivedTime)
        Else
            Return "Adding the wait time failed"
        End If


    End Function


    Sub CreateDealRecord(Mail As Outlook.MailItem)
        Dim NewDealForm As New AddDeal(Mail)

        NewDealForm.Show()





    End Sub



    Public Function MakeTicketData(DealDict As Dictionary(Of String, String), email As Outlook.MailItem) As Dictionary(Of String, String)

        Dim requestor As Outlook.ExchangeUser
        requestor = email.Recipients(1).AddressEntry.GetExchangeUser



        MakeTicketData = New Dictionary(Of String, String) From {
            {"Short Description", DealDict("Vendor") & " Bid for " & DealDict("Customer")},
            {"Vendor", DealDict("Vendor")},
            {"Client Name", DealDict("Customer")},
            {"Sales Name", requestor.Name},
            {"Sales Number", requestor.BusinessTelephoneNumber},
            {"Sales Email", requestor.PrimarySmtpAddress},
            {"Description", "This is a copy of a request sent in by email, the original email will be attached. The request has been completed, and the results of these actions will be automatically added when ready."}
        }

    End Function

    Private Function MakeTicketData(DealID As String) As Dictionary(Of String, String)

        Dim tmp = sqlInterface.SelectData_Dict("*", "DealID = '" & DealID & "'")

        Dim DealDict As Dictionary(Of String, String) = tmp(0)

        Dim requestor As Outlook.ExchangeUser

        requestor = Globals.ThisAddIn.Application.Session.CreateRecipient(DealDict("AM")).AddressEntry.GetExchangeUser

        MakeTicketData = New Dictionary(Of String, String) From {
            {"Short Description", DealDict("Vendor") & " Bid for " & DealDict("Customer")},
            {"Vendor", DealDict("Vendor")},
            {"Client Name", DealDict("Customer")},
            {"Sales Name", requestor.Name},
            {"Sales Number", requestor.BusinessTelephoneNumber},
            {"Sales Email", requestor.PrimarySmtpAddress},
            {"Sales Alias", requestor.Alias},
            {"Description", "We have received an expiry notice from the vendor as attached."}
        }


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


    Function IsDealDead(DealID As String) As Boolean

        Dim tmp As String
        tmp = sqlInterface.SelectData("Status", "DealID = '" & DealID & "'")

        'if it's set not to be reminded again - then it can temporarily be "dead" such that no reminders are sent.

        Return (tmp.ToLower.Contains("deal lost") Or tmp.ToLower.Contains("clone pending")) Or tmp.ToLower.Contains("dead")


    End Function

    Function UpdateStatus(DealID As String, NewStatus As String) As Boolean

        If DealExists(DealID) Then
            Try
                sqlInterface.Update_Data("Status = '" & NewStatus & "'", "DealID = '" & DealID & "'")
                sqlInterface.Update_Data("StatusDate = '" & Now() & "'", "DealID = '" & DealID & "'")
                Return True
            Catch
                Return False
            End Try
        Else
            Return False
        End If

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
        If lookupName.ToLower = "amy stonestreet" Then newLookupName = "Kim Woodward"
        If lookupName.ToLower = "andy walsh" Then newLookupName = "Walsh, Andrew"
        If lookupName.ToLower = "nhs solutions" Then newLookupName = "NHSSolutions@Insight.com"
        If lookupName.ToLower = "insight police team" Then newLookupName = "ipt@insight.com"
        If lookupName.ToLower = "josh smith" Then newLookupName = "josh.smith@insight.com"


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

    Public Function WriteSubmitMessage(ByVal DealDetails As Dictionary(Of String, String), Optional fromTicket As Boolean = False) As String

        WriteSubmitMessage = Replace(SubmitMessage, "%DEALID%", DealDetails("DealID"))
        WriteSubmitMessage = Replace(WriteSubmitMessage, "%VENDOR%", DealDetails("Vendor"))

        If Not fromTicket Then
            If DealDetails.ContainsKey("NDT") Then
                WriteSubmitMessage = Replace(WriteSubmitMessage, "%NDT%", Replace(NDTCreateMessage, "%NDT%", DealDetails("NDT")))
            Else
                WriteSubmitMessage = Replace(WriteSubmitMessage, "%NDT%", NoNDTMessage)
            End If
        Else
            WriteSubmitMessage = Replace(WriteSubmitMessage, "%NDT%", Replace(NDTUseMessage, "%NDT%", DealDetails("NDT")))
        End If

        WriteSubmitMessage &= drloglink
    End Function

    Public Function WriteTicketMessage(ByVal DealDetails As Dictionary(Of String, String)) As String

        WriteTicketMessage = Replace(TicketSubmitMessage, "%DEALID%", DealDetails("DealID"))
        WriteTicketMessage = Replace(WriteTicketMessage, "%VENDOR%", DealDetails("Vendor"))



        WriteTicketMessage &= drloglink
    End Function

    Public Function WriteReqMessage(DealID As String, AttBelow As String) As String

        WriteReqMessage = Replace(MoreInfoRequested, "%DEALID%", DealID)
        WriteReqMessage = Replace(WriteReqMessage, "%BELOW%", AttBelow)

        WriteReqMessage = Replace(WriteReqMessage, "%NDT%", GetNDTbyDeal(DealID))

        WriteReqMessage = Replace(WriteReqMessage, "%VENDOR%", GetVendor(DealID))




    End Function



    Public Function WriteInfoMessage(DealID As String, AttachedOrBelow As String) As String

        WriteInfoMessage = Replace(VendorInfoUpdate, "%DEALID%", DealID)
        WriteInfoMessage = Replace(WriteInfoMessage, "%BELOW%", AttachedOrBelow)

        WriteInfoMessage = Replace(WriteInfoMessage, "%NDT%", GetNDTbyDeal(DealID))

        WriteInfoMessage = Replace(WriteInfoMessage, "%VENDOR%", GetVendor(DealID))




    End Function
    Public Function FileFromResource(resource As Byte(), resourceFileName As String) As String



        Dim tempPath As String = Environ("TEMP") & "\bid-desk\" & RandomString(18) & "\"

        If (Not System.IO.Directory.Exists(tempPath)) Then
            System.IO.Directory.CreateDirectory(tempPath)
        End If
        Dim filename As String = tempPath & resourceFileName

        System.IO.File.WriteAllBytes(filename, resource)

        Return filename
    End Function

    Function TicketNumberFromSubject(MsgSubject As String) As String
        Dim ndt As String
        If MsgSubject.StartsWith("[nextDesk]", ThisAddIn.searchType) Then
            ndt = MsgSubject.Substring(InStr(MsgSubject, "#"), 7)
        Else
            ndt = ""
        End If

        Return ndt

    End Function

    Function TicketNumberFromSubject(Msg As Outlook.MailItem) As String
        Dim ndt As String
        If Msg.Subject.StartsWith("[nextDesk]", ThisAddIn.searchType) Then
            ndt = Msg.Subject.Substring(InStr(Msg.Subject, "#"), 7)
        Else
            ndt = ""
        End If

        Return ndt

    End Function

    Sub UpdateTicket(DealID As String, Message As String)
        Dim ticketNum As String = GetNDTbyDeal(DealID)
        If ticketNum <> "" Then
            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket With {
            .TicketNumber = ticketNum,
            .VisibleBrowser = False,
            .TimeOperations = True,
            .TimingOutputFile = ThisAddIn.timingFile
        }


            ndt.UpdateNextDesk(Message)
        End If
    End Sub
End Class



