Imports System.Diagnostics
Imports System.Threading
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions

Partial Class ThisAddIn
    Public sqlInterface As New ClsDatabase(ThisAddIn.server, ThisAddIn.user,
                                   ThisAddIn.database, ThisAddIn.port)


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
                If Not suppressWarnings Then Debug.WriteLine("Couldn't move the mail for some reason...")
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


        If sqlInterface.Add_Data(tmpDict, "wait_times") Then
            Return PrettyString(completedTime - receivedTime)
        Else
            Return "Adding the wait time failed"
        End If


    End Function


    Function CreateDealRecord(Mail As Outlook.MailItem) As Dictionary(Of String, String)
        Dim NewDealForm As New AddDeal
        Dim requestorName As String, Vendor As String, ccNames As String
        Dim ReplyMail As MailItem = Mail.ReplyAll
        Dim tCreateDealRecord As Dictionary(Of String, String)

        If NewDealForm.ShowDialog() = Windows.Forms.DialogResult.OK Then ' Show and wait
            Dim toNames As String(), rName() As String

            toNames = Split(ReplyMail.To, ";") ' Split out each recipient

            If InStr(toNames(0), ",") > 1 Then ' Some email names are "fName, lName" others aren't

                rName = Split(toNames(0), ",")
                requestorName = TrimExtended(rName(1)) & " " & TrimExtended(rName(0))
            Else
                requestorName = TrimExtended(toNames(0))
            End If

            If NewDealForm.DellOption.Checked Then
                Vendor = "Dell"
            ElseIf NewDealForm.HPIOption.Checked Then
                Vendor = "HPI"
            Else
                Vendor = "HPE"
            End If

            ccNames = ReplyMail.CC
            For i = 1 To UBound(toNames) ' append the second and later "to" names to the CC list
                ccNames = ccNames & "; " & toNames(i)
            Next

            Dim bIngram As Byte = BooltoByte(NewDealForm.cIngram.Checked)
            Dim bTechData As Byte = BooltoByte(NewDealForm.cTechData.Checked)
            Dim bWestcoast As Byte = BooltoByte(NewDealForm.cWestcoast.Checked)

            tCreateDealRecord = New Dictionary(Of String, String) From {
                {"AMEmailAddress", Mail.SenderEmailAddress},
                {"AM", requestorName},
                {"Customer", NewDealForm.CustomerName.Text},
                {"Vendor", Vendor},
                {"DealID", NewDealForm.DealID.Text},
                {"Ingram", bIngram},
                {"Techdata", bTechData},
                {"Westcoast", bWestcoast},
                {"CC", ccNames},
                {"Status", "Submitted to Vendor"},
                {"StatusDate", DateTime.Now().ToString("yyyyMMdd HH:mm:ss")},
                {"Date", DateTime.Now().ToString("yyyyMMdd HH:mm:ss")}
            }

            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket(False, True, ThisAddIn.timingFile)

            tCreateDealRecord.Add("NDT", ndt.CreateTicket(1, MakeTicketData(tCreateDealRecord, ReplyMail)).ToString)
            ndt.Move("Public Sector")

            Dim aliases As String = ""
            'add people to notify
            For Each recipient As Outlook.Recipient In ReplyMail.Recipients
                Try
                    aliases &= recipient.AddressEntry.GetExchangeUser.Alias & ";"
                Catch
                    ShoutError("Could not find alias for: " & recipient.ToString)
                End Try
            Next
            ndt.AddToNotify(aliases)

            'update ticket with bid number & original email
            ndt.AttachMail(Mail, "Deal ID  " & tCreateDealRecord("DealID") & "was submitted to " & tCreateDealRecord("Vendor") & " based on the information in the attached email")

            tCreateDealRecord.Remove("AMEmailAddress")

            If sqlInterface.Add_Data(tCreateDealRecord) Then
                tCreateDealRecord.Add("Result", "Success")
            Else
                tCreateDealRecord.Add("Result", "Failed")
            End If

            Return tCreateDealRecord

        Else
            Return New Dictionary(Of String, String) From {
                {"Result", "Cancelled"}
            }



        End If






    End Function



    Private Function MakeTicketData(DealDict As Dictionary(Of String, String), email As Outlook.MailItem) As Dictionary(Of String, String)

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

        Return tmp.ToLower.Contains("deal lost")


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

    Function WriteGreeting(myTime As Date, Optional toName As String = "") As String
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

    Function BooltoByte(ByVal tBoolean As Boolean) As Byte
        If tBoolean Then
            BooltoByte = 1
        Else
            BooltoByte = 0
        End If
    End Function

    Function WriteSubmitMessage(ByVal DealDetails As Dictionary(Of String, String)) As String
        WriteSubmitMessage = Replace(SubmitMessage, "%DEALID%", DealDetails("DealID"))
        WriteSubmitMessage = Replace(WriteSubmitMessage, "%VENDOR%", DealDetails("Vendor"))

        If DealDetails.ContainsKey("NDT") Then
            WriteSubmitMessage = Replace(WriteSubmitMessage, "%NDT%", Replace(NDTCreateMessage, "%NDT%", DealDetails("NDT")))
        Else
            WriteSubmitMessage = Replace(WriteSubmitMessage, "%NDT%", NoNDTMessage)
        End If

        WriteSubmitMessage = WriteSubmitMessage & drloglink
    End Function
End Class



