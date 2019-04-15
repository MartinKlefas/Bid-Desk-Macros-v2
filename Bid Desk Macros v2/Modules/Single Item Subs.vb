Partial Class ThisAddIn

    Public Function DoOneExpiry(DealID As String, msg As Outlook.MailItem, Optional SuppressWarnings As Boolean = True) As Boolean

        Dim msgReply As Outlook.MailItem, success As Boolean = True
        Dim TargetFolder As String


        TargetFolder = GetFolderbyDeal(DealID, True)

        If TargetFolder <> "" AndAlso Not IsDealDead(DealID) Then
            msgReply = msg.Forward
            Dim CCList As String = GetCCbyDeal(DealID)
            With msgReply
                .HTMLBody = WriteGreeting(Now(), Split(TargetFolder)(0)) & Replace(Replace(DRExpire, "%dealID%", DealID), "%customer%", GetCustomerbyDeal(DealID)) & .HTMLBody
                .To = TargetFolder
                .CC = CCList
                Try
                    .Send()
                Catch
                    .Display()
                End Try
            End With

            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket(False)
            Dim TicketNum As Integer
            Try
                Dim DealData As Dictionary(Of String, String) = MakeTicketData(DealID)
                TicketNum = ndt.CreateTicket(1, DealData)

                If TicketNum = 0 Then
                    ShoutError("Adding the new ticketID failed", SuppressWarnings)
                    success = False
                Else
                    ndt.Move("Public Sector")
                    'update notify to include everyone.
                    Dim aliases As String = DealData("Sales Alias")
                    For Each ccPerson In Split(CCList, ";")
                        If ccPerson <> "" Then
                            Try
                                aliases &= ";" & MyResolveName(ccPerson).Alias
                            Catch
                                ShoutError("Could not find alias for: " & ccPerson, SuppressWarnings)
                            End Try
                        End If
                    Next

                    'attach the notification with an explanation
                    ndt.AttachMail(msg, "This is the vendor's original expiration notification")

                    'Ask the CC List what to do.
                    ndt.UpdateNextDesk("Please let me know if you would like to renew " & DealID & " or if it can be marked as Dead/Won in the portal.")
                End If

                AddNewTicketToDeal(DealID, TicketNum)
                UpdateStatus(DealID, "Expiration notice with AM")

                Try
                    Globals.ThisAddIn.MoveToFolder(TargetFolder, msg, SuppressWarnings)
                Catch ex As Exception
                    ShoutError("Could not move to folder: " & TargetFolder, SuppressWarnings)
                End Try
            Catch
                Return False
            End Try

        End If

        Return success
    End Function

    Public Function DoOneFwd(DealID As String, msg As Outlook.MailItem, messageBodyAddition As String, Optional SuppressWarnings As Boolean = True, Optional CompleteAutonomy As Boolean = False) As Boolean

        Dim fNames As String()

        Dim msgFwdOne As Outlook.MailItem

        Dim TargetFolder As String, myGreeting As String




        RecordWaitTime(GetSubmitTime(DealID), msg.ReceivedTime, GetVendor(DealID))

        TargetFolder = GetFolderbyDeal(DealID, SuppressWarnings)

        msgFwdOne = msg.Forward

        fNames = Split(TargetFolder, " ")
        myGreeting = WriteGreeting(Now(), CStr(fNames(0)))

        With msgFwdOne
            .To = MyResolveName(TargetFolder).PrimarySmtpAddress
            .CC = GetCCbyDeal(DealID)
            .HTMLBody = myGreeting & messageBodyAddition & GetFact(DealID) & drloglink & .HTMLBody
            Try
                .Send()  'or .Display
            Catch
                .Display()
            End Try
        End With

        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket(False) With {
            .TicketNumber = GetNDTbyDeal(DealID)
        }
        Dim browser As OpenQA.Selenium.Chrome.ChromeDriver = ndt.GiveMeChrome(False)

        If ndt.TicketNumber <> 0 AndAlso Not ndt.IsClosed(browser) Then
            ndt.AttachMail(msg, "The attached pricing has been received from distritbution", browser)
        End If

        If Not CompleteAutonomy AndAlso MsgBox("Would you like to close the ticket", vbYesNo) = vbYes Then
            ndt.CloseTicket(browser:=browser)
        End If

        Return MoveToFolder(TargetFolder, msg, SuppressWarnings)
    End Function

    Public Function DoOneDistiReminder(DealID As String, msg As Outlook.MailItem, Optional SuppressWarnings As Boolean = True) As Boolean

        Dim msgFwdOne As Outlook.MailItem

        Dim myGreeting As String




        Try
            If IsWestcoast(DealID) Then
                msgFwdOne = msg.Forward
                msgFwdOne.To = "quotes@westcoast.co.uk"
                myGreeting = WriteGreeting(Now(), "All")
                msgFwdOne.Send()
            End If
        Catch
            Return False
        End Try

        Try
            If IsIngram(DealID) Then
                msgFwdOne = msg.Forward
                msgFwdOne.To = "Insight.UK@ingrammicro.co.uk"
                myGreeting = WriteGreeting(Now(), "All")
                msgFwdOne.Send()
            End If
        Catch
            Return False
        End Try

        Try
            If IsTechData(DealID) Then
                msgFwdOne = msg.Forward
                msgFwdOne.To = "insightsales@techdata.co.uk"
                myGreeting = WriteGreeting(Now(), "All")
                msgFwdOne.Send()
            End If
        Catch
            Return False
        End Try

        Return True

    End Function
    Public Sub DoOneMove(Message As Outlook.MailItem, DealID As String)
        Dim TargetFolder As String
        If DealExists(DealID) Then

            TargetFolder = GetFolderbyDeal(DealID)
        Else
            TargetFolder = "Not Defined"
        End If
        Dim success As Boolean = MoveToFolder(TargetFolder, Message)

    End Sub

    Public Sub OneMarkedWon(message As Outlook.MailItem, DealID As String)
        Dim TargetFolder As String
        TargetFolder = GetFolderbyDeal(DealID, False)


        Dim msgFwdOne As Outlook.MailItem = message.Forward


        With msgFwdOne
            .To = MyResolveName(targetFolder).PrimarySmtpAddress
            .CC = GetCCbyDeal(DealID)
            .HTMLBody = WriteGreeting(Now(), CStr(Split(targetFolder)(0))) & WonMessage & drloglink & .HTMLBody
            .Send()
        End With
        UpdateStatus(DealID, "Marked as Won in the Portal")
        MoveToFolder(TargetFolder, message)
    End Sub
    Public Sub OneMarkedDead(msg As Outlook.MailItem, DealID As String)
        Dim TargetFolder As String

        If DealID = "" Then Exit Sub
        targetFolder = GetFolderbyDeal(DealID, False)


        Dim msgReplyOne As Outlook.MailItem = msg.ReplyAll


        With msgReplyOne

            .CC = GetCCbyDeal(DealID)
            .HTMLBody = WriteGreeting(Now(), CStr(Split(targetFolder)(0))) & DeadMessage & drloglink & .HTMLBody
            .Send()
        End With

        MoveToFolder(targetFolder, msg)


    End Sub

    Sub DoOneExtensionMessage(msg As Outlook.MailItem, DealID As String)
        Dim msgReply As Outlook.MailItem = msg.ReplyAll



        Dim AM As String = GetAMbyDeal(DealID)
        Dim replyText, myGreeting As String

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
        UpdateStatus(DealID, "Extension requested online")


        MoveToFolder(AM, msg)
    End Sub

End Class
