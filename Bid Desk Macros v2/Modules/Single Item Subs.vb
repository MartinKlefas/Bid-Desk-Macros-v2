Partial Class ThisAddIn

    Private Function DoOneExpiry(msg As Outlook.MailItem, Optional SuppressWarnings As Boolean = True) As Boolean

        Dim msgReply As Outlook.MailItem, success As Boolean = True
        Dim DealID As String, TargetFolder As String

        DealID = FindDealID(msg.Subject, msg.Body, True)
        TargetFolder = GetFolderbyDeal(DealID, True)

        If TargetFolder <> "" AndAlso Not IsDealDead(DealID) Then
            msgReply = msg.Forward
            Dim CCList As String = GetCCbyDeal(DealID)
            With msgReply
                .HTMLBody = WriteGreeting(Now(), Split(TargetFolder)(0)) & Replace(Replace(DRExpire, "%dealID%", DealID), "%customer%", GetCustomerbyDeal(DealID)) & .HTMLBody
                .To = TargetFolder
                .CC = CCList
                .Send()
            End With

            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket(False)
            Dim TicketNum As Integer
            Try
                TicketNum = ndt.CreateTicket(1, MakeTicketData(DealID))

                If TicketNum <> 0 AndAlso AddNewTicketToDeal(DealID, TicketNum) <> 1 Then
                    ShoutError("Adding the new ticketID failed", SuppressWarnings)
                    success = False
                End If

                'update notify to include everyone.
                Dim aliases As String = TargetFolder
                aliases = MyResolveName(TargetFolder).GetExchangeUser.Alias
                For Each ccPerson In Split(CCList, ";")
                    Try
                        aliases &= ";" & MyResolveName(ccPerson).GetExchangeUser.Alias
                    Catch
                        ShoutError("Could not find alias for: " & ccPerson, SuppressWarnings)
                    End Try
                Next

                'attach the notification with an explanation
                ndt.AttachMail(msg, "This is the vendor's original expiration notification")

                'Ask the CC List what to do.
                ndt.UpdateNextDesk("Please let me know if you would like to renew " & DealID & " or if it can be marked as Dead/Won in the portal.")

            Catch
                Return False
            End Try

        End If

        Return success
    End Function

    Private Function DoOneFwd(msg As Outlook.MailItem, messageBodyAddition As String, Optional SuppressWarnings As Boolean = True) As Boolean

        Dim fNames As String()

        Dim msgFwdOne As Outlook.MailItem

        Dim DealID As String, TargetFolder As String, myGreeting As String


        DealID = FindDealID(msg.Subject, msg.Body)
        If DealID = "" Then Return False
        RecordWaitTime(GetSubmitTime(DealID), msg.ReceivedTime, GetVendor(DealID))

        TargetFolder = GetFolderbyDeal(DealID, SuppressWarnings)

        msgFwdOne = msg.Forward

        fNames = Split(TargetFolder, " ")
        myGreeting = WriteGreeting(Now(), CStr(fNames(0)))

        With msgFwdOne
            .To = MyResolveName(TargetFolder).Address
            .CC = GetCCbyDeal(DealID)
            .HTMLBody = myGreeting & messageBodyAddition & getfact(DealID) & drloglink & .HTMLBody
            .Send() ' or .Display
        End With

        Return MoveToFolder(TargetFolder, msg, SuppressWarnings)
    End Function

    Private Function DoOneDistiReminder(msg As Outlook.MailItem, Optional SuppressWarnings As Boolean = True) As Boolean

        Dim msgFwdOne As Outlook.MailItem

        Dim DealID As String, myGreeting As String


        DealID = FindDealID(msg.Subject, msg.Body)

        If DealID = "" Then Return False

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


End Class
