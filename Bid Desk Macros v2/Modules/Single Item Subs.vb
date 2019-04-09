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
                Try
                    .Display() '.Send()
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

    Private Function DoOneFwd(msg As Outlook.MailItem, messageBodyAddition As String, Optional SuppressWarnings As Boolean = True, Optional CompleteAutonomy As Boolean = False) As Boolean

        Dim fNames As String()

        Dim msgFwdOne As Outlook.MailItem

        Dim DealID As String, TargetFolder As String, myGreeting As String


        DealID = FindDealID(msg.Subject, msg.Body, CompleteAutonomy)
        If DealID = "" OrElse dealExists(DealID) Then
            Return False
        End If

        RecordWaitTime(GetSubmitTime(DealID), msg.ReceivedTime, GetVendor(DealID))

        TargetFolder = GetFolderbyDeal(DealID, SuppressWarnings)

        msgFwdOne = msg.Forward

        fNames = Split(TargetFolder, " ")
        myGreeting = WriteGreeting(Now(), CStr(fNames(0)))

        With msgFwdOne
            .To = MyResolveName(TargetFolder).PrimarySmtpAddress
            .CC = GetCCbyDeal(DealID)
            .HTMLBody = myGreeting & messageBodyAddition & GetFact(DealID) & drloglink & .HTMLBody
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
