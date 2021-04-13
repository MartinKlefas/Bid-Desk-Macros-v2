
Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions

Public Class DealIdent
    Private ReadOnly MessagesList As List(Of MailItem)
    Private Mode As String
    Private MessageNumber As Integer
    Private ReadOnly CompleteAutonomy As Boolean

    Public Sub New(messagesList As List(Of MailItem), OpMode As String, Optional Autonomy As Boolean = False, Optional DealID As String = "")
        InitializeComponent()
        Me.MessagesList = messagesList
        Me.Mode = OpMode
        Me.MessageNumber = 0
        Me.CompleteAutonomy = Autonomy
        Me.DealID.Text = DealID

    End Sub

    Private Sub DealID_KeyDown(sender As Object, e As KeyEventArgs) Handles DealID.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1_Click()
        End If
    End Sub


    Private Sub Button1_Click() Handles OKButton.Click
        Me.DialogResult = DialogResult.OK
        If MessageNumber < MessagesList.Count Then
            DisableButtons()
            Dim tDealID As String = TrimExtended(Me.DealID.Text)

            If tDealID = "" Then
                CloseMe()
                Exit Sub
            End If


            Dim tMsg As Outlook.MailItem = MessagesList(MessageNumber)

            If Not Globals.ThisAddIn.DealExists(tDealID) AndAlso Mode <> "FindOPG" AndAlso Not tMsg.SenderEmailAddress.ToLower.Contains("cisco.com") Then Mode = "Move"

            Select Case Mode
                Case "Move"
                    Call Globals.ThisAddIn.DoOneMove(tMsg, tDealID)
                Case "MoveAttach"
                    Call Globals.ThisAddIn.DoOneAttach(tMsg, tDealID)
                    Call Globals.ThisAddIn.DoOneMove(tMsg, tDealID)
                Case "FwdHP"
                    Call Globals.ThisAddIn.DoOneDistiReminder(tDealID, tMsg)
                    Call Globals.ThisAddIn.DoOneFwd(tDealID, tMsg, HPPublishMessage)
                Case "MarkedWon"
                    Call Globals.ThisAddIn.OneMarkedWon(tMsg, tDealID)
                Case "MarkedDead"
                    Call Globals.ThisAddIn.OneMarkedDead(tMsg, tDealID)
                Case "ExtensionMessage"
                    Call Globals.ThisAddIn.DoOneExtensionMessage(tMsg, tDealID)
                Case "ForwardPricing"
                    If tMsg.Subject.ToLower.Contains("opg") Then
                        Globals.ThisAddIn.DoOneFwd(tDealID, tMsg, opgFwdMessage, True, CompleteAutonomy)
                        Globals.ThisAddIn.UpdateStatus(tDealID, "OPG pricing with AM")
                    Else
                        Globals.ThisAddIn.DoOneFwd(tDealID, tMsg, sqFwdMessage, True, CompleteAutonomy)
                        Globals.ThisAddIn.UpdateStatus(tDealID, "Disti pricing with AM")
                    End If
                Case "DRDecision"
                    Globals.ThisAddIn.DoOneFwd(tDealID, tMsg, drDecision, True, CompleteAutonomy)
                    Globals.ThisAddIn.UpdateStatus(tDealID, "DR Decision with AM")

                    If Globals.ThisAddIn.GetVendor(tDealID) = "Lenovo" Then
                        'Get the Lenovo portal to forward the quote to Distribution.
                        Globals.ThisAddIn.UpdateTicket(tDealID, LenovoBotAttempt)
                        Dim frm As New LenovoBrowserController("SendToDisti", tDealID)
                        If frm.Runcode() Then
                            Globals.ThisAddIn.UpdateTicket(tDealID, LenovoBotSuccess)
                        Else
                            Globals.ThisAddIn.UpdateTicket(tDealID, LenovoBotFail)
                        End If

                    End If

                Case "Expiry"
                    Globals.ThisAddIn.DoOneExpiry(tDealID, tMsg, CompleteAutonomy, True)
                Case "ExpiryQuote"
                    Globals.ThisAddIn.DoOneExpiry(tDealID, tMsg, CompleteAutonomy, False)
                Case "FindOPG"
                    Dim newOPGForm As New NewOPGForm(tDealID)
                    newOPGForm.Show()

                Case "CloneLater"
                    Globals.ThisAddIn.UpdateStatus(tDealID, "Clone requested on deal expiry")
                    Globals.ThisAddIn.CloseAllTickets(tDealID, CloneTicketMessage)
                    Globals.ThisAddIn.DoOneMove(tMsg, tDealID)

                Case "ReqMoreInfo"
                    Call Globals.ThisAddIn.DoOneFwd(tDealID, tMsg, Globals.ThisAddIn.WriteReqMessage(tDealID, "Below"), SuppressWarnings:=True, CompleteAutonomy:=True)

                    Globals.ThisAddIn.UpdateStatus(tDealID, "More Info Requested")

                Case "Forward Vendor Update"
                    Call Globals.ThisAddIn.DoOneFwd(tDealID, tMsg, Globals.ThisAddIn.WriteInfoMessage(tDealID, "Below"), SuppressWarnings:=True, CompleteAutonomy:=True)
                Case Else


            End Select
            EnableButtons()
            MessageNumber += 1
            If MessageNumber < MessagesList.Count Then
                Me.DealID.Text = FindDealID(MessagesList(MessageNumber))
                If CompleteAutonomy Then Call Button1_Click()
            Else
                CloseMe()
            End If
        Else
            CloseMe()
        End If

    End Sub

    Private Sub Button2_Click() Handles Button2.Click
        CloseMe()
    End Sub

    Private Sub DealID_MouseDown1(sender As Object, e As MouseEventArgs) Handles DealID.MouseDown
        If e.Button = MouseButtons.Right Then

            DealID.Text = TrimExtended(My.Computer.Clipboard.GetText)
        End If
    End Sub


    Private Sub DealIdent_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.DialogResult = DialogResult.None
        If Me.DealID.Text = "" Then
            Me.DealID.Text = FindDealID(MessagesList(MessageNumber))
        End If
        If CompleteAutonomy Then Call Button1_Click()

    End Sub


    Public Function FindDealID(message As Outlook.MailItem) As String
        Dim MsgSubject, msgBody As String
        Dim subjAr, bodyAr, bodyLineAr As String()
        Dim i As Integer
        Dim tempResult As String = ""
        HPExtensionSubmitted
        MsgSubject = ReplaceSpaces(message.Subject).TrimExtended
        msgBody = message.Body

        subjAr = Split(MsgSubject, " ")



        For i = LBound(subjAr) To UBound(subjAr)
            '~~> This will give you the contents of your email
            '~~> on separate lines
            subjAr(i) = TrimExtended(subjAr(i))

            If Len(subjAr(i)) > 4 Then
                If subjAr(i).StartsWith("P00", ThisAddIn.searchType) Or
                    subjAr(i).StartsWith("E00", ThisAddIn.searchType) Or
                    subjAr(i).StartsWith("NQ", ThisAddIn.searchType) Then
                    If Mid(LCase(subjAr(i)), Len(subjAr(i)) - 2, 2) = "-v" Or Mid(LCase(subjAr(i)), Len(subjAr(i)) - 2, 2) = "-0" Then subjAr(i) = Strings.Left(subjAr(i), Len(subjAr(i)) - 3)




                    tempResult = TrimExtended(subjAr(i))
                End If
                If subjAr(i).StartsWith("REGI-", ThisAddIn.searchType) Or
                    subjAr(i).StartsWith("REGE-", ThisAddIn.searchType) Or subjAr(i).StartsWith("[REGE-", ThisAddIn.searchType) Then
                    tempResult = TrimExtended(subjAr(i))
                End If
            End If
        Next i

        If tempResult = "" Then
            bodyAr = Split(msgBody, vbCrLf)
            For i = LBound(bodyAr) To UBound(bodyAr)
                '~~> This will give you the contents of your email
                '~~> on separate lines
                If Len(bodyAr(i)) > 8 AndAlso bodyAr(i).StartsWith("Deal ID:", ThisAddIn.searchType) Then
                    bodyLineAr = Split(bodyAr(i))
                    If TrimExtended(bodyLineAr(2)).Contains("https") Then
                        tempResult = Split(bodyLineAr(1), ":").Last
                    Else
                        tempResult = TrimExtended(bodyLineAr(2))
                    End If


                End If
                If Len(bodyAr(i)) > 8 AndAlso bodyAr(i).StartsWith("Quote ID :", ThisAddIn.searchType) Then
                    tempResult = Mid(bodyAr(i), 11, 10)

                End If

            Next
        End If

        i = 0
        If tempResult = "" AndAlso i + 3 < UBound(subjAr) Then
            If subjAr(i) = "Quote" And subjAr(i + 1) = "Review" And subjAr(i + 2) = "Quote" Then
                tempResult = TrimExtended(subjAr(i + 4))
            End If
            If subjAr(i) = "QUOTE" And subjAr(i + 1) = "Deal" And subjAr(i + 3) = "Version" Then
                tempResult = TrimExtended(subjAr(i + 2))
            End If

        End If

        If tempResult = "" Then
            If message.SenderEmailAddress.ToLower.Equals("smart.quotes@techdata.com", ThisAddIn.searchType) And MsgSubject.StartsWith("QUOTE Deal", ThisAddIn.searchType) Then
                tempResult = subjAr(2)


            ElseIf (message.SenderEmailAddress.ToLower.Contains("@westcoast.co.uk")) And (MsgSubject.StartsWith("HPE", ThisAddIn.searchType) Or MsgSubject.StartsWith("Deal", ThisAddIn.searchType) Or MsgSubject.StartsWith("OPG", ThisAddIn.searchType)) Then
                tempResult = subjAr(1)

            End If

        End If

        If tempResult = "" Then
            If message.SenderEmailAddress.Contains("microsoft.com") Then
                Try
                    tempResult = Mid(message.Subject, InStr(1, MsgSubject, "CAS-"), 18)
                Catch
                    Dim mailContents As String() = message.Body.Split(" ")
                    For i = 0 To mailContents.Length
                        If mailContents(i) = "reference" Then
                            tempResult = Mid(mailContents(i + 1), 2)
                            Exit For
                        End If
                    Next
                End Try

            End If
        End If


        If CompleteAutonomy AndAlso tempResult <> "" AndAlso Not Globals.ThisAddIn.DealExists(tempResult) Then
            For Each tAttachment As Attachment In message.Attachments
                tempResult = RipFromFile(tAttachment, tempResult)
            Next
        End If

        If tempResult = "" Then
            If message.SenderEmailAddress.ToLower.Equals("donotreply@cisco.com") Then
                If message.Subject.StartsWith("Cisco:") Then
                    Try
                        tempResult = CInt(message.Subject.Split(" ")(1))
                    Catch
                        tempResult = ""
                    End Try
                ElseIf message.Subject.StartsWith("Deal ID#") Then
                    Try
                        tempResult = CInt(message.Subject.Split(" ")(2))
                    Catch
                        tempResult = ""
                    End Try
                End If

            End If
        End If


        If tempResult = "" Then
            If MsgSubject.Contains("BBR-") Then
                tempResult = Mid(MsgSubject, InStr(1, MsgSubject, "BBR-"), 12)
            ElseIf MsgSubject.Contains("D-0") Then
                tempResult = Mid(MsgSubject, InStr(1, MsgSubject, "D-0"), 11).TrimExtended
            ElseIf MsgSubject.Contains("BRPE") Then
                tempResult = Mid(MsgSubject, InStr(1, MsgSubject, "BRPE"), 14)
            End If

        End If

        FindDealID = TrimExtended(tempResult)
    End Function



    Sub DisableButtons()
        If Me.SpecialMsg.InvokeRequired Then
            Dim d As New CloseMeCallback(AddressOf DisableButtons)
            Try
                Me.Invoke(d, New Object() {})
            Catch
            End Try
        Else
            OKButton.Enabled = False
            Button2.Enabled = False
        End If
    End Sub
    Sub EnableButtons()
        If Me.SpecialMsg.InvokeRequired Then
            Dim d As New CloseMeCallback(AddressOf EnableButtons)
            Try
                Me.Invoke(d, New Object() {})
            Catch
            End Try
        Else
            OKButton.Enabled = True
            Button2.Enabled = True
        End If
    End Sub


    Private Sub CloseMe()

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.SpecialMsg.InvokeRequired Then
            Dim d As New CloseMeCallback(AddressOf CloseMe)
            Me.Invoke(d, New Object() {})
        Else

            Me.Close()

        End If
    End Sub
    Delegate Sub CloseMeCallback()
End Class