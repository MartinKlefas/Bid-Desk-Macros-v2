﻿Imports System.ComponentModel
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook

Public Class DealIdent
    Private MessagesList As List(Of MailItem)
    Private Mode As String
    Private MessageNumber As Integer
    Private CompleteAutonomy As Boolean

    Public Sub New(messagesList As List(Of MailItem), OpMode As String, Optional Autonomy As Boolean = False)
        Me.MessagesList = messagesList
        Me.Mode = OpMode
        Me.MessageNumber = 0
        Me.CompleteAutonomy = Autonomy

        InitializeComponent()
    End Sub

    Private Sub DealID_KeyDown(sender As Object, e As KeyEventArgs) Handles DealID.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1_Click()
        End If
    End Sub

    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        Call Button2_Click()
    End Sub

    Private Sub Button1_Click() Handles OKButton.Click
        Me.DialogResult = DialogResult.OK
        If MessageNumber < MessagesList.Count Then
            DisableButtons()
            Dim tDealID As String = Trim(Me.DealID.Text)
            Dim tMsg As Outlook.MailItem = MessagesList(MessageNumber)
            Select Case Mode
                Case "Move"
                    Call Globals.ThisAddIn.DoOneMove(tMsg, tDealID)
                Case "FwdHP"
                    Call Globals.ThisAddIn.DoOneDistiReminder(tDealID, tMsg)
                    Call Globals.ThisAddIn.DoOneFwd(tDealID, tMsg, HPPublishMessage)
                Case "MarkedWon"
                    Call Globals.ThisAddIn.OneMarkedWon(tMsg, tDealID)
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
                Case "Expiry"
                    Globals.ThisAddIn.DoOneExpiry(tDealID, tMsg, CompleteAutonomy)

                Case Else


            End Select
            EnableButtons()
            MessageNumber += 1
            If MessageNumber < MessagesList.Count Then
                Me.DealID.Text = FindDealID(MessagesList(MessageNumber))
                If CompleteAutonomy Then Call Button1_Click()
            Else
                Me.Close()
            End If
        Else
            Me.Close()
        End If

    End Sub

    Private Sub Button2_Click() Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub DealID_MouseDown1(sender As Object, e As MouseEventArgs) Handles DealID.MouseDown
        If e.Button = MouseButtons.Right Then

            DealID.Text = Trim(My.Computer.Clipboard.GetText)
        End If
    End Sub


    Private Sub DealIdent_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.DialogResult = DialogResult.None
        Me.DealID.Text = FindDealID(MessagesList(MessageNumber))

        If CompleteAutonomy Then Call Button1_Click()

    End Sub


    Public Function FindDealID(message As Outlook.MailItem) As String
        Dim MsgSubject, msgBody As String
        Dim subjAr, bodyAr, bodyLineAr As String()
        Dim i As Integer
        Dim tempResult As String = ""

        MsgSubject = Replace(message.Subject, " ", " ")
        msgBody = message.Body

        subjAr = Split(MsgSubject, " ")



        For i = LBound(subjAr) To UBound(subjAr)
            '~~> This will give you the contents of your email
            '~~> on separate lines
            subjAr(i) = Trim(subjAr(i))

            If Len(subjAr(i)) > 4 Then
                If subjAr(i).StartsWith("P00", ThisAddIn.searchType) Or
                    subjAr(i).StartsWith("E00", ThisAddIn.searchType) Or
                    subjAr(i).StartsWith("NQ", ThisAddIn.searchType) Then
                    If Mid(LCase(subjAr(i)), Len(subjAr(i)) - 2, 2) = "-v" Then subjAr(i) = Strings.Left(subjAr(i), Len(subjAr(i)) - 3)
                    tempResult = Trim(subjAr(i))
                End If
                If subjAr(i).StartsWith("REGI-", ThisAddIn.searchType) Or
                    subjAr(i).StartsWith("REGE-", ThisAddIn.searchType) Then
                    tempResult = Trim(subjAr(i))
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
                    tempResult = Trim(bodyLineAr(2))

                End If
            Next
        End If

        i = 0
        If tempResult = "" AndAlso i + 3 < UBound(subjAr) Then
            If subjAr(i) = "Quote" And subjAr(i + 1) = "Review" And subjAr(i + 2) = "Quote" Then
                tempResult = Trim(subjAr(i + 4))
            End If
            If subjAr(i) = "QUOTE" And subjAr(i + 1) = "Deal" And subjAr(i + 3) = "Version" Then
                tempResult = Trim(subjAr(i + 2))
            End If

        End If

        If tempResult = "" Then
            If message.SenderEmailAddress.Equals("smart.quotes@techdata.com", ThisAddIn.searchType) And MsgSubject.StartsWith("QUOTE Deal", ThisAddIn.searchType) Then
                tempResult = subjAr(2)


            ElseIf message.SenderEmailAddress.Equals("Neil.Large@westcoast.co.uk", ThisAddIn.searchType) And MsgSubject.StartsWith("Deal", ThisAddIn.searchType) And msgSubject.ToLower.Contains("for reseller insight direct") Then
                tempResult = subjAr(1)

            End If

        End If



        FindDealID = tempResult
    End Function
    Sub DisableButtons()
        OKButton.Enabled = False
        Button2.Enabled = False
    End Sub
    Sub EnableButtons()
        OKButton.Enabled = True
        Button2.Enabled = True
    End Sub
End Class