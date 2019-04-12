Imports System.ComponentModel
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
                    Else
                        Globals.ThisAddIn.DoOneFwd(tDealID, tMsg, sqFwdMessage, True, CompleteAutonomy)
                    End If
                Case "DRDecision"
                    Globals.ThisAddIn.DoOneFwd(tDealID, tMsg, drDecision, True, CompleteAutonomy)
                Case "Expiry"
                    Globals.ThisAddIn.DoOneExpiry(tDealID, tMsg)

                Case Else


            End Select
            EnableButtons()
            MessageNumber += 1
            Me.DealID.Text = FindDealID(MessagesList(MessageNumber).Subject, MessagesList(MessageNumber).Body)
            If CompleteAutonomy Then Call Button1_Click()
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
        Me.DealID.Text = FindDealID(MessagesList(0).Subject, MessagesList(0).Body)

        If CompleteAutonomy Then Call Button1_Click()

    End Sub


    Public Function FindDealID(MsgSubject As String, msgBody As String) As String
        Dim myAr As String(), i As Integer, myArTwo As String()

        Dim tempResult As String = ""

        myAr = Split(MsgSubject, " ")



        For i = LBound(myAr) To UBound(myAr)
            '~~> This will give you the contents of your email
            '~~> on separate lines
            myAr(i) = Trim(myAr(i))

            If Len(myAr(i)) > 4 Then
                If myAr(i).StartsWith("P00", ThisAddIn.searchType) Or
                    myAr(i).StartsWith("E00", ThisAddIn.searchType) Or
                    myAr(i).StartsWith("NQ", ThisAddIn.searchType) Then
                    If Mid(LCase(myAr(i)), Len(myAr(i)) - 2, 2) = "-v" Then myAr(i) = Strings.Left(myAr(i), Len(myAr(i)) - 3)
                    tempResult = Trim(myAr(i))
                End If
                If myAr(i).StartsWith("REGI-", ThisAddIn.searchType) Or
                    myAr(i).StartsWith("REGE-", ThisAddIn.searchType) Then
                    tempResult = Trim(myAr(i))
                End If
            End If
        Next i

        If tempResult = "" Then
            myAr = Split(msgBody, vbCrLf)
            For i = LBound(myAr) To UBound(myAr)
                '~~> This will give you the contents of your email
                '~~> on separate lines
                If Len(myAr(i)) > 8 Then
                    If myAr(i).StartsWith("Deal ID:", ThisAddIn.searchType) Then
                        myArTwo = Split(myAr(i))
                        tempResult = Trim(myArTwo(2))
                    End If
                End If
            Next
        End If

        If tempResult = "" AndAlso i + 3 < UBound(myAr) Then
            If myAr(i) = "Quote" And myAr(i + 1) = "Review" And myAr(i + 2) = "Quote" Then
                tempResult = Trim(myAr(i + 4))
            End If
            If myAr(i) = "QUOTE" And myAr(i + 1) = "Deal" And myAr(i + 3) = "Version" Then
                tempResult = Trim(myAr(i + 2))
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