
Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions
Imports System.Xml

Public Class DealIdent
    Private ReadOnly Message As MailItem
    Public Mode As String

    Private ReadOnly CompleteAutonomy As Boolean

    Public Sub New(message As MailItem, OpMode As String, Optional Autonomy As Boolean = True)
        Me.Message = message
        Me.Mode = OpMode

        Me.CompleteAutonomy = Autonomy

        InitializeComponent()

        Me.DealID.Text = FindDealID(message)

    End Sub

    Private Sub DealID_KeyDown(sender As Object, e As KeyEventArgs) Handles DealID.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1_Click()
        End If
    End Sub


    Public Sub Button1_Click() Handles OKButton.Click
        Me.DialogResult = DialogResult.OK

        DisableButtons()
        Dim tDealID As String = TrimExtended(Me.DealID.Text)

        If tDealID = "" Then
            CloseMe()
            Exit Sub
        End If



        Dim remoteAddMail As Outlook.MailItem

        remoteAddMail = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem)

        remoteAddMail.To = "martin.klefas@insight.com"

        remoteAddMail.Body = "If this is still in the inbox, the deal needs to be added manually"
        Dim xmlFileName As String = ""

        Select Case Mode
            Case "Extended"
                remoteAddMail.Subject = "[dbaddition] Extension to process"
                xmlFileName = WriteXMlOutput(Mode, tDealID)

            Case "Clone"
                remoteAddMail.Subject = "[dbaddition] Extension to process"
                xmlFileName = WriteXMlOutput(Mode, tDealID)

            Case Else


        End Select

        remoteAddMail.Attachments.Add(xmlFileName)
        remoteAddMail.Attachments.Add(Message)
        EnableButtons()

        remoteAddMail.Send()

        Message.Delete()

        CloseMe()
        Exit Sub
    End Sub

    Private Function WriteXMlOutput(Method As String, tDealID As String) As String
        Try
            Dim settings As XmlWriterSettings = New XmlWriterSettings With {
                .Indent = True
            }

            Dim filePath As String = IO.Path.GetTempPath & "dealinfo.xml"

            ' Create XmlWriter.
            Using writer As XmlWriter = XmlWriter.Create(filePath, settings)

                writer.WriteStartDocument()
                writer.WriteStartElement("Deal")




                writer.WriteElementString("DealID", tDealID)
                writer.WriteElementString("Action", Method)



                writer.WriteEndElement()
                writer.WriteEndDocument()

            End Using
            Return filePath
        Catch
            Return ""
        End Try
    End Function

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
            subjAr(i) = TrimExtended(subjAr(i))

            If Len(subjAr(i)) > 4 Then
                If subjAr(i).StartsWith("P00", ThisAddIn.searchType) Or
                    subjAr(i).StartsWith("E00", ThisAddIn.searchType) Or
                    subjAr(i).StartsWith("NQ", ThisAddIn.searchType) Then
                    If Mid(LCase(subjAr(i)), Len(subjAr(i)) - 2, 2) = "-v" Or Mid(LCase(subjAr(i)), Len(subjAr(i)) - 2, 2) = "-0" Then subjAr(i) = Strings.Left(subjAr(i), Len(subjAr(i)) - 3)




                    tempResult = TrimExtended(subjAr(i))
                End If
                If subjAr(i).StartsWith("REGI-", ThisAddIn.searchType) Or
                    subjAr(i).StartsWith("REGE-", ThisAddIn.searchType) Then
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
            If message.SenderEmailAddress.Equals("smart.quotes@techdata.com", ThisAddIn.searchType) And MsgSubject.StartsWith("QUOTE Deal", ThisAddIn.searchType) Then
                tempResult = subjAr(2)


            ElseIf message.SenderEmailAddress.Equals("Neil.Large@westcoast.co.uk", ThisAddIn.searchType) And (MsgSubject.StartsWith("Deal", ThisAddIn.searchType) Or MsgSubject.StartsWith("OPG", ThisAddIn.searchType)) And MsgSubject.ToLower.Contains("for reseller insight direct") Then
                tempResult = subjAr(1)

            End If

        End If

        If tempResult = "" Then
            If message.SenderEmailAddress.Contains("microsoft.com") Then
                tempResult = Mid(message.Subject, InStr(1, message.Subject, "CAS-"), 18)
            End If
        End If




        If tempResult = "" Then
            If message.SenderEmailAddress.ToLower.Equals("donotreply@cisco.com") AndAlso message.Subject.StartsWith("Cisco:") Then
                Try
                    tempResult = CInt(message.Subject.Split(" ")(1))
                Catch
                    tempResult = ""
                End Try
            End If
        End If

        FindDealID = tempResult
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