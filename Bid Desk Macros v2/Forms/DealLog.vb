Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions

Public Class AddDeal
    Private ReadOnly mail As MailItem
    Private myContinue As Boolean

    Public Sub New(mail As MailItem)
        InitializeComponent()
        Me.mail = mail
    End Sub

    Private Sub CommandButton1_Click() Handles OKButton.Click
        CustomerName.Text = TrimExtended(CustomerName.Text)
        DealID.Text = TrimExtended(DealID.Text)

        DisableButtons()

        myContinue = True
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Button2_Click() Handles tCancelButton.Click

        Me.DialogResult = DialogResult.Cancel
        Me.Hide()
    End Sub

    Private Sub UserForm_Activate() Handles Me.Activated


        Dim strClip As String

        If Me.DealID.Text = "" And Me.CustomerName.Text = "" Then ' Only read clipboard if nothing has been typed into the boxes already
            strClip = My.Computer.Clipboard.GetText

            Me.DealID.Text = FindDealID(strClip)

            Me.CustomerName.Text = FindCustomer(strClip)
            Select Case FindVendor(strClip)
                Case "HPI"
                    Call CheckOnly(HPIOption)
                Case "HPE"
                    Call CheckOnly(HPEOption)
                Case "Dell"
                    Call CheckOnly(DellOption)
                Case "Microsoft"
                    Call CheckOnly(btnMS)
                Case "Lenovo"
                    Call CheckOnly(LenovoOption)


            End Select
        End If
    End Sub
    Private Sub CheckOnly(toCheck As RadioButton)
        For Each tControl As Control In VendorGroupBox.Controls
            If TypeName(tControl) = "RadioButton" Then
                Dim rButton As RadioButton = tControl
                rButton.Checked = False
            End If
        Next
        toCheck.Checked = True
    End Sub


    Private Sub TextBox1_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles CustomerName.KeyDown
        If e.KeyCode = Keys.Enter Then
            CommandButton1_Click()
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles CustomerName.KeyDown
        If e.KeyCode = Keys.Enter Then
            CommandButton1_Click()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles tCancelButton.Click
        CloseMe()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim requestorName As String, Vendor As String, ccNames As String
        Dim ReplyMail As MailItem = mail.ReplyAll
        Dim ReplyMailTwo As MailItem = mail.ReplyAll
        Dim tCreateDealRecord As Dictionary(Of String, String)


        Dim toNames As String(), rName() As String

        UpdateTitle("Preparing Details...")

        toNames = Split(ReplyMail.To, ";") ' Split out each recipient

        If InStr(toNames(0), ",") > 1 Then ' Some email names are "fName, lName" others aren't

            rName = Split(toNames(0), ",")
            requestorName = TrimExtended(rName(1)) & " " & TrimExtended(rName(0))
        Else
            requestorName = TrimExtended(toNames(0))
        End If

        If Me.DellOption.Checked Then
            Vendor = "Dell"
        ElseIf Me.HPIOption.Checked Then
            Vendor = "HPI"
        ElseIf Me.HPEOption.Checked Then
            Vendor = "HPE"
        ElseIf Me.lenovooption.checked Then
            Vendor = "Lenovo"
        Else
            Vendor = "Microsoft"
        End If

        ccNames = ReplyMail.CC
        For i = 1 To UBound(toNames) ' append the second and later "to" names to the CC list
            ccNames = ccNames & "; " & toNames(i)
        Next

        Dim bIngram As Byte = Globals.ThisAddIn.BooltoByte(Me.cIngram.Checked)
        Dim bTechData As Byte = Globals.ThisAddIn.BooltoByte(Me.cTechData.Checked)
        Dim bWestcoast As Byte = Globals.ThisAddIn.BooltoByte(Me.cWestcoast.Checked)

        tCreateDealRecord = New Dictionary(Of String, String) From {
                {"AMEmailAddress", mail.SenderEmailAddress},
                {"AM", requestorName},
                {"Customer", Me.CustomerName.Text},
                {"Vendor", Vendor},
                {"DealID", Me.DealID.Text},
                {"Ingram", bIngram},
                {"Techdata", bTechData},
                {"Westcoast", bWestcoast},
                {"CC", ccNames},
                {"Status", "Submitted to Vendor"},
                {"StatusDate", DateTime.Now().ToString("yyyyMMdd HH:mm:ss")},
                {"Date", DateTime.Now().ToString("yyyyMMdd HH:mm:ss")}
            }



        If Not myContinue Then
            Call ExitEarly()
            Exit Sub
        End If

        If Not DoNewCreation(tCreateDealRecord, ReplyMail, ReplyMailTwo) Then

            Call ExitEarly()
            Exit Sub
        End If

        CloseMe()

    End Sub


    Public Function DoNewCreation(DealData As Dictionary(Of String, String), ByRef replyMail As Outlook.MailItem, ByRef replyMailTwo As Outlook.MailItem) As Boolean

        myContinue = True
        'Dim newNDT As Integer


        If Not DealData.ContainsKey("NDT") OrElse DealData("NDT") = "" Then
            'UpdateTitle("Creating Ticket...")




            'newNDT = ndt.CreateTicket(1, Globals.ThisAddIn.MakeTicketData(DealData, replyMail))

            'If Not myContinue Then
            '    Return False
            '    Exit Function
            'End If

            'If newNDT = 0 Or newNDT = 404 Then ' retry on first fail
            '    newNDT = ndt.CreateTicket(1, Globals.ThisAddIn.MakeTicketData(DealData, replyMail))
            'End If




            'If newNDT <> 0 And newNDT <> 404 Then ' continue on second

            '    If Not DealData.ContainsKey("NDT") Then DealData.Add("NDT", newNDT)

            '    If DealData("NDT") = "" Then DealData("NDT") = newNDT

            '    ndt.Move("Public Sector - Special Bid")

            '    If Not myContinue Then
            '        Return False
            '        Exit Function
            '    End If

            '    Dim aliases As String = ""
            '    'add people to notify
            '    For Each recipient As Outlook.Recipient In replyMail.Recipients
            '        Try
            '            aliases &= recipient.AddressEntry.GetExchangeUser.Alias & ";"
            '        Catch
            '            Globals.ThisAddIn.ShoutError("Could not find alias for: " & recipient.ToString)
            '        End Try
            '    Next

            '    UpdateTitle("Adding Notify...")

            '    ndt.AddToNotify(aliases)
            '    If Not myContinue Then
            '        Return False
            '        Exit Function
            '    End If

            '    UpdateTitle("Attaching Info...")

            '    'update ticket with bid number & original email
            '    ndt.AttachMail(mail, "Deal ID  " & DealData("DealID") & " was submitted to " & DealData("Vendor") & " based on the information in the attached email")

            '    If Not myContinue Then
            '        Return False
            '        Exit Function
            '    End If

            '    
            'End If
        Else

            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket(False, True, ThisAddIn.timingFile)
            ndt.TicketNumber = DealData("NDT")

            ndt.UpdateNextDesk(Globals.ThisAddIn.WriteTicketMessage(DealData))
        End If

        DealData.Remove("AMEmailAddress")

        If Globals.ThisAddIn.sqlInterface.Add_Data(DealData) > 0 Then
            Dim rFName As String() = Split(DealData("AM"))
            Dim mygreeting As String
            mygreeting = Globals.ThisAddIn.WriteGreeting(Now(), CStr(rFName(0)))



            Try
                    Globals.ThisAddIn.MoveToFolder(TrimExtended(DealData("AM")), mail, True)
                Catch
                End Try

            Else
                DealData.Add("Result", "Failed")
        End If


        UpdateTitle("All Done!")
        Return True

    End Function


    Sub DisableButtons()

        OKButton.Enabled = False
        tCancelButton.Text = "Cancel Logging"
        CustomerName.Enabled = False
        HPIOption.Enabled = False
        HPEOption.Enabled = False
        DellOption.Enabled = False
        LenovoOption.Enabled = False
        btnMS.Enabled = False
        cIngram.Enabled = False
        cTechData.Enabled = False
        cWestcoast.Enabled = False
        DealID.Enabled = False
        chkExertis.Enabled = False


    End Sub

    Private Sub ExitEarly()

        CloseMe()
        MsgBox("Process terminated before completion")
    End Sub


    Private Sub CloseMe()

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New CloseMeCallback(AddressOf CloseMe)
            Me.Invoke(d, New Object() {})
        Else

            Me.Close()

        End If
    End Sub
    Delegate Sub CloseMeCallback()
    Private Sub UpdateTitle(ByVal [NewTitle] As String)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New UpdateTitleCallback(AddressOf UpdateTitle)
            Me.Invoke(d, New Object() {[NewTitle]})
        Else

            Me.Text = NewTitle

        End If
    End Sub
    Delegate Sub UpdateTitleCallback(ByVal [NewTitle] As String)



    Private Sub AddDeal_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        myContinue = False
    End Sub


End Class