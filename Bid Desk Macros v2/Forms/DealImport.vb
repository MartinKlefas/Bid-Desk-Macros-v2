Imports System.Windows.Forms
Imports String_Extensions

Public Class ImportDeal
    Public Sub New(Optional AMEmailAddress As String = "", Optional ndt As String = "")

        ' This call is required by the designer.
        InitializeComponent()


        ' Add any initialization after the InitializeComponent() call.
        Me.AMMail.Text = AMEmailAddress
        Me.NDTNumber.Text = ndt
    End Sub

    Private Sub CommandButton1_Click() Handles OKButton.Click
        CustomerName.Text = TrimExtended(CustomerName.Text)
        DealID.Text = TrimExtended(DealID.Text)



        Dim DealData As Dictionary(Of String, String)

        Dim AmName, Vendor, bIngram, bWestCoast, bTechData As String

        Vendor = "Unknown"
        If HPIOption.Checked Then Vendor = "HPI"
        If HPEOption.Checked Then Vendor = "HPE"
        If DellOption.Checked Then Vendor = "Dell"
        If MSOption.Checked Then Vendor = "Microsoft"
        If LenovoOption.Checked Then Vendor = "Lenovo"

        Dim AmExUser = Globals.ThisAddIn.MyResolveName(AMMail.Text)

        AmName = AmExUser.FirstName & " " & AmExUser.LastName

        If cIngram.Checked Then
            bIngram = 1
        Else
            bIngram = 0
        End If
        If cWestcoast.Checked Then
            bWestCoast = 1
        Else
            bWestCoast = 0
        End If
        If cTechData.Checked Then
            bTechData = 1
        Else
            bTechData = 0
        End If

        DealData = New Dictionary(Of String, String) From {
                {"AM", AmName},
                {"Customer", CustomerName.Text},
                {"Vendor", Vendor},
                {"DealID", DealID.Text},
                {"Ingram", bIngram},
                {"Techdata", bTechData},
                {"Westcoast", bWestCoast},
                {"NDT", NDTNumber.Text},
                {"CC", ccList.Text},
                {"Status", "Submitted to Vendor"},
                {"StatusDate", DateTime.Now().ToString("yyyyMMdd HH:mm:ss")},
                {"Date", DateTime.Now().ToString("yyyyMMdd HH:mm:ss")}
            }

        Dim tmpCountAdded As Integer = Globals.ThisAddIn.sqlInterface.Add_Data(DealData)
        If tmpCountAdded <> 1 Then
            MsgBox(tmpCountAdded & " lines added")
        End If


        If Me.ChkTicket.Checked AndAlso NDTNumber.Text <> "" AndAlso IsNumeric(NDTNumber.Text) Then

            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket(False, True, ThisAddIn.timingFile) With {
                .TicketNumber = NDTNumber.Text
            }

            ndt.UpdateNextDesk(Globals.ThisAddIn.WriteTicketMessage(DealData))
        End If

        Me.Close()
    End Sub

    Private Sub Button2_Click() Handles tCancelButton.Click

        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub UserForm_Activate() Handles Me.Activated


        Dim strClip As String


        strClip = My.Computer.Clipboard.GetText
        If Me.DealID.Text = "" Then Me.DealID.Text = FindDealID(strClip)
        If Me.CustomerName.Text = "" Then Me.CustomerName.Text = FindCustomer(strClip)

        Select Case FindVendor(strClip)
            Case "HPI"
                Call CheckOnly(HPIOption)
            Case "HPE"
                Call CheckOnly(HPEOption)
            Case "Dell"
                Call CheckOnly(DellOption)
            Case "Microsoft"
                Call CheckOnly(MSOption)
            Case "Lenovo"
                Call CheckOnly(LenovoOption)

        End Select
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

    Private Sub AMMail_MouseDown1(sender As Object, e As MouseEventArgs) Handles AMMail.MouseDown
        If e.Button = MouseButtons.Right Then

            AMMail.Text = TrimExtended(My.Computer.Clipboard.GetText)
        End If
    End Sub
    Private Sub CustomerName_MouseDown1(sender As Object, e As MouseEventArgs) Handles CustomerName.MouseDown
        If e.Button = MouseButtons.Right Then

            CustomerName.Text = TrimExtended(My.Computer.Clipboard.GetText)
        End If
    End Sub
    Private Sub DealID_MouseDown1(sender As Object, e As MouseEventArgs) Handles DealID.MouseDown
        If e.Button = MouseButtons.Right Then

            DealID.Text = TrimExtended(My.Computer.Clipboard.GetText)
        End If
    End Sub
    Private Sub NDTNumber_MouseDown1(sender As Object, e As MouseEventArgs) Handles NDTNumber.MouseDown
        If e.Button = MouseButtons.Right Then

            NDTNumber.Text = TrimExtended(My.Computer.Clipboard.GetText)
        End If
    End Sub


End Class