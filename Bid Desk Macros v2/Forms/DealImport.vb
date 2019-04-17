Imports System.Windows.Forms

Public Class ImportDeal
    Public Sub New(Optional AMEmailAddress As String = "")

        ' This call is required by the designer.
        InitializeComponent()
        Me.AMMail.Text = AMEmailAddress

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub CommandButton1_Click() Handles OKButton.Click
        CustomerName.Text = Trim(CustomerName.Text)
        DealID.Text = Trim(DealID.Text)
        Me.DialogResult = DialogResult.OK
        'do the actual work!
    End Sub

    Private Sub Button2_Click() Handles tCancelButton.Click

        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub UserForm_Activate() Handles Me.Activated


        Dim strClip As String


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

    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
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

            AMMail.Text = Trim(My.Computer.Clipboard.GetText)
        End If
    End Sub
    Private Sub CustomerName_MouseDown1(sender As Object, e As MouseEventArgs) Handles CustomerName.MouseDown
        If e.Button = MouseButtons.Right Then

            CustomerName.Text = Trim(My.Computer.Clipboard.GetText)
        End If
    End Sub
    Private Sub DealID_MouseDown1(sender As Object, e As MouseEventArgs) Handles DealID.MouseDown
        If e.Button = MouseButtons.Right Then

            DealID.Text = Trim(My.Computer.Clipboard.GetText)
        End If
    End Sub
    Private Sub NDTNumber_MouseDown1(sender As Object, e As MouseEventArgs) Handles NDTNumber.MouseDown
        If e.Button = MouseButtons.Right Then

            NDTNumber.Text = Trim(My.Computer.Clipboard.GetText)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        CheckOnly(HPIOption)
    End Sub
End Class