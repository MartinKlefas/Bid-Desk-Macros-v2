Imports System.Windows.Forms

Public Class AddDeal
    Private Sub CommandButton1_Click() Handles OKButton.Click
        CustomerName.Text = Trim(CustomerName.Text)
        DealID.Text = Trim(DealID.Text)
        Me.DialogResult = DialogResult.OK
        Me.Hide()
    End Sub

    Private Sub Button2_Click() Handles tCancelButton.Click

        Me.DialogResult = DialogResult.Cancel
        Me.Hide()
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
        Me.Hide()
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

    End Sub
End Class