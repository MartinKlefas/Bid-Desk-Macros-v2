Imports System.Windows.Forms

Public Class newDeal
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


        Dim strClip As String, strArry As Object


        strClip = My.Computer.Clipboard.GetText

        If InStr(1, strClip, "SQ-") > 0 Then
            DealID.Text = Mid(strClip, InStr(1, strClip, "SQ-"), 10)
            DellOption.Checked = False
            HPIOption.Checked = True
        End If

        If InStr(1, strClip, "Full Legal Name") > 0 Then
            strArry = Split(Mid(strClip, InStr(1, strClip, "Full Legal Name")), vbCrLf)
            CustomerName.Text = StrConv(strArry(2), vbProperCase)
        End If

        If InStr(1, strClip, "End User Account Name") > 0 Then
            strArry = Split(Mid(strClip, InStr(1, strClip, "End User Account Name")), vbTab)
            CustomerName.Text = strArry(1)
        End If

        If InStr(1, strClip, "Deal ID") > 0 Then
            strArry = Split(Mid(strClip, InStr(1, strClip, "Deal ID")), vbTab)
            DealID.Text = strArry(1)
        End If

        If InStr(1, strClip, "HP Opportunity ID") > 0 Then
            strArry = Split(Mid(strClip, InStr(1, strClip, "HP Opportunity ID")), vbCrLf)
            DealID.Text = strArry(5)
            CustomerName.Text = strArry(13)
            DellOption.Checked = False
            HPIOption.Checked = True
        End If

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