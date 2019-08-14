Public Class TicketActions
    Private Action As String
    Private Ticket As String
    Private Autonomy As Boolean

    Public Sub New(Action As String, Optional Ticket As String = "", Optional CompleteAutonomy As Boolean = False)
        Me.Action = Action
        Me.Ticket = Ticket
        Me.NDTNum.Text = Ticket
        Me.Autonomy = CompleteAutonomy
        InitializeComponent()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Ticket = NDTNum.Text
        If Ticket <> "" Then BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub TicketActions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.NDTNum.Text = Ticket
        If Autonomy Then
            OKButton.PerformClick()
        End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket With {
            .TicketNumber = Ticket
        }
        Select Case Action
            Case "MS-More-Info"
                ndt.UpdateNextDesk(PreSubMoreInfo)
        End Select
    End Sub
End Class