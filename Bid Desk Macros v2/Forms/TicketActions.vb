Public Class TicketActions
    Private ReadOnly Action As String
    Private Ticket As String
    Private ReadOnly Autonomy As Boolean
    Private ReadOnly Comment As String

    Public Sub New(Action As String, Optional Ticket As String = "", Optional Comment As String = "", Optional CompleteAutonomy As Boolean = False)
        Me.Action = Action
        Me.Ticket = Ticket
        Me.NDTNum.Text = Ticket
        Me.Autonomy = CompleteAutonomy
        Me.Comment = Comment
        InitializeComponent()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles CnclButton.Click
        Me.Close()
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Me.OKButton.Enabled = False
        Me.CnclButton.enabled = False
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
            Case "Close"
                ndt.CloseTicket(Comment)
            Case "AttachCisco"
                'in this instance, the Cisco Deal ID is stored as the ticket number, and the filename of the quote is in the comment field.
                ndt.TicketNumber = ndt.FindTicket(0, Ticket)
                ndt.UpdateNextDeskAttach(Comment, "Please See attached the vendor quote in dollars")
                ndt.Move("Pre-sales triage")

        End Select
    End Sub
End Class