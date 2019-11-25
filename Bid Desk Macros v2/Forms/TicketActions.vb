Public Class TicketActions
    Private ReadOnly Action As String
    Private Ticket As String
    Private ReadOnly Autonomy As Boolean
    Private ReadOnly Comment As String

    Public Sub New(Action As String, Optional Ticket As String = "", Optional Comment As String = "", Optional CompleteAutonomy As Boolean = False)
        InitializeComponent()
        Me.Action = Action
        Me.Ticket = Ticket
        Me.NDTNum.Text = Ticket
        Me.Autonomy = CompleteAutonomy
        Me.Comment = Comment

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles CnclButton.Click
        Me.Close()
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Me.OKButton.Enabled = False
        Me.CnclButton.enabled = False
        Ticket = NDTNum.Text

        If Ticket <> "" Then
            If Autonomy Then
                BackgroundWorker1_DoWork(Nothing, Nothing)
            Else
                BackgroundWorker1.RunWorkerAsync()
            End If

        End If
    End Sub

    Private Sub TicketActions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.NDTNum.Text = Ticket
        If Autonomy Then
            OKButton.PerformClick()
        End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket With {
            .TicketNumber = Ticket,
            .VisibleBrowser = False,
            .TimeOperations = True,
            .TimingOutputFile = ThisAddIn.timingFile
        }
        Select Case Action
            Case "MS-More-Info"
                ndt.UpdateNextDesk(PreSubMoreInfo)
            Case "Cisco-DR-Type"

                ndt.UpdateNextDesk(CiscoDRType)

                ndt.UpdateNextDeskAttach(Globals.ThisAddIn.FileFromResource(My.Resources.Hunting_Questionnaire_Blank, "Hunting Questionnaire Blank.docx"))
                ndt.UpdateNextDeskAttach(Globals.ThisAddIn.FileFromResource(My.Resources.Teaming_Blank, "Teaming Questionnaire Blank.docx"))

            Case "Close"
                ndt.CloseTicket(Comment)
            Case "AttachCisco"
                'in this instance, the Cisco Deal ID is stored as the ticket number, and the filename of the quote is in the comment field.
                ndt.TicketNumber = ndt.FindTicket(0, Ticket, openOnly:=False)
                If ndt.TicketNumber <> 0 Then
                    Try
                        ndt.UpdateNextDeskAttach(Comment, CiscoAttachComment)
                        If Not ndt.InBin("Pre-sales triage") AndAlso MoveBack(ndt) Then
                            ndt.Move("Pre-sales triage")
                        End If

                    Catch
                    End Try

                Else
                    MsgBox("Could not find the right ticket for this")
                End If


        End Select

        closeme()
    End Sub


    Function MoveBack(ticket As clsNextDeskTicket.ClsNextDeskTicket) As Boolean

        Dim UpdatesDict As Dictionary(Of String, String) = ticket.GetUpdates

        For Each update As String In UpdatesDict.Values
            If update.ToLower.Contains("no need to move back") Then Return False

        Next
        Return True
    End Function

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