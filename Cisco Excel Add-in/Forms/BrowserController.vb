Public Class BrowserController
    Private Mode As String

    Public Sub New(thisMode As String)
        InitializeComponent()
        Me.Mode = thisMode
    End Sub

    Private Sub BrowserController_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.LblStatus.Text = LabelMessages(Mode)
    End Sub
End Class