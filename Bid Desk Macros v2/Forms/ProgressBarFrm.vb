Public Class ProgressBarFrm
    Private ReadOnly lblSnippet As String
    Public taskNum As Integer

    ''' <summary>
    ''' Save the information about the tasks being run into internal variables
    ''' </summary>
    ''' <param name="NumberofTasks">The total number of tasks as an integer</param>
    ''' <param name="LabelMessage">The missing part of the phrase "Processing [...] x of y"</param>
    Public Sub New(NumberofTasks As Integer, LabelMessage As String)
        Me.ProgressBar1.Maximum = NumberofTasks
        Me.lblSnippet = LabelMessage
        Me.Label1.Text = "Processing " & lblSnippet & " 0 of " & ProgressBar1.Maximum
        InitializeComponent()
        taskNum = 0
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        While taskNum < ProgressBar1.Maximum
            Call SetProgress(taskNum)
            Threading.Thread.Sleep(TimeSpan.FromSeconds(1))
        End While
        Call CloseMe()
    End Sub
    Private Sub SetProgress(ByVal [progress] As Integer)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.Label1.InvokeRequired Then
            Dim d As New SetProgressCallback(AddressOf SetProgress)
            Me.Invoke(d, New Object() {[progress]})
        Else

            Me.Label1.Text = "Processing " & lblSnippet & " " & [progress] & " of " & ProgressBar1.Maximum
            Me.ProgressBar1.Value = [progress]

        End If
    End Sub
    Delegate Sub SetProgressCallback(ByVal [progress] As Integer)

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

End Class