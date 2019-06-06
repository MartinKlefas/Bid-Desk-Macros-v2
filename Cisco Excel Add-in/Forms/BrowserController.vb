
Imports clsNextDeskTicket
Imports OpenQA.Selenium.Chrome
Public Class BrowserController

    Private ReadOnly Mode As String

    Public Sub New(thisMode As String)
        InitializeComponent()
        Me.Mode = thisMode
    End Sub

    Private Sub BrowserController_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UpdateLabel("Getting Everything Ready...")
        Me.TopMost = True
        BackgroundWorker1.RunWorkerAsync()

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim browser As ChromeDriver = Nothing

        If Mode = "Login" Or Mode = "NewDeal" Then
            UpdateLabel(LabelMessages("Login"))
            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

            browser = ndt.GiveMeChrome(True, True)

            DoLogin(browser)
        End If


        If Mode = "NewDeal" Then
            UpdateLabel(LabelMessages("NewDeal1"))

            ND_PageOne(browser)

            ND_PageTwo(browser)

            ND_PageThree(browser)

            ND_PageFour(browser)
        End If
    End Sub

    Sub DoLogin(Optional WithBrowser As ChromeDriver = Nothing)

        If IsNothing(WithBrowser) Then

            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

            WithBrowser = ndt.GiveMeChrome(True, True)
        End If

        WithBrowser.Navigate.GoToUrl("https://apps.cisco.com/ccw/cpc/home")

        WithBrowser.FindElementByClassName("username-input").SendKeys("martinklefas")
        WithBrowser.FindElementByClassName("password-input").SendKeys("Mpintranp23")

        WithBrowser.FindElementByName("login-button").Click()
    End Sub




    Private Sub CloseMe()

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.LblStatus.InvokeRequired Then
            Dim d As New CloseMeCallback(AddressOf CloseMe)
            Me.Invoke(d, New Object() {})
        Else

            Me.Close()

        End If
    End Sub
    Delegate Sub CloseMeCallback()
    Private Sub UpdateLabel(ByVal [NewText] As String)

        ' InvokeRequired required compares the thread ID of the'
        ' calling thread to the thread ID of the creating thread.'
        ' If these threads are different, it returns true.'
        If Me.LblStatus.InvokeRequired Then
            Dim d As New UpdateLabelCallback(AddressOf UpdateLabel)
            Me.Invoke(d, New Object() {[NewText]})
        Else

            Me.LblStatus.Text = NewText

        End If
    End Sub
    Delegate Sub UpdateLabelCallback(ByVal [NewText] As String)
End Class