
Imports System.Diagnostics
Imports System.IO
Imports clsNextDeskTicket
Imports OpenQA.Selenium.Chrome

Public Class LenovoBrowserController

    Private ReadOnly Mode As String
    Private ReadOnly QuoteNum As String

    Public Sub New(thisMode As String, Optional thisQuoteNum As String = "")
        InitializeComponent()
        Me.Mode = thisMode
        Me.QuoteNum = thisQuoteNum
    End Sub
    Private Sub LenovoBrowserController_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UpdateLabel("Getting Everything Ready...")
        Me.TopMost = True
        BackgroundWorker1.RunWorkerAsync()

    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim browser As ChromeDriver = Nothing


        If Mode = "Login" Or Mode = "NewDeal" Or Mode = "DownloadQuote" Then
            UpdateLabel(LabelMessages("Login"))
            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

            browser = ndt.GiveMeChrome(True, True)

            DoLogin(browser)
        End If


        Call CloseMe()
    End Sub

    Sub DoLogin(Optional WithBrowser As ChromeDriver = Nothing)

        If IsNothing(WithBrowser) Then

            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

            WithBrowser = ndt.GiveMeChrome(True, True)
        End If

        WithBrowser.Navigate.GoToUrl("https://www.lenovopartner.com/?loginfail=true")

        Try
            Threading.Thread.Sleep(TimeSpan.FromSeconds(3))
            WithBrowser.FindElementByName("username").SendKeys("martin.klefas@insight.com")
            WithBrowser.FindElementByName("password").SendKeys("07^%*zZt2M6q&M")

            WithBrowser.FindElementByName("password").SendKeys(OpenQA.Selenium.Keys.Enter)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



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