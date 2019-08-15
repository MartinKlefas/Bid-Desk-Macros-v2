
Imports System.Diagnostics
Imports System.IO
Imports clsNextDeskTicket
Imports OpenQA.Selenium.Chrome
Public Class BrowserController

    Private ReadOnly Mode As String
    Private ReadOnly QuoteNum As String
    Private ReadOnly EmailMessage As Outlook.MailItem
    Public Sub New(thisMode As String, Optional thisQuoteNum As String = "", Optional Sourcemessage As Outlook.MailItem = Nothing)
        InitializeComponent()
        Me.Mode = thisMode
        Me.QuoteNum = thisQuoteNum
        Me.EmailMessage = Sourcemessage
    End Sub

    Public Sub RunCode()
        'UpdateLabel("Getting Everything Ready...")
        Me.TopMost = True

        Dim browser As ChromeDriver = Nothing


        If Mode = "Login" Or Mode = "NewDeal" Or Mode = "DownloadQuote" Then
            'UpdateLabel(LabelMessages("Login"))
            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

            browser = ndt.GiveMeChrome(True, True)

            DoLogin(browser)
        End If


        If Mode = "NewDeal" Then
            'UpdateLabel(LabelMessages("NewDeal1"))
            ND_PageOne(browser)

            'UpdateLabel(LabelMessages("NewDeal2"))
            ND_PageTwo(browser)

            ' UpdateLabel(LabelMessages("NewDeal3"))
            ND_PageThree(browser)

            ' UpdateLabel(LabelMessages("NewDeal4"))
            ND_PageFour(browser)
        End If

        If Mode = "DownloadQuote" Then



            ' UpdateLabel(LabelMessages("DL1"))
            DL_PageOne(browser, QuoteNum)

            Threading.Thread.Sleep(TimeSpan.FromSeconds(5))

            Dim result As String = GetQuote(browser)
            browser.Quit()

            If result.Equals("Not Approved", StringComparison.CurrentCultureIgnoreCase) Then
                Debug.WriteLine("This quote is not yet approved.")
                Exit Sub
            Else

                Dim ticketForm As New TicketActions("AttachCisco", QuoteNum, result, True)
                ticketForm.Show()
                EmailMessage.Delete()

            End If
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


End Class