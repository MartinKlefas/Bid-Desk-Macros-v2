
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
        Dim browser As ChromeDriver = Nothing
        Try

            'UpdateLabel("Getting Everything Ready...")
            Me.TopMost = True




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

                If QuoteNum <> "0" Then

                    ' UpdateLabel(LabelMessages("DL1"))
                    DL_PageOne(browser, QuoteNum)



                    Dim result As String = GetQuote(browser)
                    browser.Quit()

                    If result.Equals("Not Approved", StringComparison.CurrentCultureIgnoreCase) Then
                        Debug.WriteLine("This quote is not yet approved.")
                        EmailMessage.Categories = "Cisco Not Approved"
                        EmailMessage.UnRead = False
                        Exit Sub
                    Else

                        Dim ticketForm As New TicketActions("AttachCisco", QuoteNum, result, True)
                        ticketForm.Show()
                        EmailMessage.Delete()

                    End If
                End If
            End If
        Catch
            browser.Quit()
        End Try

    End Sub

    Sub DoLogin(Optional WithBrowser As ChromeDriver = Nothing)

        If IsNothing(WithBrowser) Then

            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

            WithBrowser = ndt.GiveMeChrome(True, True)
        End If

        WithBrowser.Navigate.GoToUrl("https://apps.cisco.com/ccw/cpc/home")

        Try
            Threading.Thread.Sleep(TimeSpan.FromSeconds(3))
            WithBrowser.FindElementByName("pf.username").SendKeys("martinklefas")

            WithBrowser.FindElementByName("login-button").Click()

            Threading.Thread.Sleep(TimeSpan.FromSeconds(4))
startpassword:
            Try
                WithBrowser.FindElementByName("password").SendKeys("sis1898p7aPu")
            Catch
                Threading.Thread.Sleep(TimeSpan.FromSeconds(3))
                GoTo startpassword
            End Try

            WithBrowser.FindElementById("kc-login").Click()

            Catch ex As Exception
                WithBrowser.FindElementByClassName("username-input").SendKeys("martinklefas")
            WithBrowser.FindElementByClassName("password-input").SendKeys("sis1898p7aPu")

            WithBrowser.FindElementByName("login-button").Click()
        End Try

    End Sub


End Class