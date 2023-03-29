﻿
Imports System.Diagnostics
Imports System.IO
Imports clsNextDeskTicket
Imports OpenQA.Selenium.Chrome
Public Class BrowserController

    Private ReadOnly Mode As String
    Private ReadOnly QuoteNum As String

    Public Sub New(thisMode As String, Optional thisQuoteNum As String = "")
        InitializeComponent()
        Me.Mode = thisMode
        Me.QuoteNum = thisQuoteNum
    End Sub

    Private Sub BrowserController_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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


        If Mode = "NewDeal" Then
            UpdateLabel(LabelMessages("NewDeal1"))
            ND_PageOne(browser)

            UpdateLabel(LabelMessages("NewDeal2"))
            ND_PageTwo(browser)

            UpdateLabel(LabelMessages("NewDeal3"))
            ND_PageThree(browser)

            UpdateLabel(LabelMessages("NewDeal4"))
            ND_PageFour(browser)
        End If

        If Mode = "DownloadQuote" Then



            UpdateLabel(LabelMessages("DL1"))
            DL_PageOne(browser, QuoteNum)

            Threading.Thread.Sleep(TimeSpan.FromSeconds(5))

            Dim result As String = GetQuote(browser)

            If result.Equals("Not Approved", StringComparison.CurrentCultureIgnoreCase) Then
                Debug.WriteLine("This quote is not yet approved.")
            Else
                Debug.WriteLine(result)
            End If
        End If
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