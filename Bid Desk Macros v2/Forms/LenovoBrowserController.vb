
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

    Public Function Runcode() As Boolean
        Try
            Dim browser As ChromeDriver = Nothing


            If Mode = "Login" Or Mode = "ShowBid" Or Mode = "SendToDisti" Then
                UpdateLabel(LabelMessages("LenovoLogin"))
                Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

                browser = ndt.GiveMeChrome(True, True)

                DoLogin(browser)
            End If

            If Mode = "ShowBid" Or Mode = "SendToDisti" Then
                UpdateLabel(LabelMessages("Searching"))
                Call FindBid(browser)
            End If

            If Mode = "SendToDisti" Then
                UpdateLabel(LabelMessages("Sending"))
                Call SendToDistribution(browser)

                browser.Quit()
                browser.Dispose()

            End If


            Call CloseMe()
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Sub SendToDistribution(WithBrowser As ChromeDriver)
        Dim MainWindowID As String = WithBrowser.CurrentWindowHandle()


        WithBrowser.FindElementByName("email_quotation_to_distributor").Click()

        Threading.Thread.Sleep(TimeSpan.FromSeconds(3))

        WithBrowser.SwitchTo().Alert().Accept()

        Threading.Thread.Sleep(TimeSpan.FromSeconds(3))

        For Each handle As String In WithBrowser.WindowHandles()

            If Not handle.Equals(MainWindowID) Then
                'set controlled window to the new child.
                WithBrowser.SwitchTo().Window(handle)

                'send the quote to everyone (why the heck not)
                WithBrowser.FindElementByName("j_id0:theForm:j_id11:j_id38").Click()
                WithBrowser.FindElementByClassName("btn").Click()
            End If

        Next
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

    Sub FindBid(WithBrowser As ChromeDriver)



        WithBrowser.Navigate.GoToUrl("https://lbp.force.com/apex/XBEMyBidRequestList?bbrtype=PCD")

        WithBrowser.FindElementByName("j_id0:form:theBlock:j_id66:j_id70").SendKeys(QuoteNum)
        WithBrowser.FindElementByName("j_id0:form:theBlock:j_id66:j_id70").SendKeys(OpenQA.Selenium.Keys.Enter)

        Try
            WithBrowser.FindElementByLinkText(QuoteNum).Click()
        Catch
            MsgBox("Could not find bid")
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