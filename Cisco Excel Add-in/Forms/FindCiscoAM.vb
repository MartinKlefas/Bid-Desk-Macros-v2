Imports System.Diagnostics
Imports System.IO
Imports clsNextDeskTicket
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Public Class FindCiscoAM

    Private ReadOnly QuoteNum As String

    Public Sub New(Optional thisQuoteNum As String = "")
        InitializeComponent()
        'Me.Mode = thisMode
        Me.QuoteNum = thisQuoteNum
    End Sub

    Private Sub BrowserController_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UpdateLabel("Getting Everything Ready...")
        Me.TopMost = True
        BackgroundWorker1.RunWorkerAsync()

    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim browser As ChromeDriver = Nothing

        UpdateLabel(LabelMessages("Login"))
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

        browser = ndt.GiveMeChrome(True, True)

        DoLogin(browser)

        Threading.Thread.Sleep(TimeSpan.FromSeconds(3))

        UpdateLabel(LabelMessages("DL1"))
        DL_PageOne(browser, QuoteNum)

        Threading.Thread.Sleep(TimeSpan.FromSeconds(5))

        Dim tmpAM As String



        UpdateLabel(LabelMessages("AM1"))
        tmpAM = ReadAMText(browser)

        browser.Close()
        browser.Dispose()

        MsgBox("tmpAM")

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

    Function ReadAMText(Browser As ChromeDriver) As String


        Dim expandButtons = Browser.FindElementsByClassName("accordion-toggle")
        Dim buttonText As String = ""

        For Each elemnt In expandButtons
            Try

                buttonText = elemnt.Text
            Catch ex As Exception
                buttonText = ""
            End Try

            If buttonText = "Who's Involved" Then
                elemnt.Click()
                Continue For
            End If
        Next

        Threading.Thread.Sleep(TimeSpan.FromSeconds(3))

        Dim elements = Browser.FindElementsByClassName("ng-scope")
        Dim kdfid As String
        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("data-ng-if")
            Catch ex As Exception
                kdfid = ""
            End Try
            If kdfid = "whoIsInvovled.am" Then
                Return elemnt.Text
            End If
            If kdfid = "whoIsInvovled.amTeamName" Then
                Dim childelements = elemnt.FindElements(By.XPath("*"))

                For Each child In childelements
                    If child.Text = "View Members" Then
                        child.Click()
                        Dim tableCandidates = Browser.FindElementsByClassName("table-responsive")

                        For Each table In tableCandidates
                            If table.Text.Contains("Job Category") Then
                                Return table.Text
                            End If
                        Next
                    End If

                Next
            End If
        Next

        Return ""

    End Function

    Friend Sub DL_PageOne(Browser As ChromeDriver, SearchFor As String)
        Dim elements = Browser.FindElementsByClassName("commonGlobalSearch")
        Dim kdfid As String
        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("kdfpage")
            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "BUILDANDPRICE"
                        elemnt.SendKeys(SearchFor)
                End Select
            Catch
            End Try
        Next

        elements = Browser.FindElementsByClassName("commonSearchButton")

        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("kdfpage")
            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "searchButton"
                        elemnt.Click()
                End Select
            Catch
            End Try
        Next
        Threading.Thread.Sleep(TimeSpan.FromSeconds(1))
        Browser.FindElementByLinkText(SearchFor).Click()

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