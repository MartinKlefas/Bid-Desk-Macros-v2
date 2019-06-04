
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
        BackgroundWorker1.RunWorkerAsync()

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        If Mode = "Login" Then
            UpdateLabel(LabelMessages("Login"))
            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

            Dim browser As ChromeDriver = ndt.GiveMeChrome(True, True)

            DoLogin(browser)
        End If

        If Mode = "NewDeal" Then
            UpdateLabel(LabelMessages("Login"))
            Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

            Dim browser As ChromeDriver = ndt.GiveMeChrome(True, True)

            DoLogin(browser)

            UpdateLabel(LabelMessages("NewDeal1"))

            ND_PageOne(browser)




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

    Private Sub ND_PageOne(Browser As ChromeDriver)
        Browser.Navigate.GoToUrl("https://apps.cisco.com/ICW/PDR/deal#/createdeal")


        Dim elements = Browser.FindElementsByClassName("form-control")
        Dim kdfid As String
        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("kdfid")
            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "dealName"
                        elemnt.SendKeys("Test Deal / Martin Klefas / IMA")
                    Case "customerContactFirstName"
                        elemnt.SendKeys("customer firstname")
                    Case "customerContactLastName"
                        elemnt.SendKeys("customer lastname")
                    Case "customerContactJobTitle"
                        elemnt.SendKeys("Job Title")
                    Case "customerContactPhoneNumber"
                        elemnt.SendKeys("Phone Number")
                    Case "customerContactEmailAddress"
                        elemnt.SendKeys("Email Address")
                    Case "customerContactCompanyUrl"
                        elemnt.SendKeys("WebSite")
                    Case "selectedCAMId"
                        elemnt.SendKeys("mc")
                    Case ""
                End Select

            Catch
            End Try

        Next

        Browser.FindElementByPartialLinkText("Partner ").Click()
        elements = Browser.FindElementsByClassName("check")
        For Each elemnt In elements
            For Each tChild In elemnt.FindElements(OpenQA.Selenium.By.TagName("label"))
                If tChild.GetAttribute("for") = "bUseContactDetails" Then
                    tChild.Click()
                End If
            Next

        Next
        Browser.FindElementByName("camId").SendKeys("MC")




        Browser.FindElementByPartialLinkText("Faster Search").Click()

        elements = Browser.FindElementsByClassName("form-control")

        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("kdfid")
            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "addEC"
                        elemnt.SendKeys("test customer")
                        Exit For
                End Select
            Catch
            End Try

        Next

        MsgBox("Please select your customer from the list, and click OK once you've clicked ""Use Selected Address""")

        elements = Browser.FindElementsByClassName("btn")

        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("kdfid")
            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "createDealBtn"
                        elemnt.Click()
                        Exit For
                    Case Else
                End Select
            Catch
            End Try

        Next
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