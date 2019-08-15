Imports clsNextDeskTicket
Imports OpenQA.Selenium.Chrome
Partial Class BrowserController
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
                        elemnt.SendKeys("test@insight.com")
                    Case "customerContactCompanyUrl"
                        elemnt.SendKeys("WebSite")
                    Case "selectedCAMId"
                        elemnt.SendKeys("mc")
                    Case ""
                End Select

            Catch
            End Try

        Next


        Browser.FindElementByName("camId").SendKeys("MC")




        Browser.FindElementByPartialLinkText("Faster Search").Click()

WaitForEndUserBox:
        Try
            Threading.Thread.Sleep(20)
            Dim testElem = Browser.FindElementByClassName("end-customer-panel")
        Catch
            GoTo WaitForEndUserBox
        End Try

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
                        elemnt.SendKeys("test customer") ' replace with customer name
                        Exit For
                End Select
            Catch
            End Try

        Next

        MsgBox("Please select your customer from the list, and click OK once you've clicked ""Use Selected Address""")

        Browser.FindElementByPartialLinkText("Partner ").Click()
        elements = Browser.FindElementsByClassName("check")
        For Each elemnt In elements
            For Each tChild In elemnt.FindElements(OpenQA.Selenium.By.TagName("label"))
                If tChild.GetAttribute("for") = "bUseContactDetails" Then
                    tChild.Click()
                End If
            Next

        Next


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

    Private Sub ND_PageTwo(Browser As ChromeDriver)

WaitForSubmit:
        Try
            Threading.Thread.Sleep(20)

            Browser.FindElementByClassName("icon-ws-datepicker").Click()



        Catch
            GoTo WaitForSubmit
        End Try

        MsgBox("Please select a close date and click ok")

        Browser.FindElementByName("deal_category").SendKeys("Oth") '.FindElement(OpenQA.Selenium.By.TagName("select"))
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
                    Case "expectedHSListAmount"
                        elemnt.SendKeys(OpenQA.Selenium.Keys.Control + "a")
                        elemnt.SendKeys("50000")
                        Exit For
                End Select

            Catch
            End Try

        Next
                          
        elements = Browser.FindElementsByClassName("btn")

        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("kdfid")
            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "saveAndContinueGDRAction"
                        elemnt.Click()
                        Exit For
                    Case Else
                End Select
            Catch
            End Try

        Next

    End Sub

    Private Sub ND_PageThree(Browser As ChromeDriver)
WaitForSubmit:

        Try
            Threading.Thread.Sleep(20)
            Dim elements = Browser.FindElementsByClassName("incentive-child")



            If elements.Count < 1 Then GoTo waitforsubmit

            Dim kdfid As String
            For Each elemnt In elements

                For Each tChild In elemnt.FindElements(OpenQA.Selenium.By.TagName("label"))
                    If tChild.GetAttribute("for") = "check-BR - PSPP - 160729 - 31128" Then
                        Try
                            tChild.Click()
                            Exit For
                        Catch
                        End Try
                    End If
                Next
            Next


            elements = Browser.FindElementsByClassName("btn")

            For Each elemnt In elements
                Try
                    kdfid = elemnt.GetAttribute("kdfid")
                Catch ex As Exception
                    kdfid = ""
                End Try
                Try
                    Select Case kdfid
                        Case "saveAndContinueGDRAction"
                            elemnt.Click()
                            Exit For
                        Case Else
                    End Select
                Catch
                End Try

            Next
        Catch
        End Try

    End Sub

    Private Sub ND_PageFour(Browser As ChromeDriver)

WaitForSubmit:
        Try
            Threading.Thread.Sleep(20)
            Browser.FindElementByName("selectedEnrollment").Click()



        Catch
            GoTo WaitForSubmit
        End Try


        Browser.FindElementByName("selectedEnrollment").SendKeys(OpenQA.Selenium.Keys.Down)
        Browser.FindElementByName("selectedEnrollment").SendKeys(OpenQA.Selenium.Keys.Down)

    End Sub



End Class
