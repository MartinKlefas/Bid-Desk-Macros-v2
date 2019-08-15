Imports System.Diagnostics
Imports clsNextDeskTicket
Imports OpenQA.Selenium.Chrome
Partial Class BrowserController
    Private Sub DL_PageOne(Browser As ChromeDriver, SearchFor As String)
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
    Private Sub DL_pageTwo(Browser As ChromeDriver, SearchFor As String)
        Dim elements = Browser.FindElementsByPartialLinkText("Export")

        For Each elemnt In elements
            If elemnt.Text = "Export" Then
                elemnt.Click()
                Exit For

            End If
        Next
        elements = Browser.FindElementsByPartialLinkText("Export")
        For Each elemnt In elements
            If elemnt.Text = "Export Quote" Then
                elemnt.Click()
                Exit For

            End If
        Next

        elements = Browser.FindElementsByClassName("commonGlobalSearch")
        Dim kdfid As String
        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("kdfpage")
            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "quote"
                        elemnt.SendKeys("C")
                End Select
            Catch

            End Try
        Next

        elements = Browser.FindElementByTagName("label")
        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("for")

            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "c-0"
                        elemnt.Click()

                End Select
            Catch
            End Try

        Next


    End Sub



End Class


