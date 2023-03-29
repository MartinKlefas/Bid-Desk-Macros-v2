Imports System.Diagnostics
Imports System.IO
Imports clsNextDeskTicket
Imports OpenQA.Selenium.Chrome
Partial Class BrowserController
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
    Private Sub DL_pageTwo(Browser As ChromeDriver)
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

        Threading.Thread.Sleep(TimeSpan.FromSeconds(2))

        elements = Browser.FindElementsByTagName("Select")
        Dim kdfid As String
        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("kdfid")
            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "fileType"
                        elemnt.SendKeys("C")
                        Exit For
                End Select
            Catch

            End Try
        Next

        elements = Browser.FindElementsByTagName("label")
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
                        Exit For
                End Select
            Catch
            End Try

        Next

        elements = Browser.FindElementsByClassName("btn")
        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("data-ng-click")

            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "exportQuote();"
                        elemnt.Click()

                End Select
            Catch
            End Try

        Next


    End Sub

    Public Function IsApproved(Browser As ChromeDriver) As Boolean
        Dim elements = Browser.FindElementsByClassName("text-danger")
        Dim kdfid As String
        For Each elemnt In elements
            Try
                kdfid = elemnt.GetAttribute("data-ng-class")

            Catch ex As Exception
                kdfid = ""
            End Try
            Try
                Select Case kdfid
                    Case "setStatusClass(dealHeader.quoteStatus)"
                        Return elemnt.Text.Equals("Approved", StringComparison.CurrentCultureIgnoreCase)
                End Select
            Catch
            End Try
        Next
        Return False
    End Function

    Function GetQuote(Browser As ChromeDriver) As String
        Dim oldfile As String
        Dim Downloads As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%\Downloads")


        If IsApproved(Browser) Then
            oldfile = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()

            UpdateLabel(LabelMessages("DL2"))
            DL_pageTwo(Browser)

            UpdateLabel(LabelMessages("DL3"))

            Do While Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First() = oldfile
                Threading.Thread.Sleep(100)
            Loop

            Return Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()
        Else
            Return "Not Approved"

        End If
    End Function

End Class


