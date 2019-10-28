Imports System.Diagnostics
Imports System.IO
Imports clsNextDeskTicket
Imports OpenQA.Selenium.Chrome
Partial Class BrowserController
    Private Sub DL_PageOne(Browser As ChromeDriver, SearchFor As String)
findsearchbox:
        Dim elements = Browser.FindElementsByClassName("commonGlobalSearch")
        Dim kdfid As String


        If elements.Count < 1 Then
            Threading.Thread.Sleep(TimeSpan.FromSeconds(2))
            GoTo findsearchbox
        End If

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
        Threading.Thread.Sleep(TimeSpan.FromSeconds(2))
gotodealpage:
        Try
            Browser.FindElementByLinkText(SearchFor).Click()
        Catch
            Threading.Thread.Sleep(TimeSpan.FromSeconds(2))
            GoTo gotodealpage
        End Try

    End Sub
    Private Sub DL_pageTwo(Browser As ChromeDriver)

findExportLink:
        Threading.Thread.Sleep(TimeSpan.FromSeconds(1))


        Dim elements = Browser.FindElementsByPartialLinkText("Export")

        If elements.Count < 1 Then GoTo findExportLink

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

        Dim myContinue As Boolean = False

findFormatSelector:

        Threading.Thread.Sleep(TimeSpan.FromSeconds(1))

        elements = Browser.FindElementsByTagName("Select")

        If elements.Count < 1 Then GoTo findFormatSelector


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
                        myContinue = True
                        Exit For
                End Select
            Catch

            End Try
        Next

        If Not myContinue Then GoTo findFormatSelector

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
                        Exit For
                End Select
            Catch
            End Try

        Next


    End Sub

    Public Function IsApproved(Browser As ChromeDriver) As Boolean

waitforquoteload:
        Threading.Thread.Sleep(TimeSpan.FromSeconds(2))
        Dim elements = Browser.FindElementsByClassName("text-danger")

        If elements.Count < 1 Then GoTo waitforquoteload

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

            'UpdateLabel(LabelMessages("DL2"))
            DL_pageTwo(Browser)

            'UpdateLabel(LabelMessages("DL3"))
            Dim newfile As String = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()

            Do While newfile = oldfile OrElse Strings.Right(newfile, 10) = "crdownload" OrElse Strings.Right(newfile, 3) = "tmp"
                Threading.Thread.Sleep(100)
                newfile = Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()
            Loop
            Threading.Thread.Sleep(TimeSpan.FromSeconds(2))
            Return Directory.GetFiles(Downloads).OrderByDescending(Function(f) New FileInfo(f).LastWriteTime).First()
        Else
            Return "Not Approved"

        End If
    End Function

End Class


