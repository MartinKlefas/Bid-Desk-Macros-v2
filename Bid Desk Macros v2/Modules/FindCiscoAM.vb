Imports clsNextDeskTicket
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Partial Class BrowserController
    Function ReadAMText(Browser As ChromeDriver) As String


        Dim expandButtons = Browser.FindElementsByClassName("accordion-toggle")
        Dim buttonText As String

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


    Function WriteAMMessage(AMString As String) As String
        If AMString.Contains(vbCr) Or AMString.Contains(vbLf) Or AMString.Contains(vbCrLf) Then
            Return Replace(CiscoAMTeamMessage, "%AM%", AMString)

        Else
            AMString = Replace(AMString, ")", "@cisco.com)")
            Return Replace(CiscoAMMessage, "%AM%", AMString)
        End If
    End Function
End Class
