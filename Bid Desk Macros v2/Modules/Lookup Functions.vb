Imports System.IO
Imports System.Net
Imports OpenQA.Selenium.Chrome
Imports String_Extensions.StringExtensions

Partial Class ThisAddIn
    Public Function GetAMbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String
        Return GetFolderbyDeal(dealID, SuppressWarnings)

    End Function

    Public Function GetFolderbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String

        Try
            Return sqlInterface.SelectData("AM", "DealID = '" & dealID & "'")
        Catch
            ShoutError("there was an error there was an error looking up the AM", SuppressWarnings)

            Return ""
        End Try
    End Function

    Public Function QuotesReceived(dealID As String,
                                    Optional SuppressWarnings As Boolean = True) As Integer
        Try
            Return sqlInterface.SelectData("QuotesReceived", "DealID = '" & dealID & "'")

        Catch
            ShoutError("there was an error there was an error looking up the Number of quotes received", SuppressWarnings)

            Return 0
        End Try

    End Function


    Public Function GetCustomerbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String

        Try
            Return sqlInterface.SelectData("Customer", "DealID = '" & dealID & "'")
        Catch

            ShoutError("there was an error there was an error getting the customer name", SuppressWarnings)

            Return ""
        End Try
    End Function
    Public Function GetCCbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String

        Try
            Dim tmp As String = sqlInterface.SelectData("CC", "DealID = '" & dealID & "'")
            If TrimExtended(tmp) = "0" Then tmp = ""
            Return tmp
        Catch
            ShoutError("there was an error getting the CC details", SuppressWarnings)

            Return ""
        End Try
    End Function

    Public Function NoOpenTickets(DealID As String) As Boolean
        Dim allTickets As String() = GetNDTbyDeal(DealID, True).Split(";")
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

        Dim wd As ChromeDriver = ndt.GiveMeChrome(False)

        For Each ticket As String In allTickets
            ndt.TicketNumber = ticket
            If Not ndt.IsClosed(wd) Then
                wd.Close()
                Return False
            End If
        Next

        wd.Quit()
        Return True

    End Function

    Public Function GetOpenTicket(DealID As String) As String
        Dim allTickets As String() = GetNDTbyDeal(DealID, True).Split(";")
        Dim ndt As New clsNextDeskTicket.ClsNextDeskTicket

        Dim wd As ChromeDriver = ndt.GiveMeChrome(False)

        For Each ticket As String In allTickets
            ndt.TicketNumber = ticket
            If Not ndt.IsClosed(wd) Then
                wd.Close()
                Return ticket
            End If
        Next

        wd.Quit()
        Return ""

    End Function

    Public Function GetNDTbyDeal(DealID As String, Optional AllTickets As Boolean = False, Optional SuppressWarnings As Boolean = True) As String
        Try
            Dim allData As String = sqlInterface.SelectData("NDT", "DealID = '" & DealID & "'")

            If AllTickets Or Not allData.Contains(";") Then
                Return allData.TrimExtended
            Else
                Return Split(allData, ";").Last.TrimExtended
            End If

        Catch
            ShoutError("there was an error getting the ticket number", SuppressWarnings)

            Return ""
        End Try

    End Function

    Public Function AddNewTicketToDeal(DealID As String, TicketNumber As Integer) As Integer
        Dim oldNDT As String, newNDT As String

        oldNDT = GetNDTbyDeal(DealID)
        newNDT = oldNDT & ";" & TicketNumber

        Return sqlInterface.Update_Data("NDT = '" & newNDT & "'", "DealID = '" & DealID & "'")
    End Function

    Public Function AddOPG(DealID As String, OPG As String) As Integer
        Return sqlInterface.Update_Data("OPGID = '" & OPG & "'", "DealID = '" & DealID & "'")
    End Function

    Public Function AddQuoteReceived(DealID As String) As Boolean
        Dim currentRecieved As Integer
        currentRecieved = QuotesReceived(DealID)
        Return sqlInterface.Update_Data("QuotesReceived = " & currentRecieved + 1, "DealID = '" & DealID & "'") = 1
    End Function

    Public Function ChangeAM(OldAM As String, NewAM As String) As Integer
        Return sqlInterface.Update_Data("AM = '" & OldAM & "'", "AM = '" & NewAM & "'")
    End Function

    Public Function GetSubmitTime(DealID As String, Optional SuppressWarnings As Boolean = True) As Date
        Try
            Return sqlInterface.SelectData_Date("Date", "DealID = '" & DealID & "'")
        Catch

            ShoutError("there was an error getting the deal submission time", SuppressWarnings)
            Return ""
        End Try
    End Function
    Public Function GetVendor(DealID As String, Optional SuppressWarnings As Boolean = True) As String
        Try
            Return sqlInterface.SelectData("Vendor", "DealID = '" & DealID & "'")
        Catch
            ShoutError("there was an error getting the vendor", SuppressWarnings)

            Return ""
        End Try
    End Function

    Public Function GetFact(ByVal DealID As String) As String
        Dim i As Integer = 1, number As Integer
        DealID = TrimExtended(DealID)
        If DealID.ToLower.Contains("cas") Then
            DealID = Mid(DealID, 4, 6)
        End If
        While Not RegularExpressions.Regex.IsMatch(Mid(DealID, i), "^[0-9]+$")
            i += 1
            If DealID = "" Or i >= Len(DealID) - 1 Then Exit While
        End While

        Try
            number = CInt(Mid(DealID, i))

            Dim myURL, a As String, shortnumber As Long
            myURL = "http://numbersapi.com/" & number & "/trivia?fragment"

            a = LoadSiteContents(myURL)

            If a = "" Then
                GetFact = "the interesting number facts service is currently broken!"
            ElseIf a = "a boring number" Or a = "an uninteresting number" Or a = "an unremarkable number" Or a = "a number for which we're missing a fact (submit one to numbersapi at google mail!)" Then
                If number > 99 Then
                    shortnumber = CLng(Right(CStr(number), 2)) ' get the last 2 digits of the number
                    GetFact = "Sadly " & number & " is unremarkable. " & GetFact(shortnumber)
                Else
                    GetFact = "Unfortunately " & number & " is too!"
                End If
            Else

                GetFact = "Interestingly " & number & " is " & a

            End If
        Catch
            GetFact = "the interesting number facts service is currently broken!"
        End Try

    End Function

    ''' <summary>
    ''' method for retrieving the data from the provided URL
    ''' </summary>
    ''' <param name="url">URL we're scraping</param>
    ''' <returns></returns>
    Private Function LoadSiteContents(ByVal url As String) As String
        Try
            'create a new WebRequest object
            Dim request As WebRequest = WebRequest.Create(url)

            'create StreamReader to hold the returned request
            Dim stream As New StreamReader(request.GetResponse().GetResponseStream())

            Dim text As String = stream.ReadToEnd()
            Return text
        Catch ex As Exception
            'put your error handling here
            Return String.Empty
        End Try
    End Function

    Public Function IsWestcoast(DealID As String) As Boolean
        Dim tmpResult As String, intresult As Integer
        tmpResult = sqlInterface.SelectData("Westcoast", "DealID = '" & DealID & "'")

        Try
            intresult = CInt(tmpResult)
        Catch
            Return False
        End Try

        Return intresult = 1
    End Function
    Public Function IsTechData(DealID As String) As Boolean
        Dim tmpResult As String, intresult As Integer
        tmpResult = sqlInterface.SelectData("Techdata", "DealID = '" & DealID & "'")

        Try
            intresult = CInt(tmpResult)
        Catch
            Return False
        End Try

        Return intresult = 1
    End Function
    Public Function IsIngram(DealID As String) As Boolean
        Dim tmpResult As String, intresult As Integer
        tmpResult = sqlInterface.SelectData("Ingram", "DealID = '" & DealID & "'")

        Try
            intresult = CInt(tmpResult)
        Catch
            Return False
        End Try

        Return intresult = 1
    End Function

    Public Function DealExists(ByRef dealID As String) As Boolean
        Dim dealID_exists As Boolean = sqlInterface.ValueExists(dealID, "DealID")

        If dealID_exists Then
            Return True
        Else
            If sqlInterface.ValueExists(dealID, "OPGID") Then
                dealID = GetDealfromOPG(dealID)
                Return True
            Else
                Return False
            End If
        End If

    End Function
    Public Function GetDealfromOPG(ByRef dealID As String) As String
        Return sqlInterface.SelectData("DealID", "OPGID = '" & dealID & "'")
    End Function

End Class
