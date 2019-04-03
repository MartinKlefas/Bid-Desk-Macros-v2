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
            Return sqlInterface.SelectData("CC", "DealID = '" & dealID & "'")
        Catch
            ShoutError("there was an error getting the CC details", SuppressWarnings)

            Return ""
        End Try
    End Function

    Public Function GetNDTbyDeal(DealID As String, Optional AllTickets As Boolean = False, Optional SuppressWarnings As Boolean = True) As String
        Try
            Dim allData As String = sqlInterface.SelectData("NDT", "DealID = '" & DealID & "'")

            If AllTickets Or Not allData.Contains(";") Then
                Return allData
            Else
                Return Split(allData, ";").Last
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

        Return sqlInterface.Update_Data("NDT = " & newNDT, "DealID = '" & DealID & "'")
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

    Public Function GetFact(DealID As String) As String
        Dim i As Integer = 0, number As Integer

        While Not RegularExpressions.Regex.IsMatch(Mid(DealID, i), "^[0-9]+$")
            i = i + 1

        End While

        number = CInt(Mid(DealID, i))

        Dim myURL As String, a As String
        myURL = "http://numbersapi.com/" & number & "/trivia?fragment"

        Dim WinHttpReq As Object, shortNumber As Long

        WinHttpReq = CreateObject("Microsoft.XMLHTTP")
        WinHttpReq.Open("GET", myURL, False, "username", "password")
        WinHttpReq.Send

        'myURL = WinHttpReq.responseBody
        a = ""
        If WinHttpReq.status = 200 Then
            a = System.Text.Encoding.Unicode.GetString(WinHttpReq.responseBody)
        End If

        If a = "" Then
            GetFact = "the interesting number facts service is currently broken!"
        ElseIf a = "a boring number" Or a = "an uninteresting number" Or a = "an unremarkable number" Or a = "a number for which we're missing a fact (submit one to numbersapi at google mail!)" Or a = "a boring number" Then
            If Len(CStr(number)) > 2 Then
                shortNumber = CLng(Right(CStr(number), 2))
                GetFact = "Sadly " & number & " is unremarkable. " & GetFact(shortNumber)
            Else
                GetFact = "Unfortunately " & number & " is too!"
            End If
        Else

            GetFact = "Interestingly " & number & " is " & a

        End If

    End Function


    Public Function IsWestcoast(DealID As Integer) As Boolean
        Dim tmpResult As String, intresult As Integer
        tmpResult = sqlInterface.SelectData("Westcoast", "DealID = '" & DealID & "'")

        Try
            intresult = CInt(tmpResult)
        Catch
            Return False
        End Try

        Return intresult = 1
    End Function
    Public Function IsTechData(DealID As Integer) As Boolean
        Dim tmpResult As String, intresult As Integer
        tmpResult = sqlInterface.SelectData("Techdata", "DealID = '" & DealID & "'")

        Try
            intresult = CInt(tmpResult)
        Catch
            Return False
        End Try

        Return intresult = 1
    End Function
    Public Function IsIngram(DealID As Integer) As Boolean
        Dim tmpResult As String, intresult As Integer
        tmpResult = sqlInterface.SelectData("Ingram", "DealID = '" & DealID & "'")

        Try
            intresult = CInt(tmpResult)
        Catch
            Return False
        End Try

        Return intresult = 1
    End Function
End Class
