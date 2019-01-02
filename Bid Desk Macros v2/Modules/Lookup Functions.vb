Partial Class ThisAddIn
    Public Function GetFolderbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String

        Try
            Return sqlInterface.SelectData("AM", "DealID = " & dealID)
        Catch
            ShoutError("there was an error there was an error looking up the AM", SuppressWarnings)

            Return ""
        End Try
    End Function

    Public Function GetCustomerbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String

        Try
            Return sqlInterface.SelectData("Customer", "DealID = " & dealID)
        Catch

            ShoutError("there was an error there was an error getting the customer name", SuppressWarnings)

            Return ""
        End Try
    End Function
    Public Function GetCCbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String

        Try
            Return sqlInterface.SelectData("CC", "DealID = " & dealID)
        Catch
            ShoutError("there was an error getting the CC details", SuppressWarnings)

            Return ""
        End Try
    End Function

    Public Function GetNDTbyDeal(DealID As String, Optional AllTickets As Boolean = False, Optional SuppressWarnings As Boolean = True) As String
        Try
            Dim allData As String = sqlInterface.SelectData("NDT", "DealID = " & DealID)

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

        Return sqlInterface.Update_Data("NDT = " & newNDT, "DealID = " & DealID)
    End Function

    Public Function GetSubmitTime(DealID As String, Optional SuppressWarnings As Boolean = True) As Date
        Try
            Return sqlInterface.SelectData_Date("Date", "DealID = " & DealID)
        Catch

            ShoutError("there was an error getting the deal submission time", SuppressWarnings)
            Return ""
        End Try
    End Function
    Public Function GetVendor(DealID As String, Optional SuppressWarnings As Boolean = True) As String
        Try
            Return sqlInterface.SelectData("Vendor", "DealID = " & DealID)
        Catch
            ShoutError("there was an error getting the vendor", SuppressWarnings)

            Return ""
        End Try
    End Function
End Class
