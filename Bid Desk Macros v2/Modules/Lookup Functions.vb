Partial Class ThisAddIn
    Public Function GetFolderbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String

        Try
            Return sqlInterface.SelectData("AM", "DealID = " & dealID)
        Catch
            If Not SuppressWarnings Then
                MsgBox("there was an error")
            End If
            Return ""
        End Try
    End Function

    Public Function GetCustomerbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String

        Try
            Return sqlInterface.SelectData("Customer", "DealID = " & dealID)
        Catch
            If Not SuppressWarnings Then
                MsgBox("there was an error")
            End If
            Return ""
        End Try
    End Function
    Public Function GetCCbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String

        Try
            Return sqlInterface.SelectData("CC", "DealID = " & dealID)
        Catch
            If Not SuppressWarnings Then
                MsgBox("there was an error")
            End If
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
            If Not SuppressWarnings Then
                MsgBox("there was an error")
            End If
            Return ""
        End Try

    End Function

    Public Function AddNewTicketToDeal(DealID As String, TicketNumber As Integer) As Integer
        Dim oldNDT As String, newNDT As String

        oldNDT = GetNDTbyDeal(DealID)
        newNDT = oldNDT & ";" & TicketNumber

        Return sqlInterface.Update_Data("NDT = " & newNDT, "DealID = " & DealID)
    End Function
End Class
