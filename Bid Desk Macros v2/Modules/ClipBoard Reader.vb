Imports String_Extensions
Module ClipBoard_Reader
    Public Function FindDealID(ClipBoardText As String) As String
        Dim strArry As String()

        'Old HP SQs
        FindDealID = ""
        If InStr(1, ClipBoardText, "SQ-") > 0 Then
            FindDealID = Mid(ClipBoardText, InStr(1, ClipBoardText, "SQ-"), 10)

        End If

        'HPI/HPE Quotes
        If InStr(1, ClipBoardText, "Quote Number") > 0 AndAlso FindDealID = "" Then
            If InStr(1, ClipBoardText, "Quote Number: ") > 0 Then
                FindDealID = Mid(ClipBoardText, InStr(1, ClipBoardText, "Quote Number: ") + 14, 10)
            Else
                FindDealID = Mid(ClipBoardText, InStr(1, ClipBoardText, "Quote Number") + 14, 10)
            End If

        End If

        'Dell Deal Registration
        If InStr(1, ClipBoardText, "Deal ID") > 0 AndAlso FindDealID = "" Then
            strArry = Split(Mid(ClipBoardText, InStr(1, ClipBoardText, "Deal ID")), vbTab)
            Try
                FindDealID = strArry(1)
            Catch
                FindDealID = ""
            End Try
        End If


        'HPI/HPE Deal Registration
        If InStr(1, ClipBoardText.ToLower, "deal registration id") > 0 AndAlso FindDealID = "" Then
            strArry = Split(Mid(ClipBoardText, InStr(1, ClipBoardText.ToLower, "deal registration id")), vbCrLf)
            Try
                FindDealID = strArry(1)
            Catch
                FindDealID = ""
            End Try
        End If

        'Miscrosoft Hardware Deal Registration

        If InStr(1, ClipBoardText, "CAS-") > 0 Then
            FindDealID = Mid(ClipBoardText, InStr(1, ClipBoardText, "CAS-"), 19)

        End If


        'Lenovo Bid Request

        If InStr(1, ClipBoardText, "BBR-") > 0 Then
            FindDealID = Mid(ClipBoardText, InStr(1, ClipBoardText, "BBR-"), 13)
        End If


        FindDealID = TrimExtended(FindDealID)
    End Function

    Function FindVendor(ClipBoardText As String) As String
        FindVendor = "Dell"


        If InStr(1, ClipBoardText, "E00") > 0 Or InStr(1, ClipBoardText, "NQ0") > 0 Then
            FindVendor = "HPE"
        End If
        If InStr(1, ClipBoardText, "P00") > 0 Then
            FindVendor = "HPI"
        End If

        If InStr(1, ClipBoardText, "Deal Registration id") > 0 Then
            FindVendor = "HPI"
        End If

        If InStr(1, ClipBoardText, "CAS-") > 0 Then
            FindVendor = "Microsoft"
        End If

        If InStr(1, ClipBoardText, "BBR-") > 0 Then
            FindVendor = "Lenovo"

        End If
    End Function

    Function FindCustomer(ClipboardText As String) As String
        Dim strArry As String()
        FindCustomer = ""

        If InStr(1, ClipboardText, "Full Legal Name") > 0 Then
            strArry = Split(Mid(ClipboardText, InStr(1, ClipboardText, "Full Legal Name")), vbCrLf)
            FindCustomer = StrConv(strArry(2), vbProperCase)
        End If

        If InStr(1, ClipboardText, "Customer: ") > 0 Then
            strArry = Split(Mid(ClipboardText, InStr(1, ClipboardText, "Customer: ") + 10), vbCrLf)
            FindCustomer = StrConv(strArry(0), vbProperCase)
        End If

        If InStr(1, ClipboardText, "End User Account Name") > 0 Then
            strArry = Split(Mid(ClipboardText, InStr(1, ClipboardText, "End User Account Name")), vbTab)
            Try
                FindCustomer = strArry(1)
            Catch
                FindCustomer = ""
            End Try
        End If



        If InStr(1, ClipboardText, "Deal Registration id") > 0 Then
            strArry = Split(Mid(ClipboardText, InStr(1, ClipboardText, "Deal Registration id")), vbCrLf)
            Try
                FindCustomer = strArry(3)
            Catch
                FindCustomer = ""
            End Try
        End If
        If InStr(1, ClipboardText, "Opportunity ID") > 0 Then
            strArry = Split(Mid(ClipboardText, InStr(1, ClipboardText, "Opportunity ID")), vbCrLf)
            Try
                FindCustomer = Left(strArry(7), InStr(strArry(7), vbTab) - 1)
            Catch
                FindCustomer = ""
            End Try
        End If



        If InStr(1, ClipboardText, "BBR-") > 0 Then
            FindCustomer = Split(Mid(ClipboardText, InStr(1, ClipboardText, "End Customer Name") + 18), vbCrLf)(0)
        End If

        FindCustomer = TrimExtended(FindCustomer)

    End Function
End Module
