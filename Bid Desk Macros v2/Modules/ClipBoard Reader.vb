Module ClipBoard_Reader
    Public Function FindDealID(ClipBoardText As String) As String
        Dim strArry As String()
        FindDealID = ""
        If InStr(1, ClipBoardText, "SQ-") > 0 Then
            FindDealID = Mid(ClipBoardText, InStr(1, ClipBoardText, "SQ-"), 10)

        End If

        If InStr(1, ClipBoardText, "Deal ID") > 0 Then
            strArry = Split(Mid(ClipBoardText, InStr(1, ClipBoardText, "Deal ID")), vbTab)
            FindDealID = strArry(1)
        End If

        If InStr(1, ClipBoardText, "HP Opportunity ID") > 0 Then
            strArry = Split(Mid(ClipBoardText, InStr(1, ClipBoardText, "HP Opportunity ID")), vbCrLf)
            FindDealID = strArry(5)

        End If


    End Function

    Function FindVendor(ClipBoardText As String) As String
        FindVendor = ""

        If InStr(1, ClipBoardText, "SQ-") > 0 Then
            FindVendor = "HPI"
        End If



        If InStr(1, ClipBoardText, "HP Opportunity ID") > 0 Then
            FindVendor = "HPI"
        End If
    End Function

    Function FindCustomer(ClipboardText As String) As String
        Dim strArry As String()
        FindCustomer = ""

        If InStr(1, ClipboardText, "Full Legal Name") > 0 Then
            strArry = Split(Mid(ClipboardText, InStr(1, ClipboardText, "Full Legal Name")), vbCrLf)
            FindCustomer = StrConv(strArry(2), vbProperCase)
        End If

        If InStr(1, ClipboardText, "End User Account Name") > 0 Then
            strArry = Split(Mid(ClipboardText, InStr(1, ClipboardText, "End User Account Name")), vbTab)
            FindCustomer = strArry(1)
        End If



        If InStr(1, ClipboardText, "HP Opportunity ID") > 0 Then
            strArry = Split(Mid(ClipboardText, InStr(1, ClipboardText, "HP Opportunity ID")), vbCrLf)

            FindCustomer = strArry(13)

        End If


    End Function
End Module
