Public Class CiscoDiscounts

    Private ReadOnly AllDiscounts As New Dictionary(Of String, String) From { 'Friendly name, KDFID'
        {"Hunting", "check-BR-OIPU-150725-10481"},
        {"Teaming", "check-BR-Team-150725-10662"},
        {"Migration", "check-BR-MIPE-180910-02363"},
        {"Competitive Migration", "check-BR-MIPC-180910-02367"},
        {"Pre-Sales", "check-BR-Publ-140726-18202"},
        {"PSPP", "check-BR-PSPP-160729-31128"},
        {"Networking Academy", "check-BR-Netw-160729-51688"}
    }

    Public DiscountsRequested As New Dictionary(Of String, Boolean)

    Public Sub New(RequestedDiscounts As String())
        For Each Discount In RequestedDiscounts
            Try
                DiscountsRequested.Add(AllDiscounts(Discount), True)
            Catch
                MsgBox("Could not find the internal discount ID for " & Discount)
            End Try
        Next
    End Sub

End Class

