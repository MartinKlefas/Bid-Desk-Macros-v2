Imports System.Diagnostics
Imports System.Xml

Partial Class ThisAddIn

    Public Sub RemoteDBAddition(inboundMail As Outlook.MailItem)
        For Each tAttachment As Outlook.Attachment In inboundMail.Attachments
            If tAttachment.FileName.ToLower.Contains("xml") Then
                Dim fileName As String = tAttachment.GetTemporaryFilePath()
                Dim doc As XmlDocument = New XmlDocument
                Dim nodeList As XmlNodeList
                doc.PreserveWhitespace = True

                Dim dealsRead As New List(Of Dictionary(Of String, String))
                Try
                    doc.Load(fileName)
                    nodeList = doc.SelectNodes("Deal")

                    For Each deal As XmlNode In nodeList

                        Dim tCreateDealRecord As New Dictionary(Of String, String) From {
                                {"AMEmailAddress", deal.SelectSingleNode("AMEmailAddress").InnerText},
                                {"AM", deal.SelectSingleNode("AM").InnerText},
                                {"Customer", deal.SelectSingleNode("Customer").InnerText},
                                {"Vendor", deal.SelectSingleNode("Vendor").InnerText},
                                {"DealID", deal.SelectSingleNode("DealID").InnerText},
                                {"Ingram", deal.SelectSingleNode("Ingram").InnerText},
                                {"Techdata", deal.SelectSingleNode("Techdata").InnerText},
                                {"Westcoast", deal.SelectSingleNode("Westcoast").InnerText},
                                {"CC", deal.SelectSingleNode("CC").InnerText},
                                {"Status", "Submitted to Vendor"},
                                {"StatusDate", deal.SelectSingleNode("Date").InnerText},
                                {"Date", deal.SelectSingleNode("Date").InnerText}
                            }

                        dealsRead.Add(tCreateDealRecord)
                    Next

                Catch ex As Exception
                    Debug.WriteLine("error, " & ex.Message)
                End Try


            End If
        Next
    End Sub


End Class
