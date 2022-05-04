Imports System.Diagnostics
Imports System.Xml
Imports String_Extensions

Partial Class ThisAddIn

    Public Sub RemoteDBAddition(ByRef inboundMail As Outlook.MailItem)

        If inboundMail.Subject.Contains("Extension") Then
            Call RemoteExtension(inboundMail)
            Exit Sub
        End If

        Dim dealsRead As New List(Of Dictionary(Of String, String))
        Dim replyMail As Outlook.MailItem
        replyMail = Nothing

        Dim replyMailTwo As Outlook.MailItem
        replyMailTwo = Nothing

        For Each tAttachment As Outlook.Attachment In inboundMail.Attachments
            If tAttachment.FileName.ToLower.Contains(".xml") Then
                Dim fileName As String = IO.Path.GetTempPath & RandomString(6) & tAttachment.FileName
                tAttachment.SaveAsFile(fileName)
                Dim doc As XmlDocument = New XmlDocument
                Dim nodeList As XmlNodeList
                doc.PreserveWhitespace = True


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
                        Try
                            tCreateDealRecord.Add("NDT", deal.SelectSingleNode("NDT").InnerText)

                        Catch ex As Exception
                            tCreateDealRecord.Add("NDT", "")

                        End Try

                        dealsRead.Add(tCreateDealRecord)
                    Next
                    My.Computer.FileSystem.DeleteFile(fileName)
                Catch ex As Exception
                    Debug.WriteLine(ex.Message)
                End Try
            ElseIf tAttachment.FileName.ToLower.Contains(".msg") Then
                Dim fileName As String
                Dim tempPath As String = Environ("TEMP") & "\clsnextdeskticket\" & RandomString(18) & "\"

                If (Not System.IO.Directory.Exists(tempPath)) Then
                    System.IO.Directory.CreateDirectory(tempPath)
                End If
                fileName = tempPath & LegalFileName(tAttachment.FileName) & ".msg"

                tAttachment.SaveAsFile(fileName)

                Try
                    Dim Mail As Outlook.MailItem = Globals.ThisAddIn.Application.GetNamespace("MAPI").OpenSharedItem(fileName)
                    replyMail = Mail.ReplyAll
                    replyMailTwo = Mail.ReplyAll
                Catch
                    Debug.WriteLine("Could Not reply To attached mail")
                End Try

                Try
                    My.Computer.FileSystem.DeleteFile(fileName)
                Catch
                    Debug.WriteLine("Can't delete File")
                End Try

            End If
        Next



        For Each deal As Dictionary(Of String, String) In dealsRead

            If Not IsNothing(replyMail) Then
                Dim tAddDeal As New AddDeal(replyMail)
                If Not tAddDeal.DoNewCreation(deal, replyMail, replyMailTwo) Then
                    Debug.WriteLine("returned false")
                Else
                    Try
                        inboundMail.Delete()
                    Catch
                    End Try
                End If
            End If

        Next

    End Sub


    Public Sub RemoteExtension(ByRef InboundMail As Outlook.MailItem)

        Dim dealsRead As New List(Of Dictionary(Of String, String))
        Dim RequestMail As Outlook.MailItem
        RequestMail = Nothing

        For Each tAttachment As Outlook.Attachment In InboundMail.Attachments
            If tAttachment.FileName.ToLower.Contains(".xml") Then
                Dim fileName As String = IO.Path.GetTempPath & RandomString(6) & tAttachment.FileName
                tAttachment.SaveAsFile(fileName)
                Dim doc As XmlDocument = New XmlDocument
                Dim nodeList As XmlNodeList
                doc.PreserveWhitespace = True


                Try
                    doc.Load(fileName)
                    nodeList = doc.SelectNodes("Deal")

                    For Each deal As XmlNode In nodeList
                        Dim dealModification As New Dictionary(Of String, String) From {
                            {"DealID", deal.SelectSingleNode("DealID").InnerText},
                            {"Action", deal.SelectSingleNode("Action").InnerText}
                        }
                        dealsRead.Add(dealModification)
                    Next
                    My.Computer.FileSystem.DeleteFile(fileName)
                Catch ex As Exception
                    Debug.WriteLine(ex.Message)
                End Try

            ElseIf tAttachment.FileName.ToLower.Contains(".msg") Then
                Dim fileName As String
                Dim tempPath As String = Environ("TEMP") & "\clsnextdeskticket\" & RandomString(18) & "\"

                If (Not System.IO.Directory.Exists(tempPath)) Then
                    System.IO.Directory.CreateDirectory(tempPath)
                End If
                fileName = tempPath & LegalFileName(tAttachment.FileName) & ".msg"

                tAttachment.SaveAsFile(fileName)
                Try
                    Dim Mail As Outlook.MailItem = Globals.ThisAddIn.Application.GetNamespace("MAPI").OpenSharedItem(fileName)
                    RequestMail = Mail
                Catch
                    Debug.WriteLine("Could Not reply To attached mail")
                End Try

                Try
                    My.Computer.FileSystem.DeleteFile(fileName)
                Catch
                    Debug.WriteLine("Can't delete File")
                End Try

            End If
        Next

        For Each deal As Dictionary(Of String, String) In dealsRead

            If Not IsNothing(RequestMail) Then
                If deal("Action") = "Extended" Then
                    Dim DealIDForm As New DealIdent(New List(Of Microsoft.Office.Interop.Outlook.MailItem) From {RequestMail}, "ExtensionMessage", True, deal("DealID"))
                    DealIDForm.Show()
                    Try
                        InboundMail.Delete()
                    Catch
                    End Try

                ElseIf deal("Action") = "Clone" Then
                    'the "new" sub contains every action needed
                    Dim cloneLaterForm As New CloneLater(ReadDate(RequestMail), InboundMail, RequestMail, True, DealID:=deal("DealID"))

                    cloneLaterForm.Dispose()

                End If
            End If
        Next

    End Sub

End Class
