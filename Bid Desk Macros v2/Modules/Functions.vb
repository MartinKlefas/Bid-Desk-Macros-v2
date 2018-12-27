Imports Microsoft.Office.Interop.Outlook

Partial Class ThisAddIn
    Public sqlInterface As New ClsDatabase(ThisAddIn.server, ThisAddIn.user,
                                   ThisAddIn.database, ThisAddIn.port)

    Public userform3 As DealIdent

    Public Function MoveToFolder(folderName As String, thisMailItem As MailItem, Optional suppressWarnings As Boolean = False) As Boolean
        Dim mailboxNameString As String
        mailboxNameString = "Martin.Klefas@insight.com"

        Dim olApp As New Outlook.Application
        Dim olNameSpace As Outlook.NameSpace

        Dim olDestFolder As Outlook.MAPIFolder


        olNameSpace = olApp.GetNamespace("MAPI")

        'OLD: Set olDestFolder = olNameSpace.Folders(mailboxNameString).Folders("Martins Emails").Folders(folderName)
        Try
            olDestFolder = olNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders("Bids").Folders(folderName)
            thisMailItem.Move(olDestFolder)
            MoveToFolder = True
        Catch
            Try
                olDestFolder = olNameSpace.Folders(mailboxNameString).Folders("Inbox").Folders("Bids").Folders.Add(folderName)
                thisMailItem.Move(olDestFolder)
                MoveToFolder = True
            Catch
                If Not suppressWarnings Then MsgBox("Couldn't move the mail for some reason...")
                MoveToFolder = False
            End Try
        End Try
    End Function

    Public Function GetFolderbyDeal(dealID As String,
                                    Optional SuppressWarnings As Boolean = False) As String

        Try
            Return sqlInterface.SelectData("AM", "DealID = 16859207")
        Catch
            If Not SuppressWarnings Then
                MsgBox("there was an error")
            End If
            Return ""
        End Try
    End Function

    Public Function FindDealID(MsgSubject As String, msgBody As String, Optional completeAutonomy As Boolean = False) As String
        Dim myAr As String(), i As Integer, myArTwo As Object
        myAr = Split(MsgSubject, " ")

        userform3.DealID.Text = ""

        For i = LBound(myAr) To UBound(myAr)
            '~~> This will give you the contents of your email
            '~~> on separate lines
            myAr(i) = Trim(myAr(i))

            If Len(myAr(i)) > 4 Then
                If myAr(i).StartsWith("P00", ThisAddIn.searchType) Or
                    myAr(i).StartsWith("E00", ThisAddIn.searchType) Then
                    If Mid(LCase(myAr(i)), Len(myAr(i)) - 2, 2) = "-v" Then myAr(i) = Left(myAr(i), Len(myAr(i)) - 3)
                    userform3.DealID.Text = Trim(myAr(i))
                End If
                If myAr(i).StartsWith("REGI-", ThisAddIn.searchType) Or
                    myAr(i).StartsWith("REGE-", ThisAddIn.searchType) Then
                    userform3.DealID.Text = Trim(myAr(i))
                End If
            End If
        Next i

        If userform3.DealID.Text = "" Then
            myAr = Split(msgBody, vbCrLf)
            For i = LBound(myAr) To UBound(myAr)
                '~~> This will give you the contents of your email
                '~~> on separate lines
                If Len(myAr(i)) > 8 Then
                    If myAr(i).StartsWith("Deal ID:", ThisAddIn.searchType) Then
                        myArTwo = Split(myAr(i))
                        userform3.DealID.Text = Trim(myArTwo(2))
                    End If
                End If
            Next
        End If

        If i + 3 < UBound(myAr) Then
            If myAr(i) = "Quote" And myAr(i + 1) = "Review" And myAr(i + 2) = "Quote" Then
                userform3.DealID.Text = Trim(myAr(i + 4))
            End If
            If myAr(i) = "QUOTE" And myAr(i + 1) = "Deal" And myAr(i + 3) = "Version" Then
                userform3.DealID.Text = Trim(myAr(i + 2))
            End If

        End If

        If Not completeAutonomy Then
            userform3.ShowDialog()

        End If



        FindDealID = userform3.DealID.Text
    End Function

    Private Function RecordWaitTime(receivedTime As Date, completedTime As Date, person As String) As String


        Dim tmpDict As New Dictionary(Of String, String) From {
            {"WaitingFor", person},
            {"Received", receivedTime},
            {"Completed", completedTime}
        }


        If sqlInterface.Add_Data(tmpDict, "wait_times") Then
            Return PrettyString(completedTime - receivedTime)
        Else
            Return ""
        End If


    End Function


    Function CreateDealRecord(ReplyMail As Outlook.MailItem) As Dictionary(Of String, String)
        Dim NewDealForm As New newDeal
        Dim requestorName As String, Vendor As String, ccNames As String

        If NewDealForm.ShowDialog() = Windows.Forms.DialogResult.OK Then ' Show and wait
            Dim toNames As String(), rName() As String

            toNames = Split(ReplyMail.To, ";") ' Split out each recipient

            If InStr(toNames(0), ",") > 1 Then ' Some email names are "fName, lName" others aren't

                rName = Split(toNames(0), ",")
                requestorName = Trim(rName(1)) & " " & Trim(rName(0))
            Else
                requestorName = Trim(toNames(0))
            End If

            If NewDealForm.DellOption.Checked Then
                Vendor = "Dell"
            ElseIf NewDealForm.HPIOption.Checked Then
                Vendor = "HPI"
            Else
                Vendor = "HPE"
            End If

            ccNames = ReplyMail.CC
            For i = 1 To UBound(toNames) ' append the second and later "to" names to the CC list
                ccNames = ccNames & "; " & toNames(i)
            Next

            CreateDealRecord = New Dictionary(Of String, String) From {
                {"AM", requestorName},
                {"Customer", NewDealForm.CustomerName.Text},
                {"Vendor", Vendor},
                {"DealID", NewDealForm.DealID.Text},
                {"Ingram", NewDealForm.cIngram.Checked},
                {"Techdata", NewDealForm.cTechData.Checked},
                {"Westcoast", NewDealForm.cWestcoast.Checked},
                {"CC", ccNames}
            }


            If sqlInterface.Add_Data(CreateDealRecord) Then
                CreateDealRecord.Add("Result", "Success")
            Else
                CreateDealRecord.Add("Result", "Failed")
            End If


        Else
                Return New Dictionary(Of String, String) From {
                {"Result", "Cancelled"}
            }



        End If






    End Function


    Function GetCurrentItem() As Object
        Select Case True
            Case IsExplorer(Application.ActiveWindow)
                GetCurrentItem = Application.ActiveExplorer.Selection.Item(1)
            Case IsInspector(Application.ActiveWindow)
                GetCurrentItem = Application.ActiveInspector.CurrentItem
            Case Else
                GetCurrentItem = Nothing
        End Select
    End Function
    Function IsExplorer(itm As Object) As Boolean
        IsExplorer = (TypeName(itm) = "Explorer")
    End Function
    Function IsInspector(itm As Object) As Boolean
        IsInspector = (TypeName(itm) = "Inspector")
    End Function
End Class
