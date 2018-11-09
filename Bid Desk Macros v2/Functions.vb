Imports Microsoft.Office.Interop.Outlook

Module Functions
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
End Module
