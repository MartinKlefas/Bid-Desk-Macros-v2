Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions

Module HP_Quote_Reader
    Public Function RipFromFile(tAttachment As Attachment, CurrentGuess As String) As String
        If tAttachment.FileName.ToLower = "quote.csv" Then
            Dim fName As String = Path.GetTempPath() & "quote.csv"
            Try
                tAttachment.SaveAsFile(fName)
                Dim quoteCsvString As String = File.ReadAllText(fName)
                quoteCsvString = Replace(quoteCsvString, vbNullChar, "")
                Dim quoteArry As String() = Split(quoteCsvString, "-")
                For Each fragment As String In quoteArry
                    If fragment.ToLower.StartsWith("p0") Or fragment.ToLower.StartsWith("e0") Then

                        Dim OPG As String = CurrentGuess

                        Globals.ThisAddIn.AddOPG(fragment, OPG)

                        CurrentGuess = fragment

                        Exit For
                    End If
                Next
                File.Delete(fName)
            Catch
                Debug.WriteLine("Error while saving/processing CSV file")
            End Try

        ElseIf tAttachment.FileName.ToLower.EndsWith("xlsx") Then
            Dim fName As String = Path.GetTempPath() & tAttachment.FileName
            Try
                tAttachment.SaveAsFile(fName)
            Catch
                Debug.WriteLine("Error while saving xlsx file")
            End Try
            Dim tmpDealID As String = ""

            Try
                tmpDealID = ReadExcel(fName, "Sheet1", 2, 2)
                tmpDealID = Strings.Left(tmpDealID, Len(tmpDealID) - 3)
            Catch
                Debug.WriteLine("Error processing xlsx file")
            End Try

            Globals.ThisAddIn.AddOPG(tmpDealID, CurrentGuess)

            Try
                File.Delete(fName)
            Catch
                Debug.WriteLine("Error deleting xlsx file")
            End Try


            CurrentGuess = tmpDealID
        End If

        Return CurrentGuess
    End Function




    Public Function ReadExcel(file As String, sheet As String, row As Integer, column As Integer) As String

        row -= 1 'db access is 0 based, excel references are 1 based
        column -= 1

        Dim conStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & file & ";Extended Properties='Excel 12.0 Xml;HDR=No;'"
        ' HDR=Yes skips first row which contains headers for the columns
        Dim conn As System.Data.OleDb.OleDbConnection ' Notice: I used a fully qualified name 
        ' because Microsoft.Office.Interop.Excel contains also a class named OleDbConnection
        Dim cmd As OleDbCommand
        Dim dataReader As OleDbDataReader
        Dim tempStr As String = ""

        ' Create a new connection object and open it
        conn = New System.Data.OleDb.OleDbConnection(conStr)
        conn.Open()
        ' Create command text with SQL-style syntax
        ' Notice: First sheet is named Sheet1. In the command, sheet's name is followed with dollar sign!
        cmd = New OleDbCommand("select * from [" & sheet & "$]", conn)
        ' Get data from Excel's sheet to OleDb datareader object
        dataReader = cmd.ExecuteReader()
        Dim curRow As Integer = 0
        ' Read rows until an empty row is found
        While (dataReader.Read())
            ' Index of column B is 0 because it is range's first column
            tempStr = dataReader.GetValue(column).ToString()
            If curRow = row Then Exit While
            curRow += 1
        End While

        If curRow = row Then
            Return tempStr
        Else
            Return ""
        End If
    End Function
End Module
