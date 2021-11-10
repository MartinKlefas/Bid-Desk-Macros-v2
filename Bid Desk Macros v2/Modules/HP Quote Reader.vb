Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports String_Extensions

Module HP_Quote_Reader
    Public Function RipFromFile(tAttachment As Attachment, CurrentGuess As String) As String
        If tAttachment.FileName.ToLower.Contains(".csv") Then
            Dim fName As String = Path.GetTempPath() & RandomString(6) & "quote.csv"
            Try
                tAttachment.SaveAsFile(fName)
                Dim quoteCsvString As String = File.ReadAllText(fName)
                quoteCsvString = Replace(quoteCsvString, vbNullChar, "")
                Dim quoteArry As String() = Split(quoteCsvString, "-")

                For Each fragment As String In quoteArry

                    If fragment.ToLower.Contains("p0") Or fragment.ToLower.Contains("e0") Or fragment.ToLower.Contains("nq0") Then
                        Dim subfragments As String() = Split(fragment, ",")
                        For Each sfrag In subfragments
                            sfrag = Replace(sfrag, Chr(34), "")
                            If sfrag <> "" AndAlso (sfrag.ToLower.StartsWith("p0") Or sfrag.ToLower.StartsWith("e0") Or sfrag.ToLower.StartsWith("nq0")) Then
                                Dim OPG As String = CurrentGuess

                                Globals.ThisAddIn.AddOPG(sfrag, OPG)

                                CurrentGuess = sfrag

                                Exit For
                            End If

                        Next

                    End If
                Next
                File.Delete(fName)
            Catch
                Debug.WriteLine("Error while saving/processing CSV file")
            End Try

        ElseIf tAttachment.FileName.ToLower.EndsWith("xlsx") Then
            Dim fName As String = Path.GetTempPath() & RandomString(6) & tAttachment.FileName.WinSafeFileName
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

        ElseIf tAttachment.FileName.ToLower.EndsWith("xls") Then
            'the techdata "xls" files are actually html
            Dim fName As String = Path.GetTempPath() & RandomString(6) & tAttachment.FileName.WinSafeFileName
            Try
                tAttachment.SaveAsFile(fName)
            Catch
                Debug.WriteLine("Error while saving xlsx file")
            End Try
            Dim tmpDealID As String = ""

            Try
                Dim AllHTMl As String = My.Computer.FileSystem.ReadAllText(fName)
                tmpDealID = Mid(AllHTMl, AllHTMl.IndexOf("NQ0"), 11)
                tmpDealID = TrimExtended(tmpDealID)
            Catch
                Debug.WriteLine("Error processing xls file")
            End Try

            Globals.ThisAddIn.AddOPG(tmpDealID, CurrentGuess)

            Try
                File.Delete(fName)
            Catch
                Debug.WriteLine("Error deleting xls file")
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
                            Using conn As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(conStr)

            ' Notice: I used a fully qualified name 
            ' because Microsoft.Office.Interop.Excel contains also a class named OleDbConnection

            conn.Open()
            Using cmd As OleDbCommand = New OleDbCommand("select * from [" & sheet & "$]", conn)

                Using dataReader As OleDbDataReader = cmd.ExecuteReader()
                    Dim tempStr As String = ""

                    ' Create a new connection object and open it


                    ' Create command text with SQL-style syntax
                    ' Notice: First sheet is named Sheet1. In the command, sheet's name is followed with dollar sign!

                    ' Get data from Excel's sheet to OleDb datareader object

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
                End Using
            End Using
        End Using
    End Function
End Module
