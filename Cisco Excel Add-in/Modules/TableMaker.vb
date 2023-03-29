Imports String_Extensions
Imports OfficeOpenXml
Imports System.Text.RegularExpressions

Module TableMaker

    Function ExcelFile(Data As List(Of Dictionary(Of String, String)), TableFields As List(Of String)) As String
        Try
            Dim tempPath As String = Environ("TEMP") & "\backorder\" & RandomString(18) & "\"

            If (Not System.IO.Directory.Exists(tempPath)) Then
                System.IO.Directory.CreateDirectory(tempPath)
            End If

            Dim FileName As String = tempPath & "PartCodes.xlsx"

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            Using p As New ExcelPackage

                Dim ws As ExcelWorksheet = p.Workbook.Worksheets.Add("Part Codes")
                ws.Row(1).Style.Font.Bold = True
                ws.Cells.Style.WrapText = False

                Dim colNum As Integer = 1

                For Each field In TableFields
                    ws.Cells(1, colNum).Value = field
                    colNum += 1
                Next



                Dim rownum As Integer = 2
                For Each deal As Dictionary(Of String, String) In Data
                    colNum = 1
                    For Each field In TableFields
                        Try
                            Dim fieldData As Object = SanitizeData(deal(field))
                            ws.Cells(rownum, colNum).Value = fieldData
                            If TypeOf fieldData Is Double AndAlso field <> "Product Code" Then
                                ws.Cells(rownum, colNum).Style.Numberformat.Format = "#,##0.00"
                            End If


                        Catch
                            ws.Cells(rownum, colNum).Value = ""
                        End Try

                        colNum += 1

                    Next


                    rownum += 1
                Next

                p.SaveAs(New IO.FileInfo(FileName))

            End Using

            Return FileName
        Catch
            Return ""
        End Try

    End Function


    Function SanitizeData(value As String) As Object

        If Regex.IsMatch(value, "([0-9])*-([0-9])*-([0-9])*T([0-9])*:([0-9])*:([0-9])*") Then
            Return Strings.Left(value, InStr(value, "T") - 1)
        Else
            Try
                If IsNumeric(value) Then
                    Return Convert.ToDouble(value)
                Else
                    Return value
                End If

            Catch ex As Exception

                Return value
            End Try
        End If


    End Function

    Function ReadTable(ws As Excel.Worksheet) As Object
        Dim row, numColumns As Integer

        Dim tHeader As String

        Dim readHeaders As New List(Of String)

        Dim readData As New List(Of Dictionary(Of String, String))

        row = 1

        While ws.Cells(row, 1).value <> ""
            tHeader = TrimNumbers(ws.Cells(row, 1).value)
            readHeaders.Add(TrimNumbers(tHeader))

            row += 1
        End While


        readHeaders = readHeaders.Distinct.ToList

        numColumns = readHeaders.Count

        Dim tCol As Integer = 1
        Dim tData As New Dictionary(Of String, String)
        row = 1

        While ws.Cells(row, 1).value <> ""

            tHeader = TrimNumbers(ws.Cells(row, 1).value)

            tData.Add(tHeader, ws.Cells(row, 2).value)




            If tCol = numColumns Then
                readData.Add(tData)

                tCol = 1

                tData = New Dictionary(Of String, String)

            Else
                tCol += 1
            End If

            row += 1
        End While

        Return {ReadData, ReadHeaders}


    End Function
End Module
