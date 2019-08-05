Imports System.Xml

Partial Class AddDeal

    Public Function WriteXMlOutput(DealData As Dictionary(Of String, String)) As String
        Try
            Dim settings As XmlWriterSettings = New XmlWriterSettings With {
                .Indent = True
            }

            Dim filePath As String = IO.Path.GetTempPath & "\dealinfo.xml"

            ' Create XmlWriter.
            Using writer As XmlWriter = XmlWriter.Create(filePath, settings)

                writer.WriteStartDocument()
                writer.WriteStartElement("Deal")

                For Each tFile As KeyValuePair(Of String, String) In DealData


                    writer.WriteElementString(tFile.Key, tFile.Value)


                Next

                writer.WriteEndElement()
                writer.WriteEndDocument()

            End Using
            Return filePath
        Catch
            Return ""
        End Try

    End Function

End Class
