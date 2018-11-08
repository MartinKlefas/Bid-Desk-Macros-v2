Imports System.Runtime.CompilerServices

Public Module StringExtensions

    <Extension()>
    Public Sub Print(ByVal aString As String)
        Console.WriteLine(aString)
    End Sub

    <Extension()>
    Public Sub Append(ByRef aString As String,
                      ByVal bString As String)
        aString = aString & bString
    End Sub

End Module