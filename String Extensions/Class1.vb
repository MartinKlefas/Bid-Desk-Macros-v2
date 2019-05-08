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

    <Extension()>
    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Return value.Count(Function(c As Char) c = ch)
    End Function

    <Extension()>
    Public Function ReplaceSpaces(ByVal aString As String) As String
        aString = Replace(aString, CStr(Chr(160)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8194)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8195)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8196)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8197)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8198)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8199)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8200)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8201)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8202)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8203)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8204)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8205)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8206)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8207)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(8239)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(12288)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(12351)), CStr(Chr(32)))
        aString = Replace(aString, CStr(Chr(65279)), CStr(Chr(32)))

        Return aString
    End Function
End Module