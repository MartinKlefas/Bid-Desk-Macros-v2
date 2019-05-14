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
        aString = Replace(aString, CStr(ChrW(8194)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8195)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8196)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8197)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8198)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8199)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8200)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8201)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8202)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8203)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8204)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8205)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8206)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8207)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(8239)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(12288)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(12351)), CStr(Chr(32)))
        aString = Replace(aString, CStr(ChrW(65279)), CStr(Chr(32)))

        Return aString
    End Function

    <Extension()>
    Public Function TrimExtended(ByVal aString As String) As String
        aString = ReplaceSpaces(aString)
        aString = Trim(aString)
        Return aString
    End Function
End Module

