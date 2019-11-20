Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Text.RegularExpressions

Public Module StringExtensions

    <Extension()>
    Public Sub Print(ByVal aString As String)
        Console.WriteLine(aString)
    End Sub

    <Extension()>
    Public Sub Append(ByRef aString As String,
                      ByVal bString As String)
        aString &= bString
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
        aString = Replace(aString, vbTab, " ")
        aString = Replace(aString, vbCrLf, " ")
        aString = Replace(aString, vbLf, " ")
        aString = Replace(aString, vbCr, " ")
        aString = Trim(aString)
        Return aString
    End Function

    <Extension()>
    Public Function WinSafeFileName(ByVal aString As String) As String

        Dim pattern As New Regex("[0-9a-zA-Z-._]")

        Dim result As String = ""

        For Each tChar As Char In aString
            Dim letter = tChar.ToString
            If pattern.IsMatch(letter) Then
                result &= letter
            End If
        Next

        Return result
    End Function

    Public Function RemoveSpaces(ByVal aString As String) As String
        aString = ReplaceSpaces(aString)
        aString = Replace(aString, vbTab, " ")
        aString = Replace(aString, vbCrLf, " ")
        aString = Replace(aString, " ", "")
        Return aString
    End Function


    Public Function RandomString(ByVal Length As Integer) As String
        Dim s As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz"
        Dim r As New Random
        Dim sb As New StringBuilder

        For i As Integer = 1 To Length
            Dim idx As Integer = r.Next(0, 61)
            sb.Append(s.Substring(idx, 1))
        Next

        Return sb.ToString
    End Function

    <Extension>
    Public Function ContainsAny(ByVal aString As String, ByVal SearchFor As IEnumerable(Of String)) As Boolean
        Dim result As Boolean = False
        For Each searchString As String In SearchFor
            If Not result Then result = aString.Contains(searchString)
        Next

        Return result
    End Function

    <Extension>
    Public Function StartsWithAny(ByVal aString As String, ByVal SearchFor As IEnumerable(Of String)) As Boolean
        Dim result As Boolean = False
        For Each searchString As String In SearchFor
            If Not result Then result = aString.StartsWith(searchString)
        Next

        Return result
    End Function
End Module

