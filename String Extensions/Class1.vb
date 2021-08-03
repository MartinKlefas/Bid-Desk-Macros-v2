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
        aString = Replace(aString, "[", " ")
        aString = Replace(aString, "]", " ")

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
    Public Function ContainsAny(ByVal aString As String, ByVal SearchFor As IEnumerable(Of String), Optional ByVal CaseSensitive As Boolean = False) As Boolean
        Dim result As Boolean = False

        If CaseSensitive Then
            For Each searchString As String In SearchFor
                If Not result Then result = aString.Contains(searchString)
            Next
        Else
            For Each searchString As String In SearchFor
                If Not result Then result = aString.ToLower.Contains(searchString.ToLower)
            Next
        End If

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

    ''' <summary>
    ''' Removes illegal characters from a string to make it acceptable for windows. Removes "\/|?*'lt''gt'"":"
    ''' Also discards any characters that make the string longer than 260
    ''' </summary>
    ''' <param name="strFileNameIn">Input filename</param>
    ''' <returns>a legal filename in a string.</returns>
    <Extension>
    Public Function LegalFileName(strFileNameIn As String) As String
        Dim i As Integer

        Const strIllegals = "\/|?*<>"":"
        LegalFileName = strFileNameIn
        For i = 1 To Len(strIllegals)
            LegalFileName = Replace(LegalFileName, Mid$(strIllegals, i, 1), "")
        Next i

        If Len(LegalFileName) > 259 Then
            LegalFileName = Left(LegalFileName, 259)
        End If

    End Function

    ''' <summary>
    ''' Removes Numbers and space-like characters from the beginning and end of a string.
    ''' </summary>
    ''' <param name="aString">input string to be trimmed</param>
    ''' <returns>number-less string</returns>
    Public Function TrimNumbers(ByVal aString As String, Optional ByVal trimspaces As Boolean = True) As String

        If trimspaces Then aString = TrimExtended(aString)

        Dim StrArry = aString.Split(" ")

        Dim strList As New List(Of String)

        For Each item In StrArry
            strList.Add(item)
        Next



        While Regex.IsMatch(strList.First, "([0-9])+") Or strList.First = ""
            strList.RemoveAt(0)

        End While

        While Regex.IsMatch(strList.Last, "([0-9])+") Or strList.Last = ""
            strList.RemoveAt(strList.Count - 1)

        End While

        Return Strings.Join(strList.ToArray, " ")
    End Function

    ''' <summary>
    ''' Lazily and Greedily Removes all HTML tags from an input string.
    ''' The function does not check if it's an actual HTML tag or similar, only that it starts with open and closed diagonal brackets before deleting it.
    ''' </summary>
    ''' <param name="aString">The String to be cleaned</param>
    ''' <returns>A plain string with no HTML in it</returns>
    Public Function StripHTML(ByVal aString As String) As String

        Dim firstOpen, nextClose As Integer

        While aString.Contains("<") And aString.Contains(">")
            Dim strBuilder As String = ""
            firstOpen = InStr(aString, "<")
            nextClose = InStr(firstOpen, aString, ">")

            If firstOpen > 0 Then
                strBuilder = Left(aString, firstOpen)

            End If

            strBuilder &= Mid(aString, nextClose)

            aString = strBuilder

        End While

        Return aString

    End Function

    <Extension()>
    Public Function SplitByWord(ByVal Input As String, ByVal Word As String) As String()

        Dim rgx As New Regex(Word)

        SplitByWord = rgx.Split(Input)

    End Function
End Module

