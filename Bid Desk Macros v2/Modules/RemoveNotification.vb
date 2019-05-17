Module RemoveNotification
    ' Code sample by Outlook MVP "Neo"
    ' Removes the New Mail icon from the Windows system tray,
    ' and resets Outlook's new mail notification engine.
    ' Tested against Outlook 2000 (IMO) and Outlook 2002 (POP Account)

    ' Send questions and comments to neo@mvps.org

    ' WARNING: Due to the use of AddressOf, code must
    ' go into a module and not ThisOutlookSession or
    ' a class module

    ' Entry Point is RemoveNewMailIcon.

    Public Const WUM_RESETNOTIFICATION As Long = &H407

    'Required Public constants, types & declares
    'for the Shell_Notify API method
    Public Const NIM_ADD As Long = &H0
    Public Const NIM_MODIFY As Long = &H1
    Public Const NIM_DELETE As Long = &H2

    Public Const NIF_ICON As Long = &H2 'adding an ICON
    Public Const NIF_TIP As Long = &H4 'adding a TIP
    Public Const NIF_MESSAGE As Long = &H1 'want return messages

    ' Structure needed for Shell_Notify API
    Structure NOTIFYICONDATA
        Dim cbSize As Long
        Dim hwnd As Long
        Dim uID As Long
        Dim uFlags As Long
        Dim uCallbackMessage As Long
        Dim hIcon As Long
        Dim szTip As String * 64
End Structure

    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Integer, ByVal wMsg As Integer,
  ByVal wParam As Integer, ByVal lParam As String) As Long

    Declare Function GetClassName Lib "user32" _
      Alias "GetClassNameA" _
      (ByVal hwnd As Long,
      ByVal lpClassName As String,
      ByVal nMaxCount As Long) As Long


    Declare Function GetWindowTextLength Lib "user32" _
      Alias "GetWindowTextLengthA" _
      (ByVal hwnd As Long) As Long

    Declare Function GetWindowText Lib "user32" _
      Alias "GetWindowTextA" _
      (ByVal hwnd As Long,
      ByVal lpString As String,
      ByVal cch As Long) As Long

    Declare Function EnumWindows Lib "user32" _
      (ByVal lpEnumFunc As Long,
      ByVal lParam As Long) As Long

    Declare Function Shell_NotifyIcon Lib "shell32.dll" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long,
      lpData As NOTIFYICONDATA) As Long

    Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
      (ByVal lpClassName As String,
      ByVal lpWindowName As String) As Long

    ' This is the entry point that makes it happen
    Sub RemoveNewMailIcon()
        EnumWindows AddressOf EnumWindowProc, 0
End Sub

    Public Function EnumWindowProc(ByVal hwnd As Long,
      ByVal lParam As Long) As Long

        'Do stuff here with hwnd
        Dim sClass As String
        Dim sIDType As String
        Dim sTitle As String
        Dim hResult As Long

        sTitle = GetWindowIdentification(hwnd, sIDType, sClass)
        If sTitle = "rctrl_renwnd32" Then
            hResult = KillNewMailIcon(hwnd)
        End If

        If hResult Then
            EnumWindowProc = False
            ' Reset the new mail notification engine
            Call SendMessage(hwnd, WUM_RESETNOTIFICATION, 0&, 0&)
        Else
            EnumWindowProc = True
        End If
    End Function

    Private Function GetWindowIdentification(ByVal hwnd As Long,
      sIDType As String,
      sClass As String) As String

        Dim nSize As Long
        Dim sTitle As String

        'get the size of the string required
        'to hold the window title
        nSize = GetWindowTextLength(hwnd)
        'if the return is 0, there is no title
        If nSize > 0 Then
            sTitle = Space$(nSize + 1)
            Call GetWindowText(hwnd, sTitle, nSize + 1)
            sIDType = "title"
            sClass = Space$(64)
            Call GetClassName(hwnd, sClass, 64)
        Else
            'no title, so get the class name instead
            sTitle = Space$(64)
            Call GetClassName(hwnd, sTitle, 64)
            sClass = sTitle
            sIDType = "class"
        End If

        GetWindowIdentification = TrimNull(sTitle)
    End Function

    Private Function TrimNull(startstr As String) As String
        Dim pos As Integer
        pos = InStr(startstr, Chr$(0))
        If pos Then
            TrimNull = Left(startstr, pos - 1)
            Exit Function
        End If

        'if this far, there was
        'no Chr$(0), so return the string
        TrimNull = startstr
    End Function

    Private Function KillNewMailIcon(ByVal hwnd As Long) As Boolean
        Dim pShell_Notify As NOTIFYICONDATA
        Dim hResult As Long

        'setup the Shell_Notify structure
        pShell_Notify.cbSize = Len(pShell_Notify)
        pShell_Notify.hwnd = hwnd
        pShell_Notify.uID = 0

        ' Remove it from the system tray and catch result
        hResult = Shell_NotifyIcon(NIM_DELETE, pShell_Notify)
        If (hResult) Then
            KillNewMailIcon = True
        Else
            KillNewMailIcon = False
        End If
    End Function
End Module
