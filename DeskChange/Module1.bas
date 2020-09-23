Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long

Public Const GWL_STYLE = (-16)
Public Const pmTitle = "Program Manager"                ' Program manager window Title
Public Const pmClass = "Progman"                        ' Program manager class name
Public Const dvClass = "SHELLDLL_DefView"               ' Shell class name
Public Const lvClass = "SysListView32"                  ' ListView class name

Public Const WM_PAINT = &HF&
Public Const WM_STYLECHANGED = &H7D&

Public Const LVM_FIRST = &H1000&                        ' ListView messages
Public Const LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
Public Const CLR_NONE = &HFFFFFFFF

Public g_TEXTBKCOLOR As Long

Public Enum ListViewStyles
    LVS_ICON = &H0&
    LVS_REPORT = &H1&
    LVS_SMALLICON = &H2&
    LVS_LIST = &H3&
End Enum

Public Function GetExplorerListViewHwnd() As Long
    Dim hwnd As Long                                            ' hwnd var.
    hwnd = GetDesktopWindow                                     ' Get hwnd of Desktop
    hwnd = FindWindowEx(hwnd, 0&, pmClass, pmTitle)             ' Get hwnd of Program Manager
    hwnd = FindWindowEx(hwnd, 0&, dvClass, "")                  ' Get hwnd of the shell default view
    hwnd = FindWindowEx(hwnd, 0&, lvClass, "")                  ' Get hwnd of ListView.
    GetExplorerListViewHwnd = hwnd                              ' Return hwnd of listview
End Function

Public Sub TextBackGroundTransparent(TransParent As Boolean)
    Dim hwnd As Long
    hwnd = GetExplorerListViewHwnd()                            ' Get Explorer Desktop ListBox
    If (g_TEXTBKCOLOR = 0) Then                                 ' Get Original Text background color
        g_TEXTBKCOLOR = SendMessage(hwnd, LVM_GETTEXTBKCOLOR, 0&, ByVal 0&)
    End If
    If TransParent Then
        SendMessage hwnd, LVM_SETTEXTBKCOLOR, 0&, ByVal CLR_NONE      ' Set Text background color to clear
    Else
        SendMessage hwnd, LVM_SETTEXTBKCOLOR, 0&, ByVal g_TEXTBKCOLOR ' Set Text background color to original color...
    End If
    InvalidateRect hwnd, 0&, 1&                                 ' Invalidate rect so a paint will occure on the entire window
    SendMessage hwnd, WM_PAINT, 0&, ByVal 0&                    ' Paint window(background listview)
End Sub

Public Sub ChangeExplorerListViewStyle(Style As ListViewStyles)
    Dim Rc As Long                                              ' API Return Code
    Dim hwnd As Long                                            ' Hwnd of listview
    Dim wStyle(1) As Long                                       ' Style bits of listview
    hwnd = GetExplorerListViewHwnd()                            ' Get Explorer Desktop Listview
    wStyle(0) = GetWindowLong(hwnd, GWL_STYLE)                  ' Get current style of listview
    wStyle(1) = (wStyle(0) \ 4) * 4                             ' Clear out the ListView Styles bits [last 2 bits]
    wStyle(1) = wStyle(1) Or Style                              ' Add new style bits...
    Rc = SendMessage(hwnd, WM_STYLECHANGED, GWL_STYLE, ByVal VarPtr(wStyle(0)))  ' Set new style bits in ListView.
End Sub
