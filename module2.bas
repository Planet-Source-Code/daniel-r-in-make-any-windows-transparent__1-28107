Attribute VB_Name = "Module2"
Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Const SW_Minimize = 6
Public Const SW_Maximize = 3
Public Const SW_Normal = 1

Public TargetList As ListBox

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
  Dim SLength As Long, Buffer As String
  Dim RetVal As Long
  Static WinNum As Integer

  WinNum = WinNum + 1
  SLength = GetWindowTextLength(hWnd) + 1
  If SLength > 1 Then
    Buffer = Space(SLength)
    RetVal = GetWindowText(hWnd, Buffer, SLength)
    TargetList.AddItem Left(Buffer, SLength - 1)
    TargetList.ItemData(TargetList.NewIndex) = hWnd
  End If

  EnumWindowsProc = 1

End Function


