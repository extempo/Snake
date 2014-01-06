'Snake v1.0, 2014-01-02
'By Matt Carleton

'Module "API" - interactions with the OS


Option Explicit

'Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

