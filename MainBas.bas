Attribute VB_Name = "MainBas"
Option Explicit

'=================================Sleep========================================
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Savetime As Double

Public Sub Sleep(n As Long)
    Savetime = timeGetTime
    While timeGetTime < Savetime + n
        DoEvents
    Wend
End Sub
