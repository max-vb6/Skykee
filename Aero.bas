Attribute VB_Name = "Aero"
Option Explicit

Public Type MARGINS
m_Left As Long
m_Right As Long
m_Top As Long
m_Button As Long
End Type

Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000

Public Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, margin As MARGINS) As Long
Public Declare Function SetLayeredWindowAttributesByColor Lib "user32" Alias "SetLayeredWindowAttributes" (ByVal hWnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Sub SetAeroForm(hWnd As Long, lColorKey As Long, lTop As Long)
On Error GoTo errH
Dim mg As MARGINS
SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
SetLayeredWindowAttributesByColor hWnd, lColorKey, 0, LWA_COLORKEY
mg.m_Left = 0
mg.m_Button = 0
mg.m_Right = 0
mg.m_Top = lTop
DwmExtendFrameIntoClientArea hWnd, mg
Exit Sub
errH:
End Sub

Function CheckAero() As Boolean
CheckAero = (Dir(Environ("SYSTEMROOT") & "\system32\dwmapi.dll") <> "")
End Function
