VERSION 5.00
Begin VB.UserControl ucDwnUI 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   ScaleHeight     =   1080
   ScaleWidth      =   6030
   Begin VB.Timer tmrSpd 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picPro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   960
      ScaleHeight     =   135
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   720
      Width           =   1215
      Begin VB.Label lblPro 
         BackColor       =   &H00008000&
         Height          =   255
         Left            =   -10
         TabIndex        =   4
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox picOpt 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   5280
      ScaleHeight     =   1095
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   0
      Width           =   735
      Begin VB.Image imgDel 
         Height          =   375
         Left            =   0
         Picture         =   "ucDwnUI.ctx":0000
         ToolTipText     =   "Delete download."
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblRem 
      BackStyle       =   0  'Transparent
      Caption         =   "0 KB/s  Total: 0 KB"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   720
      Width           =   2010
   End
   Begin VB.Label lblProg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00%"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   525
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FileName"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "ucDwnUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lByt As Long, nByt As Long, nMax As Long

Event DwnDeleted()

Sub SetDwnPath(sPath As String)
lblFile.ToolTipText = sPath
lblFile.Caption = GetFileName(sPath, True)
GetFileIcon sPath, picIcon
tmrSpd.Enabled = True
End Sub

Sub SetProgress(lNow As Long, lMax As Long)
On Error Resume Next
lblPro.Width = picPro.Width * (lNow / lMax)
lblProg.Caption = Format(lNow / lMax, "#0.00%")
nByt = lNow
nMax = lMax
picPro.Visible = (lblProg.Caption <> "100.00%")
lblProg.Visible = picPro.Visible
tmrSpd.Enabled = picPro.Visible
End Sub

Sub SetRemText(sRem As String)
lblRem.Caption = sRem
End Sub

Function GetDwnPath() As String
GetDwnPath = lblFile.ToolTipText
End Function

Function GetProgressVisible() As Boolean
GetProgressVisible = picPro.Visible
End Function

Private Sub imgDel_Click()
RaiseEvent DwnDeleted
End Sub

Private Sub picIcon_DblClick()
Dim sTmp As String
If lblFile.ToolTipText = "" Then Exit Sub
If InStr(Left(lblFile.ToolTipText, 3), ":\") = 0 Then
sTmp = App.Path & "\" & lblFile.ToolTipText
Else
sTmp = lblFile.ToolTipText
End If
If Dir(sTmp) = "" Then Exit Sub
ShellExecute 0, "open", sTmp, "", "", 1
End Sub

Private Sub tmrSpd_Timer()
lblRem.Caption = NumToByte(nByt - lByt) & "/s " & GetResString(12) & NumToByte(nMax)
lByt = nByt
End Sub

Private Sub UserControl_Initialize()
imgDel.ToolTipText = GetResString(10)
lByt = 0
tmrSpd_Timer
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
With UserControl
.Height = 1080
picOpt.Left = .ScaleWidth - picOpt.Width
lblRem.Left = .ScaleWidth - picOpt.Width - 240 - lblRem.Width
lblProg.Left = lblRem.Left
Dim lTmpPro As Long
lTmpPro = (lblPro.Width - 10) / picPro.Width
picPro.Width = .ScaleWidth - picPro.Left - picOpt.Width - lblRem.Width - 360
lblPro.Width = picPro * lTmpPro
End With
End Sub
