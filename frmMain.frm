VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Skykee"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   7335
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Tag             =   "1"
   Begin VB.PictureBox picBar 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   860
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Image imgAbout 
         Height          =   855
         Left            =   6360
         Picture         =   "frmMain.frx":4781A
         ToolTipText     =   "4"
         Top             =   0
         Width           =   420
      End
      Begin VB.Image imgAll 
         Height          =   855
         Left            =   1440
         Picture         =   "frmMain.frx":48B12
         ToolTipText     =   "3"
         Top             =   0
         Width           =   750
      End
      Begin VB.Image imgNew 
         Height          =   855
         Left            =   360
         Picture         =   "frmMain.frx":4AD2E
         ToolTipText     =   "2"
         Top             =   0
         Width           =   750
      End
      Begin VB.Image imgBg 
         Height          =   855
         Left            =   0
         Picture         =   "frmMain.frx":4CF4A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.VScrollBar sroDwn 
      Height          =   3135
      Left            =   6480
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin Skykee.Downloader Dlr 
      Left            =   0
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer tmrClp 
      Interval        =   1000
      Left            =   0
      Top             =   1560
   End
   Begin VB.PictureBox picLst 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   6735
      TabIndex        =   2
      Top             =   840
      Width           =   6735
      Begin Skykee.ucDwnUI DwnUI 
         Height          =   1080
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   1905
      End
      Begin VB.Image imgShdw 
         Height          =   300
         Left            =   0
         Picture         =   "frmMain.frx":4D070
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dlr_DownloadError(SaveFile As String)
If DwnUI.Count = 1 Then Exit Sub
On Error GoTo errRe
Dim i As Long
For i = 1 To DwnUI.UBound
If DwnUI(i).GetDwnPath = SaveFile Then
DwnUI(i).SetProgress 100, 100
DwnUI(i).SetRemText GetResString(11)
DoEvents
PlaySoundResource 102
End If
Next i
Exit Sub
errRe:
i = i + 1
Resume
End Sub

Private Sub Dlr_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
If DwnUI.Count = 1 Then Exit Sub
On Error GoTo errRe
Dim i As Long, RemBytes As Long
For i = 1 To DwnUI.UBound
If DwnUI(i).GetDwnPath = SaveFile Then
RemBytes = MaxBytes - CurBytes
DwnUI(i).SetProgress CurBytes, MaxBytes
If (CurBytes / MaxBytes) * 100 = 100 Then
DwnUI(i).SetProgress 100, 100
DwnUI(i).SetRemText GetResString(13)
DoEvents
PlaySoundResource 101
End If
End If
Next i
Exit Sub
errRe:
i = i + 1
Resume
End Sub

Private Sub DwnUI_DwnDeleted(Index As Integer)
If Dlr.CancelFileDownload(DwnUI(Index).GetDwnPath) Or Not DwnUI(Index).GetProgressVisible Then
Unload DwnUI(Index)
End If
ReAvgLst
End Sub

Private Sub Form_Load()
If App.PrevInstance Then End
LoadResStrings Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
With Me
imgBg.Width = Me.Width
imgAbout.Left = .ScaleWidth - imgAbout.Width - 360
picLst.Move 0, picBar.Height, .ScaleWidth
sroDwn.Move .ScaleWidth - sroDwn.Width, picBar.Height, sroDwn.Width, .ScaleHeight - picBar.Height
ReAvgLst
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
imgAll_Click
End Sub

Private Sub imgAbout_Click()
ShellAbout Me.hwnd, "MaxXSoft Skykee", "Copyright (C) 2013 MaxXSoft.", Me.Icon
End Sub

Private Sub imgAll_Click()
Dlr.CancelAllDownload
If DwnUI.Count = 1 Then Exit Sub
On Error GoTo errRe
Dim i As Long
For i = 1 To DwnUI.UBound
Unload DwnUI(i)
Next i
ReAvgLst
Exit Sub
errRe:
i = i + 1
Resume
End Sub

Private Sub imgNew_Click()
With frmAddDwn
Dim sClp As String
sClp = Clipboard.GetText
If IsDownUrl(sClp) Then
.txtURL.Text = Clipboard.GetText
End If
On Error Resume Next
.Show 1
End With
End Sub

Private Sub sroDwn_Change()
picLst.Top = picBar.Height - sroDwn.Value
End Sub

Private Sub sroDwn_Scroll()
sroDwn_Change
End Sub

Private Sub tmrClp_Timer()
On Error Resume Next
Static sLastClp As String
If Clipboard.GetText <> sLastClp Then
Dim sClp As String
sClp = Clipboard.GetText
If IsDownUrl(sClp) Then
Me.ZOrder 0
Me.SetFocus
imgNew_Click
End If
End If
sLastClp = Clipboard.GetText
End Sub

Sub AddNewDwn(sUrl As String, sPath As String)
Load DwnUI(DwnUI.UBound + 1)
ReAvgLst
DwnUI(DwnUI.UBound).Visible = True
DwnUI(DwnUI.UBound).SetDwnPath sPath
Dlr.BeginDownload sUrl, sPath
End Sub

Private Sub ReAvgLst()
On Error GoTo errRe
Dim i As Long, lTop As Long
lTop = 0
For i = 1 To DwnUI.UBound
DwnUI(i).Top = lTop
DwnUI(i).Width = picLst.ScaleWidth
lTop = lTop + DwnUI(0).Height
Next i
picLst.Height = (DwnUI.Count - 1) * DwnUI(0).Height + imgShdw.Height
imgShdw.Move 0, (DwnUI.Count - 1) * DwnUI(0).Height, picLst.Width
If picLst.Height > Me.ScaleHeight - picBar.Height Then
sroDwn.LargeChange = DwnUI(0).Height * 2
sroDwn.SmallChange = DwnUI(0).Height
sroDwn.Max = picLst.Height - (Me.ScaleHeight - picBar.Height)
sroDwn.Visible = True
Else
picLst.Top = picBar.Height
sroDwn.Visible = False
End If
Exit Sub
errRe:
i = i + 1
Resume
End Sub

