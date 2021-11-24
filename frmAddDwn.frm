VERSION 5.00
Begin VB.Form frmAddDwn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Download"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6255
   Icon            =   "frmAddDwn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
   Tag             =   "5"
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Tag             =   "8"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCnl 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Tag             =   "9"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label lblShow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URL :"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Tag             =   "6"
      Top             =   480
      Width           =   450
   End
   Begin VB.Label lblShow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Path :"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Tag             =   "7"
      Top             =   1200
      Width           =   480
   End
End
Attribute VB_Name = "frmAddDwn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
LoadResStrings Me
End Sub

Private Sub txtURL_Change()
With txtURL
If .Text <> "" And Left(.Text, 7) = "http://" And InStr(.Text, " ") = 0 Then
txtFile.Text = DownloadUrlToName(.Text)
End If
End With
End Sub

Private Sub cmdOK_Click()
If txtURL.Text = "" Or txtFile.Text = "" Then Beep: Exit Sub
frmMain.AddNewDwn txtURL.Text, txtFile.Text
cmdCnl_Click
End Sub

Private Sub cmdCnl_Click()
txtURL.Text = ""
txtFile.Text = ""
Unload Me
End Sub
