Attribute VB_Name = "Main"
Option Explicit
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Const MAX_PATH = 260
Public Const ILD_TRANSPARENT = &H1
Public Type SHFILEINFO
hIcon As Long
iIcon As Long
dwAttributes As Long
szDisplayName As String * MAX_PATH
szTypeName As String * 80
End Type
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal X As Long, ByVal Y As Long, ByVal Flags As Long) As Long
Public shinfo As SHFILEINFO
Public Const SHGFI_USEFILEATTRIBUTES = &H10
Public Const SHGFI_ICON = &H100

Public Function GetFileIcon(fName As String, sPicture As PictureBox)
Dim r As Long, hImgLarge As Long
hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), SHGFI_LARGEICON Or BASIC_SHGFI_FLAGS Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES)
sPicture.Picture = LoadPicture()
sPicture.AutoRedraw = True
r = ImageList_Draw(hImgLarge&, shinfo.iIcon, sPicture.hDC, 0, 0, ILD_TRANSPARENT)
Set sPicture.Picture = sPicture.Image
GetFileIcon = r
End Function

Public Function GetFileName(Path As String, Optional GetEx As Boolean) As String
On Error GoTo FileErr
Dim tstrs() As String
tstrs = Split(Path, "\")
If GetEx Then GetFileName = tstrs(UBound(tstrs)): Exit Function
tstrs = Split(tstrs(UBound(tstrs)), ".")
GetFileName = tstrs(0)
Exit Function
FileErr:
GetFileName = ""
End Function

Public Function DownloadUrlToName(URL As String) As String
On Error GoTo errH
Dim a As Integer, tmplen As Integer
Dim sTmp() As String, sFile As String
sTmp = Split(URL, "/")
sFile = sTmp(UBound(sTmp))
If InStr(sFile, "?") <> 0 Then
sTmp = Split(sFile, "?")
sFile = sTmp(0)
End If
DownloadUrlToName = sFile
Exit Function
errH:
DownloadUrlToName = ""
End Function

Public Function NumToByte(lByt As Long, Optional lLen As Long) As String
If lByt < 2 ^ 20 Then
NumToByte = Round(lByt / 2 ^ 10, lLen) & " KB"
Else
NumToByte = Round(lByt / 2 ^ 20, lLen) & " MB"
End If
End Function

Public Function IsDownUrl(sUrl As String) As Boolean
If sUrl <> "" And Left(sUrl, 7) = "http://" And _
InStr(sUrl, " ") = 0 And Right(sUrl, 1) <> "/" Then
IsDownUrl = True
End If
End Function

Sub LoadResStrings(frm As Form)
    On Error Resume Next

    Dim ctl As Control
    Dim obj As Object
    Dim sCtlType As String
    Dim nVal As Integer
    
    '设置窗体的 caption 属性
    frm.Caption = GetResString(CInt(frm.Tag))
      
    '设置控件的标题，对菜单项使用 caption 属性并对所有其他控件使用 Tag 属性
    For Each ctl In frm.Controls
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = GetResString(CInt(ctl.Tag))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = GetResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = GetResString(CInt(obj.Tag))
                obj.ToolTipText = GetResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = GetResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = GetResString(CInt(obj.Tag))
            Next
        Else
            nVal = 0
            nVal = Val(ctl.Tag)
            If nVal > 0 Then ctl.Caption = GetResString(nVal)
            nVal = 0
            nVal = Val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = GetResString(nVal)
        End If
    Next

End Sub

Function GetResString(ByVal id As Long) As String
On Error Resume Next
Dim lLag As Long
Dim lOffset As Long
lLag = GetSystemDefaultLCID
Select Case lLag
Case 2052             '简体
lOffset = 13
Case 1028             '繁体
lOffset = 26
Case 1041             '日语
lOffset = 39
Case Else             '英文
lOffset = 0
End Select
GetResString = LoadResString(id + lOffset)
End Function
