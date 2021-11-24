VERSION 5.00
Begin VB.UserControl Downloader 
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Downloader.ctx":0000
   ScaleHeight     =   2385
   ScaleWidth      =   3480
End
Attribute VB_Name = "Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************
'我为人人
'人人为我
'枕善居汉化收藏整理
'http://www.mndsoft.com/blog/
'e-mail:mnd@mndsoft.com
'2005.03.06
'******************************************************

Option Explicit

Event DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
Event DownloadError(SaveFile As String)
Event DownloadComplete(MaxBytes As Long, SaveFile As String)
Event DownloadAllComplete(FileNotDownload() As String)

Private AsyncPropertyName() As String
Private AsyncStatusCode() As Byte

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)

    On Error Resume Next

        If AsyncProp.BytesMax <> 0 Then
            RaiseEvent DownloadProgress(CLng(AsyncProp.BytesRead), CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
        End If

        Select Case AsyncProp.StatusCode
          Case vbAsyncStatusCodeSendingRequest
            Debug.Print "Attempting to connect", AsyncProp.Target
          Case vbAsyncStatusCodeConnecting
            Debug.Print "Connecting", AsyncProp.Status '显示模板IP
          Case vbAsyncStatusCodeBeginDownloadData
            Debug.Print "Begin downloading", AsyncProp.Status '显示临时保存文件路径
            'Case vbAsyncStatusCodeDownloadingData
            '  Debug.Print "Downloading", AsyncProp.Status '显示目标 URL
          Case vbAsyncStatusCodeRedirecting
            Debug.Print "Redirecting", AsyncProp.Status
          Case vbAsyncStatusCodeEndDownloadData
            Debug.Print "Download complete", AsyncProp.Status
          Case vbAsyncStatusCodeError
            Debug.Print "Error...aborting transfer", AsyncProp.Status
            CancelAsyncRead AsyncProp.PropertyName
        End Select

End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)

  Dim f() As Byte, fn As Long
  Dim i As Integer

    On Error Resume Next

        Select Case AsyncProp.StatusCode
          Case vbAsyncStatusCodeEndDownloadData
            fn = FreeFile
            f = AsyncProp.Value
            Debug.Print "Writting to file " & AsyncProp.PropertyName
            Open AsyncProp.PropertyName For Binary Access Write As #fn
            Put #fn, , f
            Close #fn

            RaiseEvent DownloadComplete(CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)

          Case vbAsyncStatusCodeError
            CancelAsyncRead AsyncProp.PropertyName
            RaiseEvent DownloadError(AsyncProp.PropertyName)
        End Select

        For i = 1 To UBound(AsyncPropertyName)
            If AsyncPropertyName(i) = AsyncProp.PropertyName Then
                AsyncStatusCode(i) = AsyncProp.StatusCode
                Exit For
            End If
        Next i

        CheckAllDownloadComplete

End Sub

Private Sub UserControl_Initialize()

    SizeIt
    ReDim AsyncPropertyName(0)
    ReDim AsyncStatusCode(0)

End Sub

Private Sub UserControl_Resize()

    SizeIt

End Sub

Private Sub UserControl_Terminate()

    If UBound(AsyncPropertyName) > 0 Then CancelAllDownload

End Sub

Private Sub SizeIt()

    On Error GoTo ErrorSizeIt
    With UserControl
        .Width = ScaleX(32, vbPixels, vbTwips)
        .Height = ScaleY(32, vbPixels, vbTwips)
    End With

Exit Sub

ErrorSizeIt:
    MsgBox Err & ":错误在调用　SizeIt()." _
           & vbCrLf & vbCrLf & "错误描述: " & Err.Description, vbCritical, "错误"

Exit Sub

End Sub

Public Sub BeginDownload(URL As String, SaveFile As String, Optional AsyncReadOptions = vbAsyncReadForceUpdate)

    On Error GoTo ErrorBeginDownload
    UserControl.AsyncRead URL, vbAsyncTypeByteArray, SaveFile, AsyncReadOptions

    ReDim Preserve AsyncPropertyName(UBound(AsyncPropertyName) + 1)
    AsyncPropertyName(UBound(AsyncPropertyName)) = SaveFile
    ReDim Preserve AsyncStatusCode(UBound(AsyncStatusCode) + 1)
    AsyncStatusCode(UBound(AsyncStatusCode)) = 255

Exit Sub

ErrorBeginDownload:
    MsgBox Err & ":错误在调用 BeginDownload()." _
           & vbCrLf & vbCrLf & "错误描述: " & Err.Description, vbCritical, "错误"

Exit Sub

End Sub

Public Function CancelAllDownload() As Boolean

  Dim i As Integer

    On Error Resume Next

        For i = 1 To UBound(AsyncPropertyName)
            CancelAsyncRead AsyncPropertyName(i)
            Debug.Print "Killing download " & AsyncPropertyName(i)
        Next i

        ReDim AsyncPropertyName(0)
        ReDim AsyncStatusCode(0)

        CancelAllDownload = True

End Function

Public Function CancelFileDownload(SaveFile As String) As Boolean
  
  On Error GoTo CnlErr
  
  CancelAsyncRead (SaveFile)
  CancelFileDownload = True
  
  Exit Function
  
CnlErr:
  CancelFileDownload = False
  
End Function

Private Function CheckAllDownloadComplete()

  Dim i As Integer
  Dim FileNotDownload() As String
  Dim AllDownloadComplete As Boolean

    ReDim FileNotDownload(0)

    AllDownloadComplete = True
    For i = 1 To UBound(AsyncStatusCode)
        If AsyncStatusCode(i) = vbAsyncStatusCodeError Then
            ReDim Preserve FileNotDownload(UBound(FileNotDownload) + 1)
            FileNotDownload(UBound(FileNotDownload)) = AsyncPropertyName(i)
          ElseIf AsyncStatusCode(i) <> vbAsyncStatusCodeEndDownloadData Then
            AllDownloadComplete = False
            Exit For
        End If
    Next i

    If AllDownloadComplete Then
        CancelAllDownload
        RaiseEvent DownloadAllComplete(FileNotDownload)
    End If

End Function
