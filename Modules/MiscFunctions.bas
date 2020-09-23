Attribute VB_Name = "MiscFunctions"
Option Explicit

Public Sub writeToLogWindow(strString As String, Optional TimeStamp As Boolean)

    Dim strTimeStamp As String

    If TimeStamp = True Then strTimeStamp = "[" & Now & "] "
    frmMain.txtSvrLog.Text = frmMain.txtSvrLog.Text & vbCrLf & strTimeStamp & strString
    frmMain.txtSvrLog.SelStart = Len(frmMain.txtSvrLog.Text)

End Sub

Public Function StripNulls(strString As Variant) As String

    If InStr(strString, vbNullChar) Then
        StripNulls = Left(strString, InStr(strString, vbNullChar) - 1)
    Else
        StripNulls = strString
    End If

End Function

Public Function getAddress(ByVal lngAddress As Long) As Long

    getAddress = lngAddress

End Function
