Attribute VB_Name = "ServerControl"
Option Explicit

Public Sub StartServer()

    'Variable to store the result of functions
    Dim r As Long

    With frmMain
        'Tell the server object which port to listen on.
        .FTPServer.Port = 21
        
        'Give the server object a hWnd of your main form.
        .FTPServer.hWndToMsg = .hWnd
    
        'Total max clients
        .FTPServer.MaxClients = 5
    
        'Start the FTP server.
        r = .FTPServer.StartServer()

        If r <> 0 Then  'Problem starting server
            MsgBox .FTPServer.ServerGetErrorDescription(r), vbCritical
        End If
    End With

End Sub

Public Sub StopServer()

    frmMain.FTPServer.ShutdownServer

End Sub
