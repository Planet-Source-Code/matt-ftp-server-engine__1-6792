VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''
'Server configuration
'''''''''''''''''''''''''''''''''''''''''''
'Port server listens on
Public Port As Long

'Allows server admin to set max connections,
'must be MAX_CONNECTIONS <= the server constant
Public MaxClients As Integer

'Specifies the packet size (in bytes) to send data
'Shouldn't really be over 4096 bytes (4K)
Public TransferBufferSize As Long

'Every time a new clients connectes ClientCounter is
'incremented by 1, that new number becomes the new
'clients unique ID number.
Public ClientCounter As Long

'Every time a new client connects ConnectedClients is
'incremented by 1, every time a client disconnects,
'the number is decremented by 1.
Public ConnectedClients As Long

'''''''''''''''''''''''''''''''''''''''''''
'Server types
'''''''''''''''''''''''''''''''''''''''''''

'The reason for this being here instead of hidden in the server class
'module is that well, I found it next to impossible to pass
'a user defined type from the class module out.
'So this will make it easy for you to access client data at anytime.

'Client actions
Private Enum ClientStatus

    'CLIENT STATUS CONSTANTS
    stat_IDLE = 0
    stat_LOGGING_IN = 1
    stat_GETTING_DIR_LIST = 2
    stat_UPLOADING = 3
    stat_DOWNLOADING = 4

End Enum

Private Type ftpClient

    inUse As Boolean                'Identifies if this slot is being used.
    ID As Long                      'Unique number to identify a client.
    UserName As String              'User name client is is logged in as.
    IPAddress As String             'IP address of the client.
    SocketHandle As Long            'Reference number to the socket (Command socket) the client is connected to.
    DataSockHandle As Long          'Reference number to the socket (Data socket) to send/receive data.
    DataPort As Long                'Port number open on the client for the server to connect to.
    ConnectedAt As String           'Time the client first connected.
    IdleSince As String             'Last recorded time the client sent a command to the server.
    TotalBytesUploaded As Long      'Total bytes uploaded by client from the current session.
    TotalBytesDownloaded As Long    'Total bytes downloaded by client from the current session.
    TotalFilesUploaded As Long      'Total files uploaded by client from the current session.
    TotalFilesDownloaded As Long    'Total files downloaded by client from the current session.
    CurrentFile As String           'Current file being transfer, if any.
    cTotalBytesXfer As Long         'Total bytes of the current file that has been transfered.
    fFile As Long                   'Reference number to an open file on the server, if any.
    ConnectMode As Long             'If the client uses PASV mode or not.
    HomeDir As String               'Initial directory client starts in when they first connect.
    CurrentDir As String            'Current directory.
    DataToSend As String            'Store data here until it is set to the client.
    Status As ClientStatus          'What the client is currently doing.

End Type

'''''''''''''''''''''''''''''''''''''''''''
'Server constants
'''''''''''''''''''''''''''''''''''''''''''
'900 seconds (15 mins.) of idle time will disconnect the client
'without regard to specific user settings
Private Const MAX_IDLE_TIME = 900

'500 simultaneous connections for the server
Private Const MAX_CONNECTIONS = 500

'CONNECTION MODE CONSTANTS
Private Const cMode_NORMAL = 0
Private Const cMode_PASV = 1

'Array that holds client information for every client.
Private client(MAX_CONNECTIONS) As ftpClient

'''''''''''''''''''''''''''''''''''''''''''
'Server events
'''''''''''''''''''''''''''''''''''''''''''
'ServerErrorOccurred() event fires when an error occurs
'in a sub/function that wasnt called outside of the object.
Public Event ServerErrorOccurred(ByVal errNumber As Long)

'ServerStarted() event fires after successful startup of Winsock
'and the specified port is successfully opened, server is ready to accept clients.
Public Event ServerStarted()

'ServerStopped() event fires after all connected clients
'have been disconnected, Winsock is shutdown and other
'misc. variables are reset.
Public Event ServerStopped()

'NewClient() event fires when a new connection is accepted.
Public Event NewClient(ByVal ClientID As Long)

'ClientLoggedIn() event fires when a connected client
'successfully logs in with a correct user name and password.
Public Event ClientLoggedIn(ByVal ClientID As Long)

'ClientLoggedOut() event fires when a client
'disconnects or is disconnected by the server.
'When this event fires, the information on that
'client will no longer be available.
Public Event ClientLoggedOut(ByVal ClientID As Long)

'ClientSentCommand() event fires when a client
'sends a FTP command to the server.
Public Event ClientSentCommand(ByVal ClientID As Long, Command As String, Args As String)

'ClientStatusChanged() event fires when a clients
'status has changed.  This event will most likely
'follow after another.  eg. client starts a file
'transfer...
Public Event ClientStatusChanged(ByVal ClientID As Long)


'================================= MISC APIS/TYPES =================================


'These three functions are for generating a list of files and directories.
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA _
    ) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA _
    ) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
    lpFileTime As FILETIME, _
    lpSystemTime As SYSTEMTIME _
    ) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Function FindClose Lib "kernel32" ( _
    ByVal hFindFile As Long _
    ) As Long

Private Const MAX_PATH = 260

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Long
End Type


'================================= WINSOCK =================================


'''''''''''''''''''''''''''''''''''''''''''
'Winsock Generic Variables
'''''''''''''''''''''''''''''''''''''''''''
Private WinsockInfo As WSADataType

'''''''''''''''''''''''''''''''''''''''''''
'Winsock Constants
'''''''''''''''''''''''''''''''''''''''''''
Private Const FD_SETSIZE = 64
Private Const hostent_size = 16
Private Const servent_size = 14

Private Const protoent_size = 10

Private Const IPPROTO_TCP = 6
Private Const IPPROTO_UDP = 17

Private Const INADDR_NONE = &HFFFF
Private Const INADDR_ANY = &H0

Private Const sockaddr_size = 16
Private saZero As sockaddr

Private Const WSA_DESCRIPTIONLEN = 256
Private Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1

Private Const WSA_SYS_STATUS_LEN = 128
Private Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Private Const INVALID_SOCKET = -1
Private Const SOCKET_ERROR = -1

Private Const SOCK_STREAM = 1
Private Const SOCK_DGRAM = 2

Private Const MAXGETHOSTSTRUCT = 1024

Private Const AF_INET = 2
Private Const PF_INET = 2

'---async notification constants
Private Const SOL_SOCKET = &HFFFF&
Private Const SO_LINGER = &H80&
Private Const FD_READ = &H1&
Private Const FD_WRITE = &H2&
Private Const FD_ACCEPT = &H8&
Private Const FD_CONNECT = &H10&
Private Const FD_CLOSE = &H20&

Private OldWndProc As Long
Public hWndToMsg As Long
Private Const GWL_WNDPROC = (-4)

'''''''''''''''''''''''''''''''''''''''''''
'Winsock Types
'''''''''''''''''''''''''''''''''''''''''''
Private Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Integer
End Type

Private Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Private Type Inet_address
    Byte4 As String * 1
    Byte3 As String * 1
    Byte2 As String * 1
    Byte1 As String * 1
End Type

Private Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type

Private Type protoent
    p_name As Long
    p_aliases As Long
    p_proto As Integer
End Type

Private Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Private Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

Private Type IP_OPTION_INFORMATION
    TTL As Byte                   ' Time to Live (used for traceroute)
    Tos As Byte                   ' Type of Service (usually 0)
    FLAGS As Byte                 ' IP header Flags (usually 0)
    OptionsSize As Long           ' Size of Options data (usually 0, max 40)
    OptionsData As String * 128   ' Options data buffer
End Type

Private Type IP_ECHO_REPLY
    Address(0 To 3) As Byte           ' Replying Address
    Status As Long                    ' Reply Status
    RoundTripTime As Long             ' Round Trip Time in milliseconds
    DataSize As Integer               ' reply data size
    Reserved As Integer               ' for system use
    Data As Long                      ' pointer to echo data
    Options As IP_OPTION_INFORMATION  ' Reply Options
End Type

'''''''''''''''''''''''''''''''''''''''''''
'Winsock API Declarations
'''''''''''''''''''''''''''''''''''''''''''
'---Windows System Functions
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'---SOCKET FUNCTIONS
Private Declare Function accept Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, addrlen As Long) As Long
Private Declare Function bind Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
Private Declare Function closesocket Lib "wsock32.dll" (ByVal S As Long) As Long
Private Declare Function connect Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
Private Declare Function ioctlsocket Lib "wsock32.dll" (ByVal S As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function getpeername Lib "wsock32.dll" (ByVal S As Long, sname As sockaddr, namelen As Long) As Long
Private Declare Function getsockname Lib "wsock32.dll" (ByVal S As Long, sname As sockaddr, namelen As Long) As Long
Private Declare Function getsockopt Lib "wsock32.dll" (ByVal S As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
Private Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Private Declare Function listen Lib "wsock32.dll" (ByVal S As Long, ByVal backlog As Long) As Long
Private Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
Private Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
Private Declare Function recv Lib "wsock32.dll" (ByVal S As Long, ByVal buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
Private Declare Function send Lib "wsock32.dll" (ByVal S As Long, ByVal buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
Private Declare Function setsockopt Lib "wsock32.dll" (ByVal S As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
'---DATABASE FUNCTIONS
Private Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function getprotobynumber Lib "wsock32.dll" (ByVal proto As Integer) As Long
Private Declare Function getprotobyname Lib "wsock32.dll" (ByVal proto_name As String) As Long
Private Declare Function getservbyport Lib "wsock32.dll" (ByVal Port As Long, ByVal proto As String) As Long
Private Declare Function getservbyname Lib "wsock32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
'---WINDOWS EXTENSIONS
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Sub WSASetLastError Lib "wsock32.dll" (ByVal iError As Long)
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal S As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'---ICMP
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal ICMPHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean

'mWndProcOrg holds the original address of the
'Window Procedure for this window. This is used to
'route messages to the original procedure after you
'process them.
Private mWndProcOrg As Long

'Handle (hWnd) of the subclassed window.
Private mHWndSubClassed As Long

Private Sub SubClass()

    'I picked up most of this subclassing code from
    'MSDN, article Q179398 at http://msdn.microsoft.com

    '-------------------------------------------------------------
    'Initiates the subclassing of this Server object (hwnd).
    'Records the original WinProc of the window in mWndProcOrg.
    'Places a pointer to the object in the window's UserData area.
    '-------------------------------------------------------------
    
    'Exit if the window is already subclassed.
    If mWndProcOrg Then Exit Sub

    'Redirect the window's messages from this object's default
    'Window Procedure to the SubWndProc function in wSubClass.bas
    'module and record the address of the previous Window
    'Procedure for this window in mWndProcOrg.
    mWndProcOrg = SetWindowLong(hWndToMsg, GWL_WNDPROC, _
                                 AddressOf SubWndProc)
    
    'Record your window handle in case SetWindowLong gave you a
    'new one. You will need this handle so that you can unsubclass.
    mHWndSubClassed = hWndToMsg
    
    'Store a pointer to this object in the UserData section of
    'this window that will be used later to get the pointer to
    'the control based on the handle (hwnd) of the window getting
    'the message.
    Call SetWindowLong(hWndToMsg, GWL_USERDATA, ObjPtr(Me))

End Sub

Private Sub UnSubClass()
    
    '-----------------------------------------------------------
    'Unsubclasses this object's window (hwnd), setting the
    'address of the Windows Procedure back to the address it was
    'at before it was subclassed.
    '-----------------------------------------------------------
    
    'Ensures that you don't try to unsubclass the window when
    'it is not subclassed.
    If mWndProcOrg = 0 Then Exit Sub
    
    'Reset the window's function back to the original address.
    SetWindowLong mHWndSubClassed, GWL_WNDPROC, mWndProcOrg

    '0 Indicates that you are no longer subclassed.
    mWndProcOrg = 0

End Sub

Private Sub Class_Terminate()

    'This event runs when the object is cleared from memory.
   UnSubClass

End Sub

'''''''''''''''''''''''''''''''''''''''''''
'Functions to control Winsock
'''''''''''''''''''''''''''''''''''''''''''
Private Function StartWinsock() As Long

    If Not WSAStartup(257, WinsockInfo) Then
        'We have to do subclassing becuase VB does not support AddressOf of
        'functions/subs that reside in class modules.
        'I hear VB7 will allow this, in which case you should be able to remove
        'SubClass and uncomment the line below.
        'OldWndProc = SetWindowLong(hWndToMsg, GWL_WNDPROC, AddressOf WindowProc)
        SubClass
    Else
        'Winsock not working
        StartWinsock = WSAGetLastError()
    End If

End Function

Private Function EndWinsock() As Long
    
    EndWinsock = WSACleanup
    SetWindowLong hWndToMsg, GWL_WNDPROC, mWndProcOrg

End Function

Private Function ListenForConnect(ByVal Port As Long) As Long
    
    Dim S As Long, dummy As Long
    Dim SelectOps As Long
    Dim saZero As sockaddr
    Dim sockin As sockaddr
    Dim errors As Long

    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    
    If sockin.sin_port = INVALID_SOCKET Then
        Debug.Print "ERROR in sockin.sin_port"
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    
    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        errors = WSAGetLastError()
        Debug.Print "ERROR in htonl"
        ListenForConnect = errors
        Exit Function
    End If
    
    S = socket(PF_INET, SOCK_STREAM, 0)
    If S < 0 Then
        errors = WSAGetLastError()
        Debug.Print "ERROR in Socket()"
        ListenForConnect = errors
        Exit Function
    End If
    
    If bind(S, sockin, sockaddr_size) Then
        errors = WSAGetLastError()
        Debug.Print "ERROR in bind()"
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ListenForConnect = errors
        Exit Function
    End If
    
    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(S, hWndToMsg, ByVal 1025, ByVal SelectOps) Then
        errors = WSAGetLastError()
        Debug.Print "ERROR in WSAAsnycSelect"
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ListenForConnect = WSAGetLastError()
        Exit Function
    End If
    
    If listen(S, 1) Then
        errors = WSAGetLastError()
        Debug.Print "ERROR in listen()"
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ListenForConnect = errors
        Exit Function
    End If
    
    'ListenForConnect = S

End Function

Private Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
    
    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetSelectEvent = lParam And &HFFFF&
    End If

End Function

Private Function WSAGetAsyncError(ByVal lParam As Long) As Integer
    
    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000

End Function

Friend Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo HandleError

    If uMsg <> 1025 Then
        WindowProc = CallWindowProc(mWndProcOrg, hWnd, uMsg, wParam, ByVal lParam)
        Exit Function
    End If

    Dim WSAEvent As Long, WSAError As Long
    WSAEvent = WSAGetSelectEvent(lParam)
    WSAError = WSAGetAsyncError(lParam)

    Select Case WSAEvent ' Determine Event

        Case FD_CLOSE ' Socket Close
            'Debug.Print "Doing FD_CLOSE"

            If GetClientArrayLocBySockH(wParam) <> -1 Then LogoutClient client(GetClientArrayLocBySockH(wParam)).ID

        Case FD_CONNECT
            'Debug.Print "Doing FD_CONNECT"
            If WSAError = 0 Then ' No Error
                Debug.Print "Connected"
            Else
                ' Not Connected
                Dim Error As String
                Error = GetWSAErrorString$(WSAError&)
                Debug.Print Error
                closesocket wParam
            End If

        Case FD_ACCEPT
            'Debug.Print "Doing FD_ACCEPT"
            Dim Sock As sockaddr
            Dim newSock As Long
            newSock = accept(wParam, Sock, sockaddr_size)

            NewClient newSock, Sock

        Case FD_READ
            'Debug.Print "Doing FD_READ"
            Dim buflen As Long
            Dim raw_data As String
            raw_data = Space(TransferBufferSize)

            buflen = recv(wParam, raw_data, TransferBufferSize, 0)
            
            If buflen > -1 Then DataArrival ByVal wParam, Left(raw_data, buflen)

        Case FD_WRITE
            'Debug.Print "Doing FD_WRITE"

            Dim cArrayLoc As Integer

            cArrayLoc = GetClientArrayLocByDataSockH(wParam)

            If cArrayLoc > -1 Then
                SendResponse client(cArrayLoc).SocketHandle, "226 Transfer complete."
                closesocket wParam
            End If

    End Select

HandleError:
    If Err.Number <> 0 Then
        MsgBox "WindowProc error:" & vbCrLf & Err.Description
    End If

End Function

Private Function GetAscIp(ByVal inn As Long) As String
    
    On Error Resume Next
    
    Dim lpStr As Long
    Dim nStr As Long
    Dim retString As String
        retString = String(32, 0)
    
    lpStr = inet_ntoa(inn)

    If lpStr = 0 Then
        GetAscIp = ""
        Exit Function
    End If
    
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    GetAscIp = retString
    
    If Err Then GetAscIp = ""

End Function

Private Function GetHostByNameAlias(ByVal hostname As String) As Long

    On Error Resume Next
    'Return IP address as a long, in network byte order

    Dim phe As Long             'pointer to host information entry
    Dim heDestHost As HostEnt   'hostent structure
    Dim addrList As Long
    Dim retIP As Long

    'first check to see if what we have been passed is a valid IP
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        'it wasn't an IP, so do a DNS lookup
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            'Pointer is non-null, so copy in hostent structure
            MemCopy heDestHost, ByVal phe, hostent_size
            'Now get first pointer in address list
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
          Else
            'its not a valid address
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = Trim$(retIP)
    If Err Then GetHostByNameAlias = INADDR_NONE

End Function

Private Function GetHostByAddress(ByVal addr As Long) As String
    
    On Error Resume Next
    
    Dim phe As Long, ret As Long
    Dim heDestHost As HostEnt
    Dim hostname As String
    
    phe = gethostbyaddr(addr, 4, PF_INET)
    If phe <> 0 Then
        MemCopy heDestHost, ByVal phe, hostent_size
        hostname = String(256, 0)
        MemCopy ByVal hostname, ByVal heDestHost.h_name, 256
        GetHostByAddress = Left(hostname, InStr(hostname, Chr(0)) - 1)
    Else
        GetHostByAddress = ""
    End If
    
    If Err Then GetHostByAddress = ""

End Function

Private Function AddrToIP(ByVal AddrOrIP As String) As String
    
    AddrToIP$ = GetAscIp(GetHostByNameAlias(AddrOrIP$))

End Function

Private Function IpToAddr(ByVal AddrOrIP As String) As String

    On Error Resume Next
    
    IpToAddr = GetHostByAddress(GetHostByNameAlias(AddrOrIP$))
    If Err Then IpToAddr = ""

End Function

Private Function DNSLookup(ByVal Address As String) As String

    If inet_addr(Address) <> INADDR_NONE Then
        DNSLookup = IpToAddr(Address)
    ElseIf gethostbyname(Address) <> 0 Then DNSLookup = AddrToIP(Address)
    End If

End Function

Private Function GetLocalHostName() As String

    On Error Resume Next

    Dim dummy As Long
    Dim LocalName As StartUpPositionConstants
    Dim S As String

    LocalName = String(256, 0)
    LocalName = ""
    dummy = 1
    S = String(256, 0)
    dummy = gethostname(S, 256)
    If dummy = 0 Then
        S = Left(S, InStr(S, Chr(0)) - 1)
        If Len(S) > 0 Then
            LocalName = S
        End If
    End If
    
    GetLocalHostName = Trim$(LocalName)
    If Err Then GetLocalHostName = ""

End Function

Private Function LongIP(ByVal IPAddress As String) As String

    On Error GoTo longipError
    
    Dim nStr As Long
    Dim retString As String
    Dim lpStr As Long
    Dim inn As Long

    If InStr(IPAddress, ".") = 0 Then
        If Val(IPAddress) > 2147483647 Then
            inn = Val(IPAddress) - 4294967296#
        Else
            inn = Val(IPAddress)
        End If
        
        inn = ntohl(inn)
        retString = String(32, 0)
        lpStr = inet_ntoa(inn)
        
        If lpStr = 0 Then
            LongIP = "0.0.0.0"
            Exit Function
        End If
        
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        MemCopy ByVal retString, ByVal lpStr, nStr
        retString = Left(retString, nStr)
        LongIP = retString
        Exit Function

longipError:
        LongIP = ""
        Exit Function
        Resume
    Else
        inn = inet_addr(IPAddress)
        inn = htonl(inn)
        If inn < 0 Then
            LongIP = CVar(inn + 4294967296#)
        Else
            LongIP = CVar(inn)
        End If
    End If

End Function

Private Function ConnectSock(ByVal Host As String, ByVal Port As Long) As Long
    
    Dim S As Long, SelectOps As Long, dummy As Long
    Dim saZero As sockaddr
    Dim sockin As sockaddr

    sockin = saZero
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    
    If sockin.sin_port = INVALID_SOCKET Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If

    sockin.sin_addr = GetHostByNameAlias(Host)
    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If

    S = socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If S < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    
    'If SetSockLinger(S, 1, 0) = SOCKET_ERROR Then
    '    If S > 0 Then
    '        dummy = closesocket(S)
    '    End If
    '    ConnectSock = INVALID_SOCKET
    '    Exit Function
    'End If
     
     SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
     If WSAAsyncSelect(S, hWndToMsg, ByVal 1025, ByVal SelectOps) Then
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    
    If connect(S, sockin, sockaddr_size) <> -1 Then
        If S > 0 Then
            dummy = closesocket(S)
        End If
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
     
     ConnectSock = S

End Function

Private Function MakeDataConnection(cArrayLoc As Integer) As Long

    With client(cArrayLoc)
        .DataSockHandle = ConnectSock(.IPAddress, .DataPort)
    End With

End Function

Private Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long
    
    Dim Linger As LingerType
    
    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime
    
    If setsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
        Debug.Print "Error setting linger info: " & WSAGetLastError()
        SetSockLinger = SOCKET_ERROR
    Else
        If getsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
            Debug.Print "Error getting linger info: " & WSAGetLastError()
            SetSockLinger = SOCKET_ERROR
        Else
            Debug.Print "Linger is on if nonzero: "; Linger.l_onoff
            Debug.Print "Linger time if linger is on: "; Linger.l_linger
        End If
    End If

End Function

Private Function GetWSAErrorString(ByVal errnum&) As String
    
    On Error Resume Next
    
    Select Case errnum
        Case 10004: GetWSAErrorString = "Interrupted system call."
        Case 10009: GetWSAErrorString = "Bad file number."
        Case 10013: GetWSAErrorString = "Permission Denied."
        Case 10014: GetWSAErrorString = "Bad Address."
        Case 10022: GetWSAErrorString = "Invalid Argument."
        Case 10024: GetWSAErrorString = "Too many open files."
        Case 10035: GetWSAErrorString = "Operation would block."
        Case 10036: GetWSAErrorString = "Operation now in progress."
        Case 10037: GetWSAErrorString = "Operation already in progress."
        Case 10038: GetWSAErrorString = "Socket operation on nonsocket."
        Case 10039: GetWSAErrorString = "Destination address required."
        Case 10040: GetWSAErrorString = "Message too long."
        Case 10041: GetWSAErrorString = "Protocol wrong type for socket."
        Case 10042: GetWSAErrorString = "Protocol not available."
        Case 10043: GetWSAErrorString = "Protocol not supported."
        Case 10044: GetWSAErrorString = "Socket type not supported."
        Case 10045: GetWSAErrorString = "Operation not supported on socket."
        Case 10046: GetWSAErrorString = "Protocol family not supported."
        Case 10047: GetWSAErrorString = "Address family not supported by protocol family."
        Case 10048: GetWSAErrorString = "Address already in use."
        Case 10049: GetWSAErrorString = "Can't assign requested address."
        Case 10050: GetWSAErrorString = "Network is down."
        Case 10051: GetWSAErrorString = "Network is unreachable."
        Case 10052: GetWSAErrorString = "Network dropped connection."
        Case 10053: GetWSAErrorString = "Software caused connection abort."
        Case 10054: GetWSAErrorString = "Connection reset by peer."
        Case 10055: GetWSAErrorString = "No buffer space available."
        Case 10056: GetWSAErrorString = "Socket is already connected."
        Case 10057: GetWSAErrorString = "Socket is not connected."
        Case 10058: GetWSAErrorString = "Can't send after socket shutdown."
        Case 10059: GetWSAErrorString = "Too many references: can't splice."
        Case 10060: GetWSAErrorString = "Connection timed out."
        Case 10061: GetWSAErrorString = "Connection refused."
        Case 10062: GetWSAErrorString = "Too many levels of symbolic links."
        Case 10063: GetWSAErrorString = "File name too long."
        Case 10064: GetWSAErrorString = "Host is down."
        Case 10065: GetWSAErrorString = "No route to host."
        Case 10066: GetWSAErrorString = "Directory not empty."
        Case 10067: GetWSAErrorString = "Too many processes."
        Case 10068: GetWSAErrorString = "Too many users."
        Case 10069: GetWSAErrorString = "Disk quota exceeded."
        Case 10070: GetWSAErrorString = "Stale NFS file handle."
        Case 10071: GetWSAErrorString = "Too many levels of remote in path."
        Case 10091: GetWSAErrorString = "Network subsystem is unusable."
        Case 10092: GetWSAErrorString = "Winsock DLL cannot support this application."
        Case 10093: GetWSAErrorString = "Winsock not initialized."
        Case 10101: GetWSAErrorString = "Disconnect."
        Case 11001: GetWSAErrorString = "Host not found."
        Case 11002: GetWSAErrorString = "Nonauthoritative host not found."
        Case 11003: GetWSAErrorString = "Nonrecoverable error."
        Case 11004: GetWSAErrorString = "Valid name, no data record of requested type."
        Case Else:  GetWSAErrorString = "Unknown error " & errnum
    End Select

End Function

Private Function SendResponse(SocketHandle As Long, raw_data As String) As Long

    Dim Data As String
    Dim DataLen As Long

    Data = raw_data & vbCrLf
    DataLen = Len(Data)

    SendResponse = send(SocketHandle, Data, DataLen, 0)

End Function

Private Function SendData(cArrayLoc As Integer) As Long

    On Error Resume Next

    Dim Data As String

    With client(cArrayLoc)
        Debug.Print "COMPARING: " & Len(.DataToSend) & " > " & TransferBufferSize
        If Len(.DataToSend) > TransferBufferSize Then
            Data = Left$(.DataToSend, TransferBufferSize)
            .DataToSend = Right$(.DataToSend, Len(.DataToSend) - TransferBufferSize)
        Else
            Debug.Print "Sending last bit."

            Data = .DataToSend
            .DataToSend = ""
        End If

        send .DataSockHandle, Data, Len(Data), 0
    End With

HandleError:
    If Err.Number <> 0 Then
        MsgBox "SendData error:" & vbCrLf & Err.Description & vbCrLf & Err.Source
    End If

End Function

Private Function GetPeerHost(socket As Long) As String
    
    Dim Sock As sockaddr
    
    Call getpeername(socket, Sock, sockaddr_size)
    GetPeerHost = GetHostByAddress(Sock.sin_addr)

End Function

Private Function GetPeerPort(socket As Long) As Integer
    
    Dim Sock As sockaddr
    
    Call getpeername(socket, Sock, sockaddr_size)
    GetPeerPort = ntohs(Sock.sin_port)

End Function

Private Function GetProtocol(Protocol)
    
    Dim peDestProt As protoent
    Dim ppe As Long
    
    If TypeName(Protocol) = "String" Then
        ppe = getprotobyname(Protocol)
        If ppe = 0 Then
            GetProtocol = -1
            Exit Function
        End If
    
        MemCopy peDestProt, ByVal ppe, protoent_size
        GetProtocol = peDestProt.p_proto
    End If

    If IsNumeric(Protocol) Then
        ppe = getprotobynumber(CInt(Protocol))
        
        If ppe = 0 Then
            GetProtocol = -1
            Exit Function
        End If

        MemCopy peDestProt, ByVal ppe, protoent_size
        Dim strProto As String
            strProto = String$(256, 0)
        MemCopy ByVal strProto, ByVal peDestProt.p_name, 256
        strProto = Left$(strProto, InStr(strProto, Chr$(0)) - 1)
        GetProtocol = strProto
    End If

End Function

Private Function GetService(Service, Protocol)
    
    Dim seDestServ As servent
    Dim pse As Long
    
    If IsNumeric(Protocol) Then Protocol = GetProtocol(CInt(Protocol))
    If TypeName(Service) = "String" Then
        pse = getservbyname(Service, Protocol)
        
        If pse = 0 Then
            GetService = -1
            Exit Function
        End If

        MemCopy seDestServ, ByVal pse, servent_size
        GetService = ntohs(seDestServ.s_port)
    End If

    If IsNumeric(Service) Then
        pse = getservbyport(htons(CLng(Service)), Protocol)
        
        If pse = 0 Then
            GetService = -1
            Exit Function
        End If
    
        MemCopy seDestServ, ByVal pse, servent_size
        Dim strService As String
        strService = String$(256, 0)
        MemCopy ByVal strService, ByVal seDestServ.s_name, 256
        strService = Left$(strService, InStr(strService, Chr$(0)) - 1)
        GetService = strService
    End If

End Function

Private Function GetLocalHost(socket As Long) As String
    
    Dim Sock As sockaddr
    
    Call getsockname(socket, Sock, sockaddr_size)
    GetLocalHost = GetHostByAddress(Sock.sin_addr)

End Function

Private Function GetLocalPort(socket As Long) As Integer
    
    Dim Sock As sockaddr
    
    Call getsockname(socket, Sock, sockaddr_size)
    GetLocalPort = ntohs(Sock.sin_port)

End Function

Private Function PingHost(Host As String)

    Dim ICMPHandle As Long
    
    ' Get IP Address
    If inet_addr(Host) = INADDR_NONE Then
        Host = DNSLookup(Host)
        If Host = "" Then Exit Function
    End If

    ICMPHandle = IcmpCreateFile()
    If ICMPHandle = 0 Then Exit Function ' Error
    
    Dim IPOption As IP_OPTION_INFORMATION
    Dim buf As String
    Dim IPReply As IP_ECHO_REPLY
    
    IPOption.TTL = "255"
    buf = String$(32, 0) ' Intialize Buffer
    
    If IcmpSendEcho(ICMPHandle, GetHostByNameAlias(Host), buf, 32, IPOption, IPReply, Len(IPReply) + 8, 2700) = False Then PingHost = "Request Timed Out"
    If PingHost = "" Then PingHost = GetICMPError$(CInt(IPReply.Status))
    If PingHost = "" Then PingHost = IPReply.RoundTripTime
    
    Call IcmpCloseHandle(ICMPHandle)

End Function

Private Function TraceRt(Host As String)
    
    Dim ICMPHandle As Long
    
    ' Get IP Address
    If inet_addr(Host) = INADDR_NONE Then
        Host = DNSLookup(Host)
        If Host = "" Then Exit Function
    End If

    ICMPHandle = IcmpCreateFile()
    If ICMPHandle = 0 Then Exit Function ' Error
    
    Dim IPOption As IP_OPTION_INFORMATION
    Dim buf As String
    Dim IPReply As IP_ECHO_REPLY
    Dim HopAddress As String
    Dim HopCounter As Integer
    
    For IPOption.TTL = 2 To 30
        buf = String$(32, 0)
        If IcmpSendEcho(ICMPHandle, GetHostByNameAlias(Host), buf, 32, IPOption, IPReply, Len(IPReply) + 8, 2700) = False Then
            GoTo EndEcho
        End If

        If HopCounter = 0 Then TraceRt = "Hop #0" Else TraceRt = TraceRt & vbCrLf & "Hop #" & HopCounter
        HopAddress = CStr(IPReply.Address(0)) & "." & CStr(IPReply.Address(1)) & "." & CStr(IPReply.Address(2)) & "." & CStr(IPReply.Address(3))
        TraceRt = TraceRt & "  " & HopAddress
        If HopAddress = Host Then Exit For
        HopCounter = HopCounter + 1
EndEcho:
        If IPOption.TTL = 30 Then TraceRt = TraceRt & vbCrLf & "Trace Route to Long."
        DoEvents
    Next IPOption.TTL

End Function

Private Function GetICMPError(Error As Integer) As String
    
    Select Case Error
        Case 11001: GetICMPError = "Buffer too Small"
        Case 11002: GetICMPError = "Destination Network Not Reachable"
        Case 11003: GetICMPError = "Destination Host Not Reachable"
        Case 11004: GetICMPError = "Destination Protocol Not Reachable"
        Case 11005: GetICMPError = "Destination Port Not Reachable"
        Case 11006: GetICMPError = "No Resources Available"
        Case 11007: GetICMPError = "Bad Option"
        Case 11008: GetICMPError = "Hardware Error"
        Case 11009: GetICMPError = "Packet too Big"
        Case 11010: GetICMPError = "Request Timed Out"
        Case 11011: GetICMPError = "Bad Request"
        Case 11012: GetICMPError = "Bad Route"
        Case 11013: GetICMPError = "TTL Expired in Transit"
        Case 11014: GetICMPError = "TTL Expired Reassembly"
        Case 11015: GetICMPError = "Parameter Problem"
        Case 11016: GetICMPError = "Source Quench"
        Case 11017: GetICMPError = "Option too Big"
        Case 11018: GetICMPError = "Bad Destination"
        Case 11019: GetICMPError = "Address Deleted"
        Case 11020: GetICMPError = "Spec MTU Change"
        Case 11021: GetICMPError = "MTU Change"
        Case 11022: GetICMPError = "Unload"
        Case 11050: GetICMPError = "General Failure"
    End Select

End Function


'================================= SERVER =================================


'''''''''''''''''''''''''''''''''''''''''''
'Server public properties
'''''''''''''''''''''''''''''''''''''''''''
'Use these properties to get information about a specific client
'outside of this class module.
Public Property Get GetClientConnectedAt(ClientID As Long) As String
    GetClientConnectedAt = client(GetClientArrayLocByID(ClientID)).ConnectedAt
End Property

Public Property Get GetClientConnectMode(ClientID As Long) As String
    GetClientConnectMode = client(GetClientArrayLocByID(ClientID)).ConnectMode
End Property

Public Property Get GetClientcTotalBytesXfer(ClientID As Long) As Long
    GetClientcTotalBytesXfer = client(GetClientArrayLocByID(ClientID)).cTotalBytesXfer
End Property

Public Property Get GetClientCurrentDir(ClientID As Long) As String
    GetClientCurrentDir = client(GetClientArrayLocByID(ClientID)).CurrentDir
End Property

Public Property Get GetClientCurrentFile(ClientID As Long) As String
    GetClientCurrentFile = client(GetClientArrayLocByID(ClientID)).CurrentFile
End Property

Public Property Get GetClientDataPort(ClientID As Long) As Long
    GetClientDataPort = client(GetClientArrayLocByID(ClientID)).DataPort
End Property

Public Property Get GetClientDataSockHandle(ClientID As Long) As Long
    GetClientDataSockHandle = client(GetClientArrayLocByID(ClientID)).DataSockHandle
End Property

Public Property Get GetClientfFile(ClientID As Long) As Long
    GetClientfFile = client(GetClientArrayLocByID(ClientID)).fFile
End Property

Public Property Get GetClientHomeDir(ClientID As Long) As String
    GetClientHomeDir = client(GetClientArrayLocByID(ClientID)).HomeDir
End Property

Public Property Get GetClientIdleSince(ClientID As Long) As Long
    GetClientIdleSince = client(GetClientArrayLocByID(ClientID)).IdleSince
End Property

Public Property Get GetClientIPAddress(ClientID As Long) As String
    GetClientIPAddress = client(GetClientArrayLocByID(ClientID)).IPAddress
End Property

Public Property Get GetClientSocketHandle(ClientID As Long) As Long
    GetClientSocketHandle = client(GetClientArrayLocByID(ClientID)).SocketHandle
End Property

Public Property Get GetClientStatus(ClientID As Long) As String
    GetClientStatus = ServerGetClientStatusDescription(client(GetClientArrayLocByID(ClientID)).Status)
End Property

Public Property Get GetClientTotalBytesDownloaded(ClientID As Long) As Long
    GetClientTotalBytesDownloaded = client(GetClientArrayLocByID(ClientID)).TotalBytesDownloaded
End Property

Public Property Get GetClientTotalBytesUploaded(ClientID As Long) As Long
    GetClientTotalBytesUploaded = client(GetClientArrayLocByID(ClientID)).TotalBytesUploaded
End Property

Public Property Get GetClientTotalFilesDownloaded(ClientID As Long) As Long
    GetClientTotalFilesDownloaded = client(GetClientArrayLocByID(ClientID)).TotalFilesDownloaded
End Property

Public Property Get GetClientTotalFilesUploaded(ClientID As Long) As Long
    GetClientTotalFilesUploaded = client(GetClientArrayLocByID(ClientID)).TotalFilesUploaded
End Property

Public Property Get GetClientUserName(ClientID As Long) As String
    GetClientUserName = client(GetClientArrayLocByID(ClientID)).UserName
End Property

'''''''''''''''''''''''''''''''''''''''''''
'Server functions
'''''''''''''''''''''''''''''''''''''''''''
Public Function StartServer() As Long

    'Make sure the server has been givin a port before doing anything.
    If Port < 1 Then
        StartServer = 1001
        Exit Function
    End If

    'Make sure that the MainFormhWnd variable is assigned the hWnd
    'of the main form, this is required by the Winsock object.
    If hWndToMsg < 1 Then
        StartServer = 1002
        Exit Function
    End If

    'If no TransferBufferSize was assigned then default to 4096 (4K)
    If TransferBufferSize = 0 Then TransferBufferSize = 4096

    'Variable to store the result of functions
    Dim r As Long

    'Start winsock
    r = StartWinsock()

    If r <> 0 Then  'Winsock did not start properly
        StartServer = r
        Exit Function
    End If

    r = ListenForConnect(Port)

    If r <> 0 Then   'Problem trying to open listen port.
        StartServer = r
        Exit Function
    End If

    'Server started with out error and is ready to accept clients.
    RaiseEvent ServerStarted

End Function

Private Sub NewClient(SocketHandle As Long, SocketInfo As sockaddr)

    On Error GoTo ShowError

    Dim tmpID As Long
    Dim i As Integer

    'Increment counters
    ConnectedClients = ConnectedClients + 1
    
    ClientCounter = ClientCounter + 1
    'This just helps to make sure that things dont get
    'screwed up incase two clients connected at nearly the exact same time.
    tmpID = ClientCounter

    'Find an open slot in the client() array.
    Do
        i = i + 1
    Loop Until client(i).inUse = False

    'Fill client data
    With client(i)
        .inUse = True
        .ConnectedAt = Now
        .ID = tmpID
        .Status = stat_LOGGING_IN
        .IdleSince = Now
        .IPAddress = GetAscIp(SocketInfo.sin_addr)
        .SocketHandle = SocketHandle
    End With

    RaiseEvent NewClient(client(i).ID)

    If ((ConnectedClients > MaxClients) And (MaxClients <> 0)) Or (ConnectedClients > MAX_CONNECTIONS) Then
        SendResponse SocketHandle, "421 Too many users - try again later."
        LogoutClient i
        Exit Sub
    End If

    SendResponse SocketHandle, "220 P1mp FTP Engine version " & App.Major & ".0" & App.Minor & " build " & App.Revision

ShowError:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " " & Err.Description
    End If

End Sub

Private Sub LoginClient(cArrayLoc As Integer, Password As String)

    'Right now the server only allows anonymous connections,
    'because of this we will allow any password to work.

    'Fill in some dummy data until i put in a user system.
    'When i do, validate passwords here, etc.
    With client(cArrayLoc)
        .HomeDir = "C:\"
        .CurrentDir = .HomeDir
        SendResponse .SocketHandle, "230 User logged in, proceed."
        .Status = stat_IDLE
    End With

    RaiseEvent ClientLoggedIn(ByVal client(cArrayLoc).ID)
    RaiseEvent ClientStatusChanged(ByVal client(cArrayLoc).ID)

End Sub

Public Sub LogoutClient(ByVal ID As Long)

    With client(GetClientArrayLocByID(ID))
        'Close the connection
        closesocket .SocketHandle

        'Reset all client info for next use.
        .ConnectedAt = ""
        .ConnectMode = 0
        .cTotalBytesXfer = 0
        .CurrentDir = ""
        .CurrentFile = ""
        .DataPort = 0
        .DataSockHandle = 0
        .DataToSend = ""
        .fFile = 0
        .HomeDir = ""
        .ID = 0
        .IdleSince = ""
        .IPAddress = ""
        .SocketHandle = 0
        .Status = stat_IDLE
        .TotalBytesDownloaded = 0
        .TotalBytesUploaded = 0
        .TotalFilesDownloaded = 0
        .TotalFilesUploaded = 0
        .UserName = ""
        .inUse = False
    End With

    ConnectedClients = ConnectedClients - 1

    'Client is now officially logged out.
    RaiseEvent ClientLoggedOut(ByVal ID)

End Sub

Public Function GetClientArrayLocByID(ByVal ID As Long) As Integer

    'This function returns the number where the information
    'of the client of the specified ID is located in the client() array.
    Dim i As Integer

    For i = 0 To UBound(client)
        If client(i).ID = ID Then
            GetClientArrayLocByID = i
            Exit Function
        End If
    Next

    'ID not found in client array, return -1 error
    GetClientArrayLocByID = -1

End Function

Private Function GetClientArrayLocBySockH(SocketHandle As Long) As Integer

    'This function returns the number where the information
    'of the client of the specified socket handle is located in the client() array.
    Dim i As Integer

    For i = 0 To UBound(client)
        If client(i).SocketHandle = SocketHandle Then
            GetClientArrayLocBySockH = i
            Exit Function
        End If
    Next

    'SocketHandle not found in client array, return -1 error
    GetClientArrayLocBySockH = -1

End Function

Private Function GetClientArrayLocByDataSockH(SocketHandle As Long) As Integer

    'This function returns the number where the information
    'of the client of the specified socket handle is located in the client() array.
    Dim i As Integer

    For i = 0 To UBound(client)
        If client(i).DataSockHandle = SocketHandle Then
            GetClientArrayLocByDataSockH = i
            Exit Function
        End If
    Next

    'SocketHandle not found in client array, return -1 error
    GetClientArrayLocByDataSockH = -1

End Function

Private Function isCommandSock(ByVal SocketHandle As Long) As Boolean

    'Looks in all the SocketHandle vars in the client() array
    'for a match.
    Dim i As Integer

    For i = 0 To UBound(client)
        If client(i).SocketHandle = SocketHandle Then
            isCommandSock = True
            Exit Function
        End If
    Next

    'SocketHandle not found in client array, return False
    isCommandSock = False

End Function

Private Function isDataSock(ByVal SocketHandle As Long) As Boolean

    'Looks in all the DataSockHandle vars in the client() array
    'for a match.
    Dim i As Integer

    For i = 0 To UBound(client)
        If client(i).DataSockHandle = SocketHandle Then
            isDataSock = True
            Exit Function
        End If
    Next

    'SocketHandle not found in client array, return False
    isDataSock = False

End Function

Private Sub DataArrival(ByVal SocketHandle As Long, ByRef raw_data As String)

    'This function determines if the incoming data is
    'data or a FTP command from the client and
    'then sends a reference of the data to the appropriet sub.
    If isCommandSock(SocketHandle) = True Then
        ProcFTPCommand SocketHandle, raw_data
    ElseIf isDataSock(SocketHandle) = True Then
        ProcData SocketHandle, raw_data
    Else
        RaiseEvent ServerErrorOccurred(1003)
    End If

End Sub

Private Sub ProcFTPCommand(ByVal SocketHandle As Long, ByRef raw_data As String)

    Dim Data
    Dim ftpCommand As String
    Dim ftpArgs As String

    Data = Replace$(raw_data, vbCrLf, "")    'Remove carriage return and line feed.

    If InStr(Data, " ") = 0 Then
        ftpCommand = Data 'Command has no arguments...
    Else
        ftpCommand = Left$(Data, (InStr(Data, " ") - 1))          'Get the command.
        ftpArgs = Right$(Data, (Len(Data) - InStr(Data, " ")))    'Get the command arguments.
    End If

    RaiseEvent ClientSentCommand(client(GetClientArrayLocBySockH(SocketHandle)).ID, ftpCommand, ftpArgs)
    client(GetClientArrayLocBySockH(SocketHandle)).IdleSince = Now

    Select Case UCase$(ftpCommand)

        Case "USER"
            'For now i am simply hardcoding an anonymous account until
            'i get around to putting in a user system.
            If ftpArgs = "anonymous" Then
                client(GetClientArrayLocBySockH(SocketHandle)).UserName = ftpArgs
                SendResponse SocketHandle, "331 User name ok, need password."
            Else
                SendResponse SocketHandle, "530 Not logged in: No such account " & ftpArgs
            End If

        Case "PASS"
            'Now have user name and password, attempt to login the client.
            LoginClient GetClientArrayLocBySockH(SocketHandle), ftpArgs
        '    OLD CODE FROM OLD FTP SERVER
        '    If ftpArgs = "" Then
        '        Dim i As Integer
        '        Dim msgArray() As String
        '            msgArray = Split(Server.WelcomeMsg, vbCrLf)
        '        For i = 0 To UBound(msgArray)
        '            msgArray(i) = "230-" & msgArray(i)
        '            SendResponse SocketHandle, msgArray(i)
        '        Next
        '        login_client Socket
        '        SendResponse SocketHandle, "230 User logged in, proceed."
        '    Else
        '        SendResponse SocketHandle, "530 Not logged in"
        '        logout_client Socket
        '    End If

        Case "TYPE"
            SendResponse SocketHandle, "200 Type set to " & ftpArgs

        Case "REST"
            SendResponse SocketHandle, "350 Restarting at " & ftpArgs & " - send STORE or RETRIEVE to initiate transfer."

        Case "PWD"
            SendResponse SocketHandle, "257 " & Chr(34) _
                & ConvPathToRelative(client(GetClientArrayLocBySockH(SocketHandle)).HomeDir, client(GetClientArrayLocBySockH(SocketHandle)).CurrentDir) _
                & Chr(34) & " is current directory."

        Case "PORT"
            Dim tmpArray() As String 'Six slots required (0 - 5)
            tmpArray = Split(ftpArgs, ",")
            client(GetClientArrayLocBySockH(SocketHandle)).DataPort = tmpArray(4) * 256 Or tmpArray(5)
            SendResponse SocketHandle, "200 Port command successful."

        Case "LIST"
            SendResponse SocketHandle, "150 Opening ASCII mode data connection for /bin/ls."
            client(GetClientArrayLocBySockH(SocketHandle)).Status = stat_GETTING_DIR_LIST
            RaiseEvent ClientStatusChanged(client(GetClientArrayLocBySockH(SocketHandle)).ID)
            GetDirectoryList GetClientArrayLocBySockH(SocketHandle)

        'Case "RETR"
        '    Dim FileSize As Long
        '        FileSize = GetMP3Info(ftpArgs, True)
            
        '    If FileSize = -1 Then
        '        SendResponse SocketHandle, "550 " & ftpArgs & ": No such file or directory."
        '        frmMain.dataSock(Socket).Close
        '    Else
        '        SendResponse SocketHandle, "150 Opening BINARY mode data connection for " & ftpArgs & " (" & FileSize & " bytes)"
        '        open_File Socket, ftpArgs
        '    End If

        Case "CWD"
            ChangeDirectory GetClientArrayLocBySockH(SocketHandle), ftpArgs

        Case "CDUP"
            Dim tmp As String

            tmp = client(GetClientArrayLocBySockH(SocketHandle)).CurrentDir
            
            If Right$(tmp, 1) = "\" Then tmp = Left$(tmp, Len(tmp) - 1)
            tmp = Left$(tmp, InStrRev(tmp, "\"))
            ChangeDirectory GetClientArrayLocBySockH(SocketHandle), ConvPathToRelative(client(GetClientArrayLocBySockH(SocketHandle)).HomeDir, tmp)

        'Case "PASV"
        '    client(Socket).ClientModeType = "PASV"
        '    SendResponse SocketHandle, "227 Entering Passive Mode (" _
        '        & Replace(frmMain.commandSock(0).LocalIP, ".", ",") & OpenLocalDataPort(Socket) & ")"

        Case "NOOP"
            SendResponse SocketHandle, "200 NOOP command successful."

        Case Else
            SendResponse SocketHandle, "502 Command not implemented."

    End Select

End Sub

Private Sub ProcData(ByVal SocketHandle As Long, ByRef raw_data As String)

End Sub

Private Sub GetDirectoryList(cArrayLoc As Integer)

    'This code tends to be very very close when it
    'comes across a directory with lots of files to list.
    'When trying to list my C:\WinNT\System32 directory on my computer
    'it took nearly 10 seconds!  :(
    'If anyone has a faster routine to do this please tell me! =)

    Dim hFile As Long
    Dim r As Long
    Dim fname As String
    Dim WFD As WIN32_FIND_DATA
    Dim dirList As String
    Dim permissions As String

    'Get the first file in the directory (it will usually return ".")
    hFile = FindFirstFile(client(cArrayLoc).CurrentDir & "*.*" + Chr$(0), WFD)
    
    'If nothing returned, bail out.
    'If hFile < 0 Then
    '    frmMain.dataSock(socket).Close
    '    Exit Sub
    'End If

    If Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> "." And Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> ".." Then
        If (WFD.dwFileAttributes And vbDirectory) Then
            permissions = "drwx------"
        Else
            permissions = "-rwx------"
        End If
        
        dirList = permissions _
            & " 1 user group " _
            & WFD.nFileSizeLow _
            & get_date(WFD.ftLastWriteTime) _
            & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) _
            & vbCrLf
    End If

    While FindNextFile(hFile, WFD)
        If Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> "." And Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> ".." Then
            If (WFD.dwFileAttributes And vbDirectory) Then
                permissions = "drwx------"
            Else
                permissions = "-rwx------"
            End If
            dirList = dirList _
                & permissions _
                & " 1 user group " _
                & WFD.nFileSizeLow _
                & get_date(WFD.ftLastWriteTime) _
                & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) _
                & vbCrLf
        End If

        DoEvents

    Wend

    'Close the search handle
    r = FindClose(hFile)
    
    MakeDataConnection cArrayLoc

    If hFile = -1 Then
        closesocket client(cArrayLoc).DataSockHandle
        Exit Sub
    End If

    send client(cArrayLoc).DataSockHandle, dirList, Len(dirList), 0

End Sub

Private Sub ChangeDirectory(cArrayLoc As Integer, ChangeTo As String)

    If Left$(ChangeTo, 1) = "/" Then
        'Client specified an absolute path.
        If FileExists(ConvPathToLocal(client(cArrayLoc).HomeDir, ChangeTo)) = True Then
            client(cArrayLoc).CurrentDir = ConvPathToLocal(client(cArrayLoc).HomeDir, ChangeTo)
        Else
            SendResponse client(cArrayLoc).SocketHandle, "550 " & ChangeTo & ": No such file or directory."
            Exit Sub
        End If
    Else
        'Client specified a path relative to their current path.
        If FileExists(ConvPathToLocal(client(cArrayLoc).CurrentDir, ChangeTo)) = True Then
            client(cArrayLoc).CurrentDir = ConvPathToLocal(client(cArrayLoc).CurrentDir, ChangeTo)
        Else
            SendResponse client(cArrayLoc).SocketHandle, "550 " & ChangeTo & ": No such file or directory."
            Exit Sub
        End If
    End If

    SendResponse client(cArrayLoc).SocketHandle, "250 Directory changed to " & ConvPathToRelative(client(cArrayLoc).HomeDir, client(cArrayLoc).CurrentDir)

End Sub

Public Function ShutdownServer() As Long

    ShutdownServer = EndWinsock
    UnSubClass
    RaiseEvent ServerStopped

End Function

Public Function ConvPathToLocal(ByVal StartPath As String, ByVal CurrentPath As String) As String

    'This function takes a relative path and converts to to a local one.
    'eg. home directory is "/", locally "/" equates to "C:\Files\".
    'Their current directory is "/Uploads/Stuff".
    'If passed through this function you will get
    '"C:\Files\Uploads\Stuff\"

    'StartPath must be a local path.  "C:\Files\" identifying the home directory.
    'CurrentPath is the clients current relative path.  "/Uploads/Stuff"

    Dim result As String

    If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"

    If Left$(CurrentPath, 1) = "/" Then CurrentPath = Right$(CurrentPath, Len(CurrentPath) - 1)
    CurrentPath = Replace$(CurrentPath, "/", "\")

    result = StartPath & CurrentPath

    If Right$(result, 1) <> "\" Then result = result & "\"

    ConvPathToLocal = result

End Function

Public Function ConvPathToRelative(ByVal StartPath As String, ByVal CurrentPath As String) As String

    'This function takes an absolute path and converts it to a relative unix style one.
    'eg. home directory is "C:\MyFiles\Music\" and they are currently in
    '"C:\MyFiles\Music\Alternative\", pass these two paths into this function
    'will give you the result /Alternative

    If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"
    If Right$(CurrentPath, 1) <> "\" Then CurrentPath = CurrentPath & "\"

    Dim strRelPath As String

    If StartPath = CurrentPath Then
        strRelPath = "/"
    Else
        strRelPath = Replace$(CurrentPath, StartPath, "/")
        strRelPath = Replace$(strRelPath, "\", "/")
    
        If Right$(strRelPath, 1) = "/" Then strRelPath = Left$(strRelPath, Len(strRelPath) - 1)
    End If

    ConvPathToRelative = strRelPath

End Function

Public Function ServerGetClientStatusDescription(ByVal stat As Integer) As String

    Select Case stat
        Case stat_IDLE: ServerGetClientStatusDescription = "Idle"
        Case stat_LOGGING_IN: ServerGetClientStatusDescription = "Connecting..."
        Case stat_GETTING_DIR_LIST: ServerGetClientStatusDescription = "Downloading list of files"
        Case stat_UPLOADING: ServerGetClientStatusDescription = "Uploading"
        Case stat_DOWNLOADING: ServerGetClientStatusDescription = "Downloading"
        Case Else: ServerGetClientStatusDescription = "Unknown status"
    End Select

End Function

Public Function ServerGetErrorDescription(ByVal errCode As Long) As String

    If errCode < 10000 Then
        Select Case errCode
    
            Case 1001: ServerGetErrorDescription = "Server failed to start becuase no port or invalid port was specified."
            Case 1002: ServerGetErrorDescription = "Server failed to start becuase a valid form handle was not assigned to hWndToMsg."
            Case 1003: ServerGetErrorDescription = "Socket could not be found from which the received data originated."
            
            Case Else: ServerGetErrorDescription = "Unknown error " & errCode
    
        End Select
    Else 'If errCode > 10000 then it is a Winsock error
        ServerGetErrorDescription = "Winsock Error:" & vbCrLf & GetWSAErrorString(errCode)
    End If

End Function


'================================= MISC CODE =================================


Private Function get_date(FT As FILETIME) As String

    'convert the FILETIME (FT) to a SYSTEMTIME (ST) type
    Dim ST As SYSTEMTIME
    Dim r As Long
    Dim ds As String
    
    r = FileTimeToSystemTime(FT, ST)

    ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)

    If DateDiff("d", ds, Date) > 365 Then
        get_date = Format$(ds, " mmm dd yyyy ")
    Else
        get_date = Format$(ds & " " & ST.wHour & ":" & ST.wMinute, " mmm dd hh:mm ")
    End If

End Function

Private Function FileExists(FileName As String) As Boolean

    Dim hFindFile As Long
    Dim FileData As WIN32_FIND_DATA
    If Right(FileName, 1) = "\" Then
        FileName = Left(FileName, Len(FileName) - 1)
    End If

    hFindFile = FindFirstFile(FileName, FileData)
    If hFindFile = -1 Then
        FileExists = False
    Else
        FileExists = True
    End If

    FindClose hFindFile

End Function
