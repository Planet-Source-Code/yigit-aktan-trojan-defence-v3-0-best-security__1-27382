Attribute VB_Name = "modWinsock"
Option Explicit

Public Const INADDR_NONE = &HFFFF
Public Const SOCKET_ERROR = -1
Public Const WSABASEERR = 10000
Public Const WSAEFAULT = (WSABASEERR + 14)
Public Const WSAEINVAL = (WSABASEERR + 22)
Public Const WSAEINPROGRESS = (WSABASEERR + 50)
Public Const WSAENETDOWN = (WSABASEERR + 50)
Public Const WSASYSNOTREADY = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Public Const WSANOTINITIALISED = (WSABASEERR + 93)
Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004

Public Type WSAData
    wVersion       As Integer
    wHighVersion   As Integer
    szDescription  As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpVendorInfo   As Long
End Type

Public Type HOSTENT
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type

Public Type servent
    s_name    As Long
    s_aliases As Long
    s_port    As Integer
    s_proto   As Long
End Type

Public Type protoent
    p_name    As String
    p_aliases As Long
    p_proto   As Long
End Type

Public Declare Function WSAStartup Lib "ws2_32.dll" _
        (ByVal wVR As Long, lpWSAD As WSAData) As Long

Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Public Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, _
        ByVal addr_len As Long, ByVal addr_type As Long) As Long
        
Public Declare Function gethostbyname Lib "ws2_32.dll" _
        (ByVal host_name As String) As Long
        
Public Declare Function gethostname Lib "ws2_32.dll" _
        (ByVal host_name As String, ByVal namelen As Long) As Long

Public Declare Function getservbyname Lib "ws2_32.dll" _
        (ByVal serv_name As String, ByVal proto As String) As Long

Public Declare Function getprotobynumber Lib "ws2_32.dll" _
        (ByVal proto As Long) As Long

Public Declare Function getprotobyname Lib "ws2_32.dll" _
        (ByVal proto_name As String) As Long

Public Declare Function getservbyport Lib "ws2_32.dll" _
        (ByVal port As Integer, ByVal proto As Long) As Long

Public Declare Function inet_addr Lib "ws2_32.dll" _
        (ByVal cp As String) As Long

Public Declare Function inet_ntoa Lib "ws2_32.dll" _
        (ByVal inn As Long) As Long

Public Declare Function htons Lib "ws2_32.dll" _
        (ByVal hostshort As Integer) As Integer

Public Declare Function htonl Lib "ws2_32.dll" _
        (ByVal hostlong As Long) As Long

Public Declare Function ntohl Lib "ws2_32.dll" _
        (ByVal netlong As Long) As Long

Public Declare Function ntohs Lib "ws2_32.dll" _
        (ByVal netshort As Integer) As Integer

Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, _
        ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
        (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" _
        (ByVal lpString As Any) As Long

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

Public Function UnsignedToLong(Value As Double) As Long
    
    If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If

End Function

Public Function LongToUnsigned(Value As Long) As Double
    
    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If

End Function

Public Function UnsignedToInteger(Value As Long) As Integer
    
    If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If

End Function

Public Function IntegerToUnsigned(Value As Integer) As Long
    
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If

End Function

Public Function StringFromPointer(ByVal lPointer As Long) As String
    Dim strTemp As String
    Dim lRetVal As Long

    lRetVal = lstrcpy(ByVal strTemp, ByVal lPointer)
    If lRetVal Then StringFromPointer = strTemp

End Function


