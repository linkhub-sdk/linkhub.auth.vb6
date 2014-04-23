VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Linkhub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================
' Class Module for develop interoperation with Linkhub APIs.
' Functionalities are authentication for Linkhub api products, and to support
' several base infomation(ex. Remain point).
'
' This module uses advapi32.dll and crypt32.dllm. most of windows system has it
' on system32 by default.
'
' http://www.linkhub.co.kr
' Author : Kim Seongjun (pallet027@gmail.com)
' Written : 2014-04-22
'
' Thanks for your interest.
'
' This module has referenced some code from VB-JSON project
' (which is from http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html)
' For the License...
'
' And cStringBuilder.cls is from Steve McMahon(steve@vbaccelerator.com)'s. It speed up
' json String processing.
'
'=================================================================================
Option Explicit


'''''''''' for json''''''''''''''''''
Const INVALID_JSON      As Long = 1
Const INVALID_OBJECT    As Long = 2
Const INVALID_ARRAY     As Long = 3
Const INVALID_BOOLEAN   As Long = 4
Const INVALID_NULL      As Long = 5
Const INVALID_KEY       As Long = 6
Const INVALID_RPC_CALL  As Long = 7

Private psErrors As String
'''''''''''''''''''''''''''''''''''''

'''''''''' for win32 ''''''''''''''''''''
Private Const CP_UTF8 As Long = 65001

Private Const CALG_RC2                   As Long = &H6602&
Private Const CALG_MD5                   As Long = &H8003&
Private Const CALG_SHA1                  As Long = &H8004&
Private Const CALG_HMAC                  As Long = &H8009&

Private Const PROV_RSA_FULL              As Long = 1
Private Const PROV_RSA_AES               As Long = 24
Private Const CRYPT_VERIFYCONTEXT        As Long = &HF0000000
Private Const MS_DEFAULT_PROVIDER        As String = _
    "Microsoft Base Cryptographic Provider v1.0"

Private Const HP_HASHVAL                 As Long = 2
Private Const HP_HASHSIZE                As Long = 4
Private Const HP_HMAC_INFO               As Long = 5

Private Const CRYPT_STRING_BASE64        As Long = &H1&
Private Const CRYPT_STRING_HEX           As Long = &H4&
Private Const CRYPT_STRING_HEXASCII      As Long = &H5&
Private Const CRYPT_STRING_HEXADDR       As Long = &HA&
Private Const CRYPT_STRING_HEXASCIIADDR  As Long = &HB&
Private Const CRYPT_STRING_HEXRAW        As Long = &HC&       'Requires Vista or later, so we emulate.
Private Const CRYPT_STRING_NOCR          As Long = &H80000000
Private Const CRYPT_STRING_NOCRLF        As Long = &H40000000 'Requires Vista or later!

Private Const CRYPT_IPSEC_HMAC_KEY       As Long = &H100&

Private Const PLAINTEXTKEYBLOB           As Byte = &H8
Private Const CUR_BLOB_VERSION           As Byte = &H2

Private Type HMAC_INFO
    HashAlgId As Long
    pbInnerString As Long
    cbInnerString As Long
    pbOuterString As Long
    cbOuterString As Long
End Type

Private Type BLOBHEADER
    bType As Byte
    bVersion As Byte
    reserved As Integer
    aiKeyAlg As Long
End Type

Private Type KEYBLOB
    hdr As BLOBHEADER
    cbKeySize As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long)

Private Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" ( _
    lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long

Private Declare Function CryptAcquireContext Lib "Advapi32" Alias "CryptAcquireContextW" ( _
    ByRef phProv As Long, _
    ByVal pszContainer As Long, _
    ByVal pszProvider As Long, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptBinaryToString Lib "Crypt32" _
    Alias "CryptBinaryToStringW" ( _
    ByRef pbBinary As Byte, _
    ByVal cbBinary As Long, _
    ByVal dwFlags As Long, _
    ByVal pszString As Long, _
    ByRef pcchString As Long) As Long

Private Declare Function CryptCreateHash Lib "Advapi32" ( _
    ByVal hProv As Long, _
    ByVal AlgId As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phHash As Long) As Long
    
Private Declare Function CryptDestroyHash Lib "Advapi32" ( _
    ByVal hHash As Long) As Long
    
Private Declare Function CryptDestroyKey Lib "Advapi32" ( _
    ByVal hKey As Long) As Long

Private Declare Function CryptGetHashParam Lib "Advapi32" ( _
    ByVal hHash As Long, _
    ByVal dwParam As Long, _
    ByRef pbData As Any, _
    ByRef pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptHashData Lib "Advapi32" ( _
    ByVal hHash As Long, _
    ByRef pbData As Any, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptImportKey Lib "Advapi32" ( _
    ByVal hProv As Long, _
    ByVal pbData As Long, _
    ByVal dwDataLen As Long, _
    ByVal hPubKey As Long, _
    ByVal dwFlags As Long, _
    ByRef phKey As Long) As Long

Private Declare Function CryptReleaseContext Lib "Advapi32" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptSetHashParam Lib "Advapi32" ( _
    ByVal hHash As Long, _
    ByVal dwParam As Long, _
    ByRef pbData As HMAC_INFO, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptStringToBinary Lib "Crypt32" _
    Alias "CryptStringToBinaryW" ( _
    ByVal pszString As Long, _
    ByVal cchString As Long, _
    ByVal dwFlags As Long, _
    ByVal pbBinary As Long, _
    ByRef pcbBinary As Long, _
    ByRef pdwSkip As Long, _
    ByRef pdwFlags As Long) As Long

Private hBaseProvider As Long
Private hKey As Long
Private hHmacHash As Long
Private blnIsWinXP As Boolean
Private blnIsWin5_1 As Boolean

Public Enum EncDecFormat
    edfBase64
    edfHex
    edfHexAscii
    edfHexAddr
    edfHexAsciiAddr
    edfHexRaw
End Enum
#If False Then
Dim edfBase64, edfHex, edfHexAscii, edfHexAddr, edfHexAsciiAddr, edfHexRaw
#End If

Public Enum EncodeFolding
    efCrLf
    efLf
    efNoFolding
End Enum
#If False Then
Dim efCrLf, efLf, efNoFolding
#End If


Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
''''''''''''''''''''''''''end of win32''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const ServiceURL = "https://auth.linkhub.co.kr"
Private m_isTest As Boolean
Private m_partnerID As String
Private m_secretKey As String

Private m_LastErrCode As Long
Private m_LastErrMessage As String



Public Property Let partnerID(ByVal value As String)
    m_partnerID = value
End Property

Public Property Let SercetKey(ByVal value As String)
    m_secretKey = value
End Property

Public Property Let IsTest(ByVal value As Boolean)
    m_isTest = value
End Property

Public Property Get LastErrCode() As Long
    LastErrCode = m_LastErrCode
End Property
Public Property Get LastErrMessage() As String
    LastErrMessage = m_LastErrMessage
End Property


Public Function getToken(ServiceID As String, access_id As String, scope As Collection) As LinkhubToken
    
    On Error GoTo errHandler
    
    Dim mlinkhubToken As New LinkhubToken
    
    Dim winhttp1
    Dim postdata As String
    
    
    Dim postObject As New Dictionary
    
    postObject.Add "access_id", access_id
    postObject.Add "scope", scope
    
    postdata = toString(postObject)
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", ServiceURL + "/" + ServiceID + "/Token")
    
    Dim xdate As String
    xdate = Format(UTCTime, "yyyy-MM-ddTHH:mm:ss.sssZ")
    
    Call winhttp1.setRequestHeader("x-lh-date", xdate)
    Call winhttp1.setRequestHeader("x-lh-version", "1.0")
    'Call winhttp1.setRequestHeader("x-lh-forwarded", "123.123.123.123")
    
    Dim target As String
    target = "POST" + Chr(10)
    target = target + Encode(MD5(ToUTF8(postdata)), edfBase64, efNoFolding) + Chr(10)
    target = target + xdate + Chr(10)
    'target = target + "123.123.123.123" + Chr(10)
    target = target + "1.0" + Chr(10)
    target = target + "/" + ServiceID + "/Token"
    
    Dim Bearer As String
    Dim btKey
    
    btKey = Decode(m_secretKey, edfBase64)
    InitHmac (btKey)
    Bearer = Encode(HmacSha1(ToUTF8(target)), edfBase64)
    
    Call winhttp1.setRequestHeader("Authorization", "LINKHUB " + m_partnerID + " " + Bearer)
    
    winhttp1.send (postdata)
    winhttp1.WaitForResponse
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = parse(Response)
        
        m_LastErrCode = parsedDic.Item("code")
        m_LastErrMessage = parsedDic.Item("message")
        
        Set getToken = Nothing
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
        
    Set parsedDic = parse(Response)
        
    mlinkhubToken.session_token = parsedDic.Item("session_token")
    mlinkhubToken.ServiceID = parsedDic.Item("serviceID")
    mlinkhubToken.partnerID = parsedDic.Item("partnerID")
    mlinkhubToken.usercode = parsedDic.Item("usercode")
    mlinkhubToken.expiration = parsedDic.Item("expiration")
    mlinkhubToken.ipaddress = parsedDic.Item("ipaddress")
    Set mlinkhubToken.scope = parsedDic.Item("scope")
       
    Set getToken = mlinkhubToken
    Exit Function
    
errHandler:
    m_LastErrCode = parsedDic.Item("code")
    m_LastErrMessage = parsedDic.Item("message")
    
    Set getToken = Nothing
    
End Function

Public Function GetBalance(BearerToken As String, ServiceID As String) As Double
    On Error GoTo errHandler
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("GET", ServiceURL + "/" + ServiceID + "/Point")
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    
    winhttp1.send
    winhttp1.WaitForResponse
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = parse(Response)
        
        m_LastErrCode = parsedDic.Item("code")
        m_LastErrMessage = parsedDic.Item("message")
        
        GetBalance = m_LastErrCode
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set parsedDic = parse(Response)
    
    GetBalance = CDbl(parsedDic.Item("remainPoint"))
    
    Exit Function
errHandler:
    m_LastErrCode = parsedDic.Item("code")
    m_LastErrMessage = parsedDic.Item("message")
    
     GetBalance = m_LastErrCode
End Function

Public Function GetPartnerBalance(BearerToken As String, ServiceID As String) As Double
    On Error GoTo errHandler
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("GET", ServiceURL + "/" + ServiceID + "/PartnerPoint")
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    
    winhttp1.send
    winhttp1.WaitForResponse
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = parse(Response)
        
        m_LastErrCode = parsedDic.Item("code")
        m_LastErrMessage = parsedDic.Item("message")
        
        GetPartnerBalance = m_LastErrCode
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set parsedDic = parse(Response)
    
    GetPartnerBalance = CDbl(parsedDic.Item("remainPoint"))
    
    Exit Function
errHandler:
    m_LastErrCode = parsedDic.Item("code")
    m_LastErrMessage = parsedDic.Item("message")
    
     GetPartnerBalance = m_LastErrCode
End Function

Private Function UTCTime() As Date
    Dim T As SYSTEMTIME
    GetSystemTime T
    UTCTime = DateSerial(T.wYear, T.wMonth, T.wDay) + TimeSerial(T.wHour, T.wMinute, T.wSecond) + T.wMilliseconds / 86400000#
End Function


Public Function Decode( _
    ByVal Encoded As String, _
    Optional ByVal Format As EncDecFormat = edfBase64) As Byte()
    
    Dim dwFlags As Long
    Dim bytBuf() As Byte
    Dim lngOutLen As Long
    Dim dwActualUsed As Long
    
    Select Case Format
        Case edfBase64
            dwFlags = CRYPT_STRING_BASE64
        Case edfHex
            dwFlags = CRYPT_STRING_HEX
        Case edfHexAscii
            dwFlags = CRYPT_STRING_HEXASCII
        Case edfHexAddr
            dwFlags = CRYPT_STRING_HEXADDR
        Case edfHexAsciiAddr
            dwFlags = CRYPT_STRING_HEXASCIIADDR
        Case edfHexRaw
            dwFlags = CRYPT_STRING_HEXRAW
    End Select
    
    If blnIsWinXP And (Format = edfHexRaw) Then
        'Emulate missing format.
        Dim i As Long
        
        ReDim bytBuf(Len(Encoded) \ 2 - 1)
        For i = 1 To Len(Encoded) Step 2
            bytBuf((i - 1) \ 2) = CLng("&H0" & Mid$(Encoded, i, 2))
        Next
        Decode = bytBuf
    ElseIf CryptStringToBinary(StrPtr(Encoded), _
                              Len(Encoded), _
                              dwFlags, _
                              0, _
                              lngOutLen, _
                              0, _
                              dwActualUsed) = 0 Then
        Err.Raise vbObjectError Or &HD302&, _
                  "HS1.Decode", _
                  "Failed to determine decoded length, system error " _
                & CStr(Err.LastDllError)
    Else
        ReDim bytBuf(lngOutLen - 1)
        If CryptStringToBinary(StrPtr(Encoded), _
                               Len(Encoded), _
                               dwFlags, _
                               VarPtr(bytBuf(0)), _
                               lngOutLen, _
                               0, _
                               dwActualUsed) = 0 Then
            Err.Raise vbObjectError Or &HD304&, _
                      "HS1.Decode", _
                      "Failed to decode value, system error " _
                    & CStr(Err.LastDllError)
        Else
            Decode = bytBuf
        End If
    End If
End Function

Public Function Encode( _
    ByRef Bytes() As Byte, _
    Optional ByVal Format As EncDecFormat = edfHexRaw, _
    Optional ByVal Folding As EncodeFolding = efCrLf) As String
    
    Dim dwFlags As Long
    Dim lngOutLen As Long
    Dim strEncoded As String
    
    Select Case Format
        Case edfBase64
            dwFlags = CRYPT_STRING_BASE64
        Case edfHex
            dwFlags = CRYPT_STRING_HEX
        Case edfHexAscii
            dwFlags = CRYPT_STRING_HEXASCII
        Case edfHexAddr
            dwFlags = CRYPT_STRING_HEXADDR
        Case edfHexAsciiAddr
            dwFlags = CRYPT_STRING_HEXASCIIADDR
        Case edfHexRaw
            dwFlags = CRYPT_STRING_HEXRAW
    End Select
    
    Select Case Folding
        Case efNoFolding
            dwFlags = dwFlags Or CRYPT_STRING_NOCRLF
        Case efLf
            dwFlags = dwFlags Or CRYPT_STRING_NOCR
    End Select

    If blnIsWinXP And (Format = edfHexRaw) Then
        'Emulate missing format.
        Dim i As Long
        
        strEncoded = String$(2 * (UBound(Bytes) - LBound(Bytes) + 1), 0)
        For i = 1 To Len(strEncoded) Step 2
            Mid$(strEncoded, i, 2) = Right$("0" & Hex$(Bytes((i - 1) \ 2 + LBound(Bytes))), 2)
        Next
        Encode = strEncoded
    ElseIf CryptBinaryToString(Bytes(LBound(Bytes)), _
                               UBound(Bytes) - LBound(Bytes) + 1, _
                               dwFlags, _
                               0&, _
                               lngOutLen) = 0 Then
        Err.Raise vbObjectError Or &HC30A&, _
                  "HS256.Encode", _
                  "Failed to determine encoded length, system error " _
                & CStr(Err.LastDllError)
    Else
        strEncoded = String$(lngOutLen - 1, 0)
        If CryptBinaryToString(Bytes(LBound(Bytes)), _
                               UBound(Bytes) - LBound(Bytes) + 1, _
                               dwFlags, _
                               StrPtr(strEncoded), _
                               lngOutLen) = 0 Then
            Err.Raise vbObjectError Or &HC30C&, _
                      "HS256.Encode", _
                      "Failed to encode value, system error " _
                    & CStr(Err.LastDllError)
        Else
            If blnIsWin5_1 Then
                Select Case Folding
                    Case efNoFolding
                        Encode = Replace$(strEncoded, vbNewLine, "")
                    Case efLf
                        Encode = Replace$(strEncoded, vbCr, "")
                    Case Else
                        Encode = strEncoded
                End Select
            Else
                Encode = strEncoded
            End If
        End If
    End If
End Function

Public Function FromUTF8(ByRef UTF8() As Byte) As String
    Dim lngOutLen As Long
    Dim strWide As String
    
    lngOutLen = MultiByteToWideChar(CP_UTF8, _
                                    0, _
                                    VarPtr(UTF8(LBound(UTF8))), _
                                    UBound(UTF8) - LBound(UTF8) + 1, _
                                    0, _
                                    0)
    If lngOutLen = 0 Then
        Err.Raise vbObjectError Or &HD312&, _
                  "HS1.FromUTF8", _
                  "Failed to decode string, system error " _
                & CStr(Err.LastDllError)
    Else
        strWide = String$(lngOutLen, 0)
        lngOutLen = MultiByteToWideChar(CP_UTF8, _
                                        0, _
                                        VarPtr(UTF8(LBound(UTF8))), _
                                        UBound(UTF8) - LBound(UTF8) + 1, _
                                        StrPtr(strWide), _
                                        lngOutLen)
        If lngOutLen = 0 Then
            Err.Raise vbObjectError Or &HD312&, _
                      "HS1.FromUTF8", _
                      "Failed to decode string, system error " _
                    & CStr(Err.LastDllError)
        Else
            FromUTF8 = strWide
        End If
    End If
End Function

Private Function HmacSha1(ByRef Data() As Byte) As Byte()
    Dim lngErr As Long
    Dim HmacInfo As HMAC_INFO
    Dim lngDataLen As Long
    Dim lngHashSize As Long
    Dim bytHashValue() As Byte
    
    If hKey = 0 Then
        Err.Raise vbObjectError Or &HD322&, _
                  "HS1.HmacSha1", _
                  "No key set, call InitHmac first"
    ElseIf CryptCreateHash(hBaseProvider, CALG_HMAC, hKey, 0, hHmacHash) = 0 Then
        lngErr = Err.LastDllError
        DestroyHandles
        Err.Raise vbObjectError Or &HD32A&, _
                  "HS1.HmacSha1", _
                  "Failed to create HMAC hash object, system error " _
                & CStr(lngErr)
    Else
        HmacInfo.HashAlgId = CALG_SHA1
        If CryptSetHashParam(hHmacHash, HP_HMAC_INFO, HmacInfo, 0&) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HD32C&, _
                      "HS1.HmacSha1", _
                      "Failed to set HMAC_INFO hash param, system error " _
                    & CStr(lngErr)
        ElseIf CryptHashData(hHmacHash, _
                             Data(LBound(Data)), _
                             UBound(Data) - LBound(Data) + 1, _
                             0&) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HD32E&, _
                      "HS1.HmacSha1", _
                      "Failed to hash data, system error " _
                    & CStr(lngErr)
        Else
            lngDataLen = 4 '4 bytes for Long length.
            If CryptGetHashParam(hHmacHash, HP_HASHSIZE, lngHashSize, lngDataLen, 0&) = 0 Then
                lngErr = Err.LastDllError
                DestroyHandles
                Err.Raise vbObjectError Or &HD332&, _
                          "HS1.HmacSha1", _
                          "Failed to obtain hash value length, system error " _
                        & CStr(lngErr)
            Else
                lngDataLen = lngHashSize
                ReDim bytHashValue(lngDataLen - 1)
                If CryptGetHashParam(hHmacHash, HP_HASHVAL, bytHashValue(0), lngDataLen, 0&) = 0 Then
                    lngErr = Err.LastDllError
                    DestroyHandles
                    Err.Raise vbObjectError Or &HD334&, _
                              "HS1.HmacSha1", _
                              "Failed to obtain hash value, system error " _
                            & CStr(lngErr)
                Else
                    DestroyHandles
                    HmacSha1 = bytHashValue
                End If
            End If
        End If
    End If
End Function

Private Sub InitHmac(ByRef key() As Byte)
    Dim hHash As Long
    Dim kbKey As KEYBLOB
    Dim bytKbKey() As Byte
    Dim lngErr As Long

    DestroyHandles
    If hBaseProvider = 0 Then
        Err.Raise vbObjectError Or &HD342&, _
                  "HS1.InitHmac", _
                  "No cryptographic Base provider context"
    Else
        With kbKey
            With .hdr
                .bType = PLAINTEXTKEYBLOB
                .bVersion = CUR_BLOB_VERSION
                .aiKeyAlg = CALG_RC2
            End With
            .cbKeySize = UBound(key) - LBound(key) + 1
            ReDim bytKbKey(LenB(kbKey) + .cbKeySize - 1)
            CopyMemory VarPtr(bytKbKey(0)), VarPtr(kbKey), LenB(kbKey)
            CopyMemory VarPtr(bytKbKey(LenB(kbKey))), VarPtr(key(LBound(key))), .cbKeySize
        End With
        If CryptImportKey(hBaseProvider, _
                          VarPtr(bytKbKey(0)), _
                          UBound(bytKbKey) + 1, _
                          0, _
                          CRYPT_IPSEC_HMAC_KEY, _
                          hKey) = 0 Then
            lngErr = Err.LastDllError
            DestroyHandles
            Err.Raise vbObjectError Or &HD344&, _
                      "HS1.InitHmac", _
                      "Failed to import key, system error " _
                    & CStr(lngErr)
        End If
    End If
End Sub

Private Function MD5(ByRef Data() As Byte) As Byte()
    Dim hHash As Long
    Dim lngDataLen As Long
    Dim lngHashSize As Long
    Dim bytHashValue() As Byte
    
    If hBaseProvider = 0 Then
        Err.Raise vbObjectError Or &HD352&, _
                  "HS1.MD5", _
                  "No cryptographic Base provider context"
    ElseIf CryptCreateHash(hBaseProvider, CALG_MD5, 0&, 0&, hHash) = 0 Then
        Err.Raise vbObjectError Or &HD354&, _
                  "HS1.MD5", _
                  "Failed to create CryptoAPI Hash object, system error " _
                & CStr(Err.LastDllError)
    ElseIf CryptHashData(hHash, _
                         Data(LBound(Data)), _
                         UBound(Data) - LBound(Data) + 1, _
                         0&) = 0 Then
        CryptDestroyHash hHash
        Err.Raise vbObjectError Or &HD356&, _
                  "HS1.MD5", _
                  "Failed to hash data, system error " _
                & CStr(Err.LastDllError)
    Else
        lngDataLen = 4 '4 bytes for Long length.
        If CryptGetHashParam(hHash, HP_HASHSIZE, lngHashSize, lngDataLen, 0&) = 0 Then
            CryptDestroyHash hHash
            Err.Raise vbObjectError Or &HD358&, _
                      "HS1.MD5", _
                      "Failed to obtain hash value length, system error " _
                    & CStr(Err.LastDllError)
        Else
            lngDataLen = lngHashSize
            ReDim bytHashValue(lngDataLen - 1)
            
            If CryptGetHashParam(hHash, HP_HASHVAL, bytHashValue(0), lngDataLen, 0&) = 0 Then
                CryptDestroyHash hHash
                Err.Raise vbObjectError Or &HD35A&, _
                          "HS1.MD5", _
                          "Failed to obtain hash value, system error " _
                        & CStr(Err.LastDllError)
            Else
                CryptDestroyHash hHash
                MD5 = bytHashValue
            End If
        End If
    End If
End Function

Public Function ToUTF8(ByVal Text As String) As Byte()
    Dim lngOutLen As Long
    Dim UTF8() As Byte
    
    lngOutLen = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), Len(Text), _
                                    0, 0, 0, 0)
    ReDim UTF8(lngOutLen - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(Text), Len(Text), _
                        VarPtr(UTF8(0)), lngOutLen, 0, 0
    ToUTF8 = UTF8
End Function

Private Sub DestroyHandles(Optional ByVal Release As Boolean = False)
    On Error Resume Next 'Handle all exceptions here!
    If hHmacHash <> 0 Then CryptDestroyHash hHmacHash: hHmacHash = 0
    If hKey <> 0 Then CryptDestroyKey hKey: hKey = 0
    If Release And (hBaseProvider <> 0) Then CryptReleaseContext hBaseProvider, 0&: hBaseProvider = 0
    Err.Clear
End Sub

Private Sub Class_Initialize()
    Dim osvResult As OSVERSIONINFO
    
    With osvResult
        .dwOSVersionInfoSize = Len(osvResult)
        .szCSDVersion = ""
        If GetVersionEx(osvResult) = 0 Then
            Err.Raise vbObjectError Or &HD362&, _
                      "HS1.Class_Initialize", _
                      "Failed to obtain Windows version, system error " _
                    & CStr(Err.LastDllError)
        Else
            If .dwMajorVersion = 5 And .dwMinorVersion >= 1 Then
                blnIsWin5_1 = True 'XP or Win 2003 Server.
                If .wProductType = 1 Then
                    'XP 32 or 64 bit.
                    blnIsWinXP = True
                Else
                    'Win 2003 Server.
                End If
            ElseIf CCur(.dwMajorVersion) + CCur(.dwMinorVersion) / 1000@ >= 5.002@ Then
                'Vista or later.
            Else
                Err.Raise vbObjectError Or &HD364&, _
                          "HS1.Class_Initialize", _
                          "Requires Windows XP or later"
            End If
        End If
    End With
    If CryptAcquireContext(hBaseProvider, _
                           0&, _
                           StrPtr(MS_DEFAULT_PROVIDER), _
                           PROV_RSA_FULL, _
                           CRYPT_VERIFYCONTEXT) = 0 Then
        Err.Raise vbObjectError Or &HD366&, _
                  "HS1.Class_Initialize", _
                  "Failed to obtain CryptoAPI Base context, system error " _
                & CStr(Err.LastDllError)
    End If
End Sub

Private Sub Class_Terminate()
    DestroyHandles Release:=True
End Sub

'''''''''''''''''''''' json

Private Function GetParserErrors() As String
   GetParserErrors = psErrors
End Function

Private Function ClearParserErrors() As String
   psErrors = ""
End Function


'
'   parse string and create JSON object
'
Public Function parse(ByRef str As String) As Object

   Dim index As Long
   index = 1
   psErrors = ""
   On Error Resume Next
   Call skipChar(str, index)
   Select Case Mid(str, index, 1)
      Case "{"
         Set parse = parseObject(str, index)
      Case "["
         Set parse = parseArray(str, index)
      Case Else
         psErrors = "Invalid JSON"
   End Select


End Function

 '
 '   parse collection of key/value
 '
Public Function parseObject(ByRef str As String, ByRef index As Long) As Variant

   Set parseObject = CreateObject("Scripting.Dictionary")
   Dim sKey As String
   
   ' "{"
   Call skipChar(str, index)
   If Mid(str, index, 1) <> "{" Then
      psErrors = psErrors & "Invalid Object at position " & index & " : " & Mid(str, index) & vbCrLf
      Exit Function
   End If
   
   index = index + 1

   Do
      Call skipChar(str, index)
      If "}" = Mid(str, index, 1) Then
         index = index + 1
         Exit Do
      ElseIf "," = Mid(str, index, 1) Then
         index = index + 1
         Call skipChar(str, index)
      ElseIf index > Len(str) Then
         psErrors = psErrors & "Missing '}': " & Right(str, 20) & vbCrLf
         Exit Do
      End If

      
      ' add key/value pair
      sKey = parseKey(str, index)
      On Error Resume Next
      
      parseObject.Add sKey, parseValue(str, index)
      If Err.Number <> 0 Then
         psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
         Exit Do
      End If
   Loop
eh:

End Function

'
'   parse list
'
Private Function parseArray(ByRef str As String, ByRef index As Long) As Collection

   Set parseArray = New Collection

   ' "["
   Call skipChar(str, index)
   If Mid(str, index, 1) <> "[" Then
      psErrors = psErrors & "Invalid Array at position " & index & " : " + Mid(str, index, 20) & vbCrLf
      Exit Function
   End If
   
   index = index + 1

   Do

      Call skipChar(str, index)
      If "]" = Mid(str, index, 1) Then
         index = index + 1
         Exit Do
      ElseIf "," = Mid(str, index, 1) Then
         index = index + 1
         Call skipChar(str, index)
      ElseIf index > Len(str) Then
         psErrors = psErrors & "Missing ']': " & Right(str, 20) & vbCrLf
         Exit Do
      End If

      ' add value
      On Error Resume Next
      parseArray.Add parseValue(str, index)
      If Err.Number <> 0 Then
         psErrors = psErrors & Err.Description & ": " & Mid(str, index, 20) & vbCrLf
         Exit Do
      End If
   Loop

End Function

'
'   parse string / number / object / array / true / false / null
'
Private Function parseValue(ByRef str As String, ByRef index As Long)

   Call skipChar(str, index)

   Select Case Mid(str, index, 1)
      Case "{"
         Set parseValue = parseObject(str, index)
      Case "["
         Set parseValue = parseArray(str, index)
      Case """", "'"
         parseValue = parseString(str, index)
      Case "t", "f"
         parseValue = parseBoolean(str, index)
      Case "n"
         parseValue = parseNull(str, index)
      Case Else
         parseValue = parseNumber(str, index)
   End Select

End Function

'
'   parse string
'
Private Function parseString(ByRef str As String, ByRef index As Long) As String

   Dim quote   As String
   Dim Char    As String
   Dim code    As String

   Dim SB As New cStringBuilder

   Call skipChar(str, index)
   quote = Mid(str, index, 1)
   index = index + 1
   
   Do While index > 0 And index <= Len(str)
      Char = Mid(str, index, 1)
      Select Case (Char)
         Case "\"
            index = index + 1
            Char = Mid(str, index, 1)
            Select Case (Char)
               Case """", "\", "/", "'"
                  SB.Append Char
                  index = index + 1
               Case "b"
                  SB.Append vbBack
                  index = index + 1
               Case "f"
                  SB.Append vbFormFeed
                  index = index + 1
               Case "n"
                  SB.Append vbLf
                  index = index + 1
               Case "r"
                  SB.Append vbCr
                  index = index + 1
               Case "t"
                  SB.Append vbTab
                  index = index + 1
               Case "u"
                  index = index + 1
                  code = Mid(str, index, 4)
                  SB.Append ChrW(Val("&h" + code))
                  index = index + 4
            End Select
         Case quote
            index = index + 1
            
            parseString = SB.toString
            Set SB = Nothing
            
            Exit Function
            
         Case Else
            SB.Append Char
            index = index + 1
      End Select
   Loop
   
   parseString = SB.toString
   Set SB = Nothing
   
End Function

'
'   parse number
'
Private Function parseNumber(ByRef str As String, ByRef index As Long)

   Dim value   As String
   Dim Char    As String

   Call skipChar(str, index)
   Do While index > 0 And index <= Len(str)
      Char = Mid(str, index, 1)
      If InStr("+-0123456789.eE", Char) Then
         value = value & Char
         index = index + 1
      Else
         parseNumber = CDec(value)
         Exit Function
      End If
   Loop
End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef str As String, ByRef index As Long) As Boolean

   Call skipChar(str, index)
   If Mid(str, index, 4) = "true" Then
      parseBoolean = True
      index = index + 4
   ElseIf Mid(str, index, 5) = "false" Then
      parseBoolean = False
      index = index + 5
   Else
      psErrors = psErrors & "Invalid Boolean at position " & index & " : " & Mid(str, index) & vbCrLf
   End If

End Function

'
'   parse null
'
Private Function parseNull(ByRef str As String, ByRef index As Long)

   Call skipChar(str, index)
   If Mid(str, index, 4) = "null" Then
      parseNull = Null
      index = index + 4
   Else
      psErrors = psErrors & "Invalid null value at position " & index & " : " & Mid(str, index) & vbCrLf
   End If

End Function

Private Function parseKey(ByRef str As String, ByRef index As Long) As String

   Dim dquote  As Boolean
   Dim squote  As Boolean
   Dim Char    As String

   Call skipChar(str, index)
   Do While index > 0 And index <= Len(str)
      Char = Mid(str, index, 1)
      Select Case (Char)
         Case """"
            dquote = Not dquote
            index = index + 1
            If Not dquote Then
               Call skipChar(str, index)
               If Mid(str, index, 1) <> ":" Then
                  psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                  Exit Do
               End If
            End If
         Case "'"
            squote = Not squote
            index = index + 1
            If Not squote Then
               Call skipChar(str, index)
               If Mid(str, index, 1) <> ":" Then
                  psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                  Exit Do
               End If
            End If
         Case ":"
            index = index + 1
            If Not dquote And Not squote Then
               Exit Do
            Else
               parseKey = parseKey & Char
            End If
         Case Else
            If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
            Else
               parseKey = parseKey & Char
            End If
            index = index + 1
      End Select
   Loop

End Function

'
'   skip special character
'
Private Sub skipChar(ByRef str As String, ByRef index As Long)
   Dim bComment As Boolean
   Dim bStartComment As Boolean
   Dim bLongComment As Boolean
   Do While index > 0 And index <= Len(str)
      Select Case Mid(str, index, 1)
      Case vbCr, vbLf
         If Not bLongComment Then
            bStartComment = False
            bComment = False
         End If
         
      Case vbTab, " ", "(", ")"
         
      Case "/"
         If Not bLongComment Then
            If bStartComment Then
               bStartComment = False
               bComment = True
            Else
               bStartComment = True
               bComment = False
               bLongComment = False
            End If
         Else
            If bStartComment Then
               bLongComment = False
               bStartComment = False
               bComment = False
            End If
         End If
         
      Case "*"
         If bStartComment Then
            bStartComment = False
            bComment = True
            bLongComment = True
         Else
            bStartComment = True
         End If
         
      Case Else
         If Not bComment Then
            Exit Do
         End If
      End Select
      
      index = index + 1
   Loop

End Sub

Public Function toString(ByRef obj As Variant) As String
   Dim SB As New cStringBuilder
   Select Case VarType(obj)
      Case vbNull
         SB.Append "null"
      Case vbDate
         SB.Append """" & CStr(obj) & """"
      Case vbString
         SB.Append """" & jEncode(obj) & """"
      Case vbObject
         
         Dim bFI As Boolean
         Dim i As Long
         
         bFI = True
         If TypeName(obj) = "Dictionary" Then

            SB.Append "{"
            Dim keys
            keys = obj.keys
            For i = 0 To obj.Count - 1
               If bFI Then bFI = False Else SB.Append ","
               Dim key
               key = keys(i)
               SB.Append """" & key & """:" & toString(obj.Item(key))
            Next i
            SB.Append "}"

         ElseIf TypeName(obj) = "Collection" Then

            SB.Append "["
            Dim value
            For Each value In obj
               If bFI Then bFI = False Else SB.Append ","
               SB.Append toString(value)
            Next value
            SB.Append "]"

         End If
      Case vbBoolean
         If obj Then SB.Append "true" Else SB.Append "false"
      Case vbVariant, vbArray, vbArray + vbVariant
         Dim sEB
         SB.Append multiArray(obj, 1, "", sEB)
      Case Else
         SB.Append Replace(obj, ",", ".")
   End Select

   toString = SB.toString
   Set SB = Nothing
   
End Function

Private Function jEncode(str) As String

   Dim SB As New cStringBuilder
   Dim i As Long
   Dim j As Long
   Dim aL1 As Variant
   Dim aL2 As Variant
   Dim c As String
   Dim p As Boolean

   aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
   aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
   For i = 1 To Len(str)
      p = True
      c = Mid(str, i, 1)
      For j = 0 To 7
         If c = Chr(aL1(j)) Then
            SB.Append "\" & Chr(aL2(j))
            p = False
            Exit For
         End If
      Next

      If p Then
         Dim a
         a = AscW(c)
         If a > 31 And a < 127 Then
            SB.Append c
         ElseIf a > -1 Or a < 65535 Then
            SB.Append "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
         End If
      End If
   Next
   
   jEncode = SB.toString
   Set SB = Nothing
   
End Function

Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition
   
   Dim iDU As Long
   Dim iDL As Long
   Dim i As Long
   
   On Error Resume Next
   iDL = LBound(aBD, iBC)
   iDU = UBound(aBD, iBC)

   Dim SB As New cStringBuilder

   Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
   If Err.Number = 9 Then
      sPB1 = sPT & sPS
      For i = 1 To Len(sPB1)
         If i <> 1 Then sPB2 = sPB2 & ","
         sPB2 = sPB2 & Mid(sPB1, i, 1)
      Next
      '        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
      SB.Append toString(aBD(sPB2))
   Else
      sPT = sPT & sPS
      SB.Append "["
      For i = iDL To iDU
         SB.Append multiArray(aBD, iBC + 1, i, sPT)
         If i < iDU Then SB.Append ","
      Next
      SB.Append "]"
      sPT = Left(sPT, iBC - 2)
   End If
   Err.Clear
   multiArray = SB.toString
   
   Set SB = Nothing
End Function


Private Function toUnicode(str As String) As String

   Dim x As Long
   Dim uStr As New cStringBuilder
   Dim uChrCode As Integer

   For x = 1 To Len(str)
      uChrCode = Asc(Mid(str, x, 1))
      Select Case uChrCode
         Case 8:   ' backspace
            uStr.Append "\b"
         Case 9: ' tab
            uStr.Append "\t"
         Case 10:  ' line feed
            uStr.Append "\n"
         Case 12:  ' formfeed
            uStr.Append "\f"
         Case 13: ' carriage return
            uStr.Append "\r"
         Case 34: ' quote
            uStr.Append "\"""
         Case 39:  ' apostrophe
            uStr.Append "\'"
         Case 92: ' backslash
            uStr.Append "\\"
         Case 123, 125:  ' "{" and "}"
            uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
         Case Is < 32, Is > 127: ' non-ascii characters
            uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
         Case Else
            uStr.Append Chr$(uChrCode)
      End Select
   Next
   toUnicode = uStr.toString
   Exit Function

End Function
