VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Jusolink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=================================================================================
' Class Module for base module for Jusolink API SDK. It include base functionality for
' RESTful web service request and parse json result. It uses Linkhub class module
' to accomplish authentication APIs.
'
' This module uses advapi32.dll and crypt32.dll. most of windows system has it
' on system32 by default.
'
' http://www.jusolink.com
' Author : Jeong Yoohan (yhjeong@linkhub.co.kr)
' Written : 2015-05-07
'
' Thanks for your interest.
'=================================================================================
Private Const ServiceID = "JUSOLINK"
Private Const ServiceURL = "https://juso.linkhub.co.kr"
Private Const APIVersion = "1.0"

Private Const adTypeBinary As Long = 1
Private Const adTypeText As Long = 2
Private Const adModeReadWrite As Long = 3

Private m_token_Dic
Private m_Linkhub As Linkhub
Private m_scope As New Collection
Private m_LastErrCode As Long
Private m_LastErrMessage As String

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Sub Class_Initialize()
    Set m_token_Dic = CreateObject("Scripting.Dictionary")
End Sub
Private Sub Class_Terminate()
     m_token_Dic.RemoveAll
    Set m_token_Dic = Nothing
End Sub
Private Function UniStrToUTF8(UniString) As Byte()
   ' Convert a Unicode string to a byte stream of UTF-8
   Dim BArray() As Byte
   Dim TempB() As Byte
   Dim i As Long
   Dim k As Long
   Dim TLen As Long
   Dim b1 As Byte
   Dim b2 As Byte
   Dim UTF16 As Long
   Dim j
   TLen = Len(UniString)  ' Obtain length of Unicode input string
   If TLen = 0 Then Exit Function  ' get out if there's nothing to convert
   k = 0
   For i = 1 To TLen
      ' Work out the UTF16 value of the Unicode character
      CopyMemory b1, ByVal StrPtr(UniString) + ((i - 1) * 2), 1
      CopyMemory b2, ByVal StrPtr(UniString) + ((i - 1) * 2) + 1, 1
      ' Combine the 2 bytes into the Unicode UTF-16
      UTF16 = b2  ' assign b2 to UTF16 before multiplying by 256 to avoid overflow
      UTF16 = UTF16 * 256 + b1
      ' Convert UTF-16 to 2 or 3 bytes of UTF-8
      TempB = ToUTF8(UTF16)
      ' Copy the resultant bytes to BArray
      For j = 0 To UBound(TempB)
         ReDim Preserve BArray(k)
         BArray(k) = TempB(j): k = k + 1
      Next
      ReDim TempB(0)
   Next
   UniStrToUTF8 = BArray  ' Return the resultant UTF-8 byte array
End Function

Private Function ToUTF8(ByVal UTF16 As Long) As Byte()
   ' Convert a 16bit UTF-16BE to 2 or 3 UTF-8 bytes
   Dim BArray() As Byte
   If UTF16 < &H80 Then
      ReDim BArray(0)  ' one byte UTF-8
      BArray(0) = UTF16  ' Use number as is
   ElseIf UTF16 < &H800 Then
      ReDim BArray(1)  ' two byte UTF-8
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(0) = &HC0 + (UTF16 And &H1F)  ' Use 5 remaining bits
   Else
      ReDim BArray(2)  ' three byte UTF-8
      BArray(2) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Use next 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits again
      BArray(0) = &HE0 + (UTF16 And &HF)  ' Use 4 remaining bits
   End If
   ToUTF8 = BArray  ' Return UTF-8 bytes in an array
End Function
'최근 오류코드 확인
Public Property Get LastErrCode() As Long
    LastErrCode = m_LastErrCode
End Property
'최근오류메시지 확인
Public Property Get LastErrMessage() As String
    LastErrMessage = m_LastErrMessage
End Property
'초기화
Public Sub Initialize(linkID As String, SecretKey As String)
    Set m_Linkhub = New Linkhub
    m_Linkhub.linkID = linkID
    m_Linkhub.SercetKey = SecretKey
       
    m_scope.Add "200"
End Sub
Private Function getSession_token() As String
    Dim refresh As Boolean
    
    refresh = False
    
    Dim m_Token As LinkhubToken
    
    Set m_Token = Nothing
    
    If m_token_Dic.Exists("token") Then
        Set m_Token = m_token_Dic.Item("token")
    End If
    
    If m_Token Is Nothing Then
        refresh = True
    Else
        Dim utcnow As String
        utcnow = Format(m_Linkhub.UTCTime, "yyyy-MM-ddTHH:mm:ss.sssZ")
        refresh = m_Token.expiration < utcnow
    End If
    
    If refresh Then
    
        Set m_Token = m_Linkhub.getToken(ServiceID, "", m_scope)
        
        If m_Token Is Nothing Then
            Err.Raise m_Linkhub.LastErrCode, "JUSOLINK", m_Linkhub.LastErrMessage
            Exit Function
        End If
        
        m_token_Dic.Add "token", m_Token
        
    End If
    
    getSession_token = m_Token.session_token

End Function


Private Function httpGET(url As String, BearerToken As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("GET", ServiceURL + url)
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-api-version", APIVersion)
    
    
    winhttp1.send
    winhttp1.WaitForResponse
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "JUSOLINK", parsedDic.Item("message")
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpGET = m_Linkhub.parse(Response)

End Function
 
Public Function URLEncode(ByVal StringToEncode As String) As String
   Dim i                As Integer
   Dim iAsc             As Long
   Dim sTemp            As String
   
   Dim ByteArrayToEncode() As Byte
 
   ByteArrayToEncode = ADO_EncodeUTF8(StringToEncode)
   
   For i = 0 To UBound(ByteArrayToEncode)
      iAsc = ByteArrayToEncode(i)
      Select Case iAsc
         Case 32 'space
            sTemp = "+"
         Case 48 To 57, 65 To 90, 97 To 122
            sTemp = Chr(ByteArrayToEncode(i))
         Case Else
            Debug.Print iAsc
            sTemp = "%" & Hex(iAsc)
      End Select
      URLEncode = URLEncode & sTemp
   Next
 
End Function
  
'Purpose: UTF16 to UTF8 using ADO
Public Function ADO_EncodeUTF8(ByVal strUTF16 As String) As Byte()
 
   Dim objStream        As Object
   Dim data()           As Byte
 
   Set objStream = CreateObject("ADODB.Stream")
   objStream.Charset = "utf-8"
   objStream.Mode = adModeReadWrite
   objStream.Type = adTypeText
   objStream.Open
   objStream.WriteText strUTF16
   objStream.Flush
   objStream.Position = 0
   objStream.Type = adTypeBinary
   objStream.Read 3 ' skip BOM
   data = objStream.Read()
   objStream.Close
   ADO_EncodeUTF8 = data
 
End Function

'주소검색
Public Function Search(Index As String, PageNum As Integer, PerPage As Integer, noDiff As Boolean, noSuggest As Boolean) As SearchResult
    On Error GoTo ErrHandler
    
    Dim Result As Variant
    Dim url As String
    
    If PageNum < 1 Then PageNum = Null
    If PerPage <> 0 Then
        If PerPage < 0 Then PerPage = 20
    End If
    
    url = "/Search?Searches=" + URLEncode(Index)
    
    If PageNum <> 0 Then
        url = url + "&PageNum=" + CStr(PageNum)
    End If
    
    If PerPage <> 0 Then
        url = url + "&PerPage=" + CStr(PerPage)
    End If
    
    If noSuggest Then
        url = url + "&noSuggest=true"
    End If
    
    If noDiff Then
        url = url + "&noDifferential=true"
    End If
    
    Set Result = httpGET(url, getSession_token())
    
    Set Search = New SearchResult
    
    Search.fromDictionary Result
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set Search = Nothing
End Function

'잔액조회
Public Function GetBalance() As Double
    On Error GoTo ErrHandler
    
    GetBalance = m_Linkhub.GetPartnerBalance(getSession_token(), ServiceID)
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetBalance = m_LastErrCode
End Function

''단가확인
Public Function GetUnitCost() As Single
    On Error GoTo ErrHandler
    Dim Result As Variant
        
    Set Result = httpGET("/Search/UnitCost", getSession_token())
    
    GetUnitCost = Result.Item("unitCost")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetUnitCost = m_LastErrCode

End Function



