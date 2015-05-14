VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SearchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public searches     As String
Public deletedWord  As Collection
Public suggest      As String
Public numFound     As String
Public listSize     As String
Public totalPage    As String
Public page         As String
Public chargeYN     As Boolean
Public juso         As Collection
Public sidoCount    As sidoCount


Public Sub fromDictionary(dic As Variant)
    searches = dic.Item("searches")
    suggest = dic.Item("suggest")
    numFound = dic.Item("numFound")
    listSize = dic.Item("listSize")
    totalPage = dic.Item("totalPage")
    page = dic.Item("page")
    chargeYN = dic.Item("chargeYN")
    
    If IsEmpty(dic.Item("sidoCount")) = False Then
    
        Set sidoCount = New sidoCount
        sidoCount.fromDictionary dic.Item("sidoCount")
    End If
    
        
    If IsEmpty(dic.Item("deletedWord")) = False Then
           Set deletedWord = dic.Item("deletedWord")
    End If
        
    
    If IsEmpty(dic.Item("juso")) = False Then
        Dim tCollection As Collection
        Set tCollection = dic.Item("juso")
    
        If (tCollection Is Nothing) = False Then
            Set juso = New Collection
            Dim T As Variant
            Dim info As JusoInfo
            
            For Each T In tCollection
                Set info = New JusoInfo
                info.fromDictionary T
                juso.Add info
            Next
        End If
                
    End If
    
End Sub
