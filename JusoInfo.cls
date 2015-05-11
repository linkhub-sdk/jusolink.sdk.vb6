VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JusoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sectionNum    As String
Public roadAddr1     As String
Public roadAddr2     As String
Public jibunAddr     As String
Public zipcode       As String
Public dongCode      As String
Public streetCode    As String
Public detailBuildingName     As Collection
Public relatedJibun     As Collection

Public Sub fromDictionary(dic As Variant)

    sectionNum = dic.Item("sectionNum")
    roadAddr1 = dic.Item("roadAddr1")
    roadAddr2 = dic.Item("roadAddr2")
    jibunAddr = dic.Item("jibunAddr")
    zipcode = dic.Item("zipcode")
    dongCode = dic.Item("dongCode")
    streetCode = dic.Item("streetCode")
        
    If IsEmpty(dic.Item("detailBuildingName")) = False Then
        Set detailBuildingName = dic.Item("detailBuildingName")
    End If

    If IsEmpty(dic.Item("relatedJibun")) = False Then
        Set relatedJibun = dic.Item("relatedJibun")
    End If
           
End Sub
