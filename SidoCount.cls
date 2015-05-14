VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SidoCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GYEONGGI         As Integer
Public GYEONGSANGBUK    As Integer
Public GYEONGSANGNAM    As Integer
Public SEOUL            As Integer
Public JEOLLANAM        As Integer
Public CHUNGCHEONGNAM   As Integer
Public JEOLLABUK        As Integer
Public BUSAN            As Integer
Public GANGWON          As Integer
Public CHUNGCHEONGBUK   As Integer
Public DAEGU            As Integer
Public INCHEON          As Integer
Public GWANGJU          As Integer
Public JEJU             As Integer
Public DAEJEON          As Integer
Public ULSAN            As Integer
Public SEJONG           As Integer

Public Sub fromDictionary(dic As Variant)

    If IsEmpty(dic.Item("GYEONGGI")) = False Then
        GYEONGGI = dic.Item("GYEONGGI")
    End If
    
    If IsEmpty(dic.Item("GYEONGSANGBUK")) = False Then
        GYEONGSANGBUK = dic.Item("GYEONGSANGBUK")
    End If
    
    If IsEmpty(dic.Item("GYEONGSANGNAM")) = False Then
        GYEONGSANGNAM = dic.Item("GYEONGSANGNAM")
    End If
    
    If IsEmpty(dic.Item("SEOUL")) = False Then
        SEOUL = dic.Item("SEOUL")
    End If
    
    If IsEmpty(dic.Item("JEOLLANAM")) = False Then
        JEOLLANAM = dic.Item("JEOLLANAM")
    End If
    
    If IsEmpty(dic.Item("CHUNGCHEONGNAM")) = False Then
        CHUNGCHEONGNAM = dic.Item("CHUNGCHEONGNAM")
    End If
    
    If IsEmpty(dic.Item("JEOLLABUK")) = False Then
        JEOLLABUK = dic.Item("JEOLLABUK")
    End If
    
    If IsEmpty(dic.Item("BUSAN")) = False Then
        BUSAN = dic.Item("BUSAN")
    End If
    
    If IsEmpty(dic.Item("GANGWON")) = False Then
        GANGWON = dic.Item("GANGWON")
    End If
    
    If IsEmpty(dic.Item("CHUNGCHEONGBUK")) = False Then
        CHUNGCHEONGBUK = dic.Item("CHUNGCHEONGBUK")
    End If
    
    If IsEmpty(dic.Item("DAEGU")) = False Then
        DAEGU = dic.Item("DAEGU")
    End If
    
    If IsEmpty(dic.Item("INCHEON")) = False Then
        INCHEON = dic.Item("INCHEON")
    End If
    
    If IsEmpty(dic.Item("GWANGJU")) = False Then
        GWANGJU = dic.Item("GWANGJU")
    End If
    
    If IsEmpty(dic.Item("JEJU")) = False Then
        JEJU = dic.Item("JEJU")
    End If
    
    If IsEmpty(dic.Item("DAEJEON")) = False Then
        DAEJEON = dic.Item("DAEJEON")
    End If
    If IsEmpty(dic.Item("ULSAN")) = False Then
        ULSAN = dic.Item("ULSAN")
    End If
    
    If IsEmpty(dic.Item("SEJONG")) = False Then
        SEJONG = dic.Item("SEJONG")
    End If
           
End Sub

