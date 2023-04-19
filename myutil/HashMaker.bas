Attribute VB_Name = "HashMaker"
Option Explicit

Public Function md5HashMaker(str As String) As String
    Dim md5 As Object
    Dim utf8 As Object
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim i As Integer
    Dim res As String

    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    bytes = utf8.GetBytes_4(str)

    Set md5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    hash = md5.ComputeHash_2(bytes)

    For i = LBound(hash) To UBound(hash)
        res = res & LCase(Right("0" & Hex(hash(i)), 2))
    Next i

    md5HashMaker = LCase(res)
End Function


Public Function sha256HashMaker(str As String) As String
    Dim sha256m As Object
    Dim utf8 As Object
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim i As Integer
    Dim res As String

    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    bytes = utf8.GetBytes_4(str)

    Set sha256m = CreateObject("System.Security.Cryptography.SHA256Managed")
    hash = sha256m.ComputeHash_2((bytes))

    For i = LBound(hash) To UBound(hash)
        res = res & LCase(Right("0" & Hex(hash(i)), 2))
    Next i

    HASH_SHA256 = LCase(res)
End Function
