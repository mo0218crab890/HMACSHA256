Imports System
Imports System.IO
Imports System.Security.Cryptography

Public Class HMACSHA256example

    Public Shared Sub Main(ByVal Fileargs() As String)
        Dim strRequest As String
        Dim strSecKey As String

        strRequest = "GET" & vbCr
        strRequest = strRequest & "ecs.amazonaws.jp" & vbCr
        strRequest = strRequest & "/onca/xml" & vbCr
        strRequest = strRequest & "AWSAccessKeyId=AKIAIW4ZB3OTQA7YVJJQ" & "&AssociateTag=gntz0220alc89-22&" & vbCr
        strRequest = strRequest & "Operation=ItemSearch" & "&SearchIndex=Books" & "&Service=AWSECommerceService&" & vbCr
        strRequest = strRequest & "2017-02-21T02%3A08%3A30.000Z&" & "&Title=Harry%20Potter"
        strSecKey = "LjJHhxF5GzlZwwhH7Ar22p1ZzpXUjMMmcMbE21cK"

        'テスト呼び出し
        Dim oCrypto As New HMACSHA256.HMACSHA256fork
        Dim sCrypt As String

        sCrypt = oCrypto.CalcHash(strRequest, strSecKey)

    End Sub 'Main

End Class 'HMACSHA256example 'end VerifyFile
