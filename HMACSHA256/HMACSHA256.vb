Imports System.Text
Imports System.Security.Cryptography

<ComClass(HMACSHA256fork.ClassId, HMACSHA256fork.InterfaceId, HMACSHA256fork.EventsId)> Public Class HMACSHA256fork
    Public Const ClassId As String = "9468DB22-4BE0-41E6-B2A2-7373D4D9E506"
    Public Const InterfaceId As String = "E5AEA0C5-21CD-477E-A7C7-0ADA7D1C5694"
    Public Const EventsId As String = "4272F0BF-2E69-43A8-9CE7-59C32B7611C2"

    'コンストラクタ 必須
    Public Sub New()
        MyBase.New()
    End Sub

    '公開するファンクション
    Public Function CalcHash(ByVal sOut As String, ByVal sKey As String, ByVal bUrlEncode As Boolean) As String

        Dim sHash As String

        sHash = CalcHash(sOut, sKey)

        If bUrlEncode Then
            '署名をURLエンコード
            CalcHash = Uri.EscapeDataString(sHash)
        Else
            CalcHash = sHash
        End If

    End Function

    Public Function CalcHash(ByVal sOut As String, ByVal sKey As String) As String
        Dim hash() As Byte
        Dim bSecret() As Byte
        Dim bRequest() As Byte
        Dim objUTF8 As New System.Text.UTF8Encoding
        Dim strBase64 As String

        'UTF8変換
        bRequest = Encoding.UTF8.GetBytes(sOut)
        bSecret = Encoding.UTF8.GetBytes(sKey)

        '暗号キー付で呼び出す
        Using objSHA256 As New System.Security.Cryptography.HMACSHA256()

            '暗号キー設定
            objSHA256.Key = bSecret

            'ハッシュ値を計算（バイナリ）
            objSHA256.ComputeHash(bRequest)
            hash = objSHA256.Hash

        End Using

        'Base64でエンコードを行う
        strBase64 = System.Convert.ToBase64String(hash)

        '// 結果を返す
        CalcHash = strBase64
    End Function
End Class