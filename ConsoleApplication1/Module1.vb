Module Module1

    Sub Main()

        'デバッグ用データ設定
        'strRequest = "GET" & vbLf
        'strRequest = strRequest & "webservices.amazon.com" & vbLf
        'strRequest = strRequest & "/onca/xml" & vbLf
        'strRequest = strRequest & "AWSAccessKeyId=00000000000000000000"
        'strRequest = strRequest & "&ItemId=0679722769"
        'strRequest = strRequest & "&Operation=ItemLookup"
        'strRequest = strRequest & "&ResponseGroup=ItemAttributes%2COffers%2CImages%2CReviews"
        'strRequest = strRequest & "&Service=AWSECommerceService"
        'strRequest = strRequest & "&Timestamp=2009-01-01T12%3A00%3A00Z"
        'strRequest = strRequest & "&Version=2009-01-06"
        'strSecKey = "1234567890"

        'リクエストの作成
        'パラメータ部分
        Dim strParam As String
        strParam = "AWSAccessKeyId=AKIAIW4ZB3OTQA7YVJJQ"
        strParam = strParam & "&AssociateTag=gntz0220alc89-22"
        strParam = strParam & "&Operation=ItemSearch"
        strParam = strParam & "&SearchIndex=Books"
        strParam = strParam & "&Service=AWSECommerceService"
        strParam = strParam & "&Timestamp=2017-02-21T02%3A08%3A30.000Z"
        strParam = strParam & "&Title=Harry%20Potter"

        Dim strSigMake As String
        strSigMake = "GET" & vbLf
        strSigMake = strSigMake & "ecs.amazonaws.jp" & vbLf
        strSigMake = strSigMake & "/onca/xml" & vbLf
        strSigMake = strSigMake & strParam

        Dim strRequest As String
        strRequest = "http://ecs.amazonaws.jp/onca/xml?"
        strRequest = strRequest & strParam

        Dim strSecKey As String
        strSecKey = "LjJHhxF5GzlZwwhH7Ar22p1ZzpXUjMMmcMbE21cK"

        'HMAC with the SHA256ハッシュアルゴリズムを使って計算し署名を作成(暗号キーにはシークレットキーを使用)
        'テスト呼び出し
        Dim objHmacSha256 As New HMACSHA256.HMACSHA256fork
        Dim sHash As String

        'sHash = objHmacSha256.CalcHash(strSigMake, strSecKey)
        'URLエンコード済みで取得する
        sHash = objHmacSha256.CalcHash(strSigMake, strSecKey, True)

        '署名をURLエンコード
        'URLエンコードを行う
        'Dim strUrlEnc As String
        'strUrlEnc = Uri.EscapeDataString(sHash)

        'シグネチャパラメータの付与
        'strRequest = strRequest & "&Signature=" & strUrlEnc
        strRequest = strRequest & "&Signature=" & sHash

    End Sub

    Function GetTimeStamp() As String

        GetTimeStamp = Format(Now(), "yyyy-mm-ddThh:mm:ssZ")

    End Function

End Module
