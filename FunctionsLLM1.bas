Attribute VB_Name = "FunctionsLLM1"
' License: public domain (https://creativecommons.org/publicdomain/zero/1.0/)
' Kita Toshihiro https://tkita.net 2024
' Mac版 Excelでも動作します。そのために、バックスラッシュ文字や
' ダブルクォーテーション文字の置き換えを行って処理しています。

Option Explicit

Const GPT_API_URL As String = "https://api.openai.com/v1/chat/completions"
'Const DEFAULT_MODEL As String = "gpt-4o"
Const DEFAULT_MODEL As String = "gpt-4o-mini"
'https://openai.com/api/pricing/

Const GEMINI_API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/"
'Const GEMINI_DEFAULT_MODEL As String = "gemini-1.5-pro-latest"
Const GEMINI_DEFAULT_MODEL As String = "gemini-1.5-flash-latest"
'https://ai.google.dev/pricing

Const DQ_ALT As String = "__@@" ' for MAC Excel

' --------------------- OpenAI GPT -------------------------------

Function GPT(prompt As String, Optional model As String = DEFAULT_MODEL) As String
    Dim json As String
    Dim response As String
    Dim apiKey As String
    Dim gptResponse As String
    
    If model = "" Then
        model = DEFAULT_MODEL
    End If
    
    'APIキーをapiシートのA2から読み込む
    apiKey = Sheets("api").Range("A2").value
    
    'max_tokens で指定すると、文章の途中で突然切れた出力になることが多い。
    prompt = "300文字以内で出力してください。" & prompt

    ' JSONリクエストを作成
    json = "{""model"":""" & model & """,""messages"":[{""role"":""user"",""content"":"""
    json = json & EscapeJsonString(prompt) & """}],""max_tokens"":350,""temperature"":0.7}"

#If Mac Then
    Dim http As String
    Dim command As String
    Dim command1 As String
    json = Replace(json, """", DQ_ALT) ' ダブルクォートを置き換え
    command1 = "echo '" & json & "' | perl -pe '$a=chr(34);s/" & DQ_ALT & "/$a/g' | "
    command1 = command1 & "curl '" & GPT_API_URL & "' --header 'Content-Type: application/json' "
    command1 = command1 & " --header 'Authorization: Bearer " & apiKey & "' --data @- -X POST"
    command = "do shell script "" " & command1 & " "" "
    http = MacScript(command)
    response = http
#Else
    Dim http As Object
    ' HTTPオブジェクト (Windows版Excel)
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", GPT_API_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send json
    response = http.responseText
#End If
    If InStr(response, """error""") > 0 Then
        GPT = response 'エラーの場合はJSONをそのまま表示
    Else
        gptResponse = ParseGPTResponse(response)
        GPT = gptResponse
    End If
End Function

' レスポンスJSONをパースする関数
Private Function ParseGPTResponse(response As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim content As String
    Dim searchString As String
    Dim tmp1 As String
    
    searchString = """content"": """
    'searchString = """content"":"""
    
    ' "content":" の位置を探す
    startPos = InStr(response, searchString)
    If startPos > 0 Then
        startPos = startPos + Len(searchString)
        endPos = FindUnescapedQuote(response, startPos)
        'endPos = InStr(startPos, response, """")
        If endPos > 0 Then
            ' コンテンツを抽出
            content = Mid(response, startPos, endPos - startPos)
        End If
    End If
   
    tmp1 = UnescapeJsonString(content)
    '末尾の改行を削除
    If Right(tmp1, 1) = Chr(10) Then
        tmp1 = Left(tmp1, Len(tmp1) - 1)
    End If
    ParseGPTResponse = tmp1
End Function

Function GPTrange(rng As Range, Optional model As String = DEFAULT_MODEL) As String
    GPTrange = GPT(RangeToStr(rng), model)
End Function

Function GPTtranslate(prompt As String, Optional lang As String = "English", Optional model As String = DEFAULT_MODEL) As String
    GPTtranslate = GPT(prompt & Chr(92) & "n" & Chr(92) & "n この文章を" & lang & "に翻訳したもの : ", model)
End Function

Function GPTsummary(prompt As String, Optional length As Integer = 150, Optional model As String = DEFAULT_MODEL) As String
    GPTsummary = GPT(prompt & Chr(92) & "n" & Chr(92) & "n この文章を" & length & "文字で要約したもの : ", model)
End Function


' --------------------- Google Gemini -------------------------------

Function Gemini(prompt As String, Optional model As String = GEMINI_DEFAULT_MODEL) As String
    Dim json As String
    Dim response As String
    Dim apiKey As String
    Dim geminiResponse As String
    Dim apiURL As String
    
    If model = "" Then
        model = GEMINI_DEFAULT_MODEL
    End If
    
    ' APIキーをapiシートのA3から読み込む
    apiKey = Sheets("api").Range("A3").value
    apiURL = GEMINI_API_URL & model & ":generateContent?key=" & apiKey
    
    'max_tokens で指定すると、文章の途中で突然切れた出力になることが多い。
    prompt = "300文字以内で出力してください。" & prompt

    ' JSONリクエストを作成
    json = "{""contents"":[{""parts"":[{""text"":""" & EscapeJsonString(prompt) & """}]}]}"

#If Mac Then
    Dim http As String
    Dim command1 As String
    Dim command  As String
    json = Replace(json, """", DQ_ALT) ' ダブルクォートを置き換え
    command1 = "echo '" & json & "' | perl -pe '$a=chr(34);s/" & DQ_ALT & "/$a/g' | "
    command1 = command1 & " curl '" & apiURL & "' --header 'Content-Type: application/json' --data @- -X POST"
    command = "do shell script "" " & command1 & " "" "
    http = MacScript(command)
    response = http
#Else
    Dim http As Object
    ' HTTPオブジェクトを作成
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", apiURL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send json
    response = http.responseText
#End If
    If InStr(response, """error""") > 0 Then
        Gemini = response 'エラーの場合はJSONをそのまま表示
    Else
        geminiResponse = ParseGeminiResponse(response)
        Gemini = geminiResponse
    End If
End Function

Private Function ParseGeminiResponse(response As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim content As String
    Dim searchString As String
    Dim tmp1 As String
    
    searchString = """text"": """

    ' "text: "の位置を探す
    startPos = InStr(response, searchString)
    If startPos > 0 Then
        startPos = startPos + Len(searchString)
        endPos = FindUnescapedQuote(response, startPos)
        'endPos = InStr(startPos, response, """")
        If endPos > 0 Then
            ' コンテンツを抽出
            content = Mid(response, startPos, endPos - startPos)
        End If
    End If

    tmp1 = UnescapeJsonString(content)
    tmp1 = Replace(tmp1, Chr(92) & "n", Chr(10))
    '末尾の改行を削除
    If Right(tmp1, 1) = Chr(10) Then
        tmp1 = Left(tmp1, Len(tmp1) - 1)
    End If
    ParseGeminiResponse = tmp1
End Function

Function GeminiRange(rng As Range, Optional model As String = GEMINI_DEFAULT_MODEL) As String
    GeminiRange = Gemini(RangeToStr(rng), model)
End Function

Function GeminiTranslate(prompt As String, Optional lang As String = "English", Optional model As String = GEMINI_DEFAULT_MODEL) As String
    GeminiTranslate = Gemini(prompt & Chr(92) & "n" & Chr(92) & "n この文章を" & lang & "に翻訳したもの : ", model)
End Function

Function GeminiSummary(prompt As String, Optional length As Integer = 150, Optional model As String = GEMINI_DEFAULT_MODEL) As String
    GeminiSummary = Gemini(prompt & Chr(92) & "n" & Chr(92) & "n この文章を" & length & "文字で要約したもの : ", model)
End Function

' --------------------------------------------------------------------------------------

Private Function FindUnescapedQuote(response As String, startPos As Long) As Long
    Dim endP As Long
    Dim startP As Long
    
    startP = startPos
    Do
        ' コンテンツの終了位置を探す
        endP = InStr(startP, response, """")
        
        ' エスケープされていないダブルクォートを見つけた場合はループを終了
        If endP = 0 Or Mid(response, endP - 1, 1) <> "\" Then
            Exit Do
        End If
        
        ' 次の位置に移動して再度検索
        startP = endP + 1
    Loop While endP > 0
    
    FindUnescapedQuote = endP
End Function

Function EscapeJsonString(value As String) As String
    value = Replace(value, Chr(92), Chr(92) & Chr(92)) ' \ -> \\
    value = Replace(value, Chr(34), Chr(92) & Chr(34)) ' " -> \"
    value = Replace(value, vbCrLf, Chr(92) & "n")      ' CRLF -> \n
    value = Replace(value, vbCr, Chr(92) & "n")        ' CR -> \n
    value = Replace(value, vbLf, Chr(92) & "n")        ' LF -> \n
    EscapeJsonString = value
End Function

Function UnescapeJsonString(escapedStr As String) As String
    Dim result As String
    result = escapedStr
    result = Replace(result, Chr(92) & Chr(34), Chr(34))  ' Unescape \"
    result = Replace(result, Chr(92) & Chr(92), Chr(92))  ' Unescape \\
    result = Replace(result, Chr(92) & "b", Chr(8))       ' Unescape \b
    result = Replace(result, Chr(92) & "f", Chr(12))      ' Unescape \f
    result = Replace(result, Chr(92) & "n", Chr(10))      ' Unescape \n
    result = Replace(result, ChrW(92) & "n", Chr(10))     ' Unescape \n
    result = Replace(result, ChrW(165) & "n", Chr(10))    ' Unescape \n
    result = Replace(result, Chr(92) & "r", Chr(13))      ' Unescape \r
    result = Replace(result, Chr(92) & "t", Chr(9))       ' Unescape \t
    UnescapeJsonString = result
End Function


'指定した範囲のセル内容をマークダウンの表形式で文字列として返す関数
Function RangeToStr(rng As Range) As String
    Dim cell As Range
    Dim row As Range
    Dim result As String
    Dim rowStr As String
    Dim colCount As Integer
    Dim rowCount As Integer
    Dim i As Integer
    
    colCount = rng.Columns.Count
    rowCount = rng.Rows.Count
    
    ' Create the header row for the markdown table
    For i = 1 To colCount
        If i = 1 Then
            rowStr = "| "
        Else
            rowStr = rowStr & " | "
        End If
        rowStr = rowStr & rng.Cells(1, i).value
    Next i
    rowStr = rowStr & " |" & Chr(92) & "n"
    result = result & rowStr
    
    ' Create the separator row for the markdown table
    rowStr = "|"
    For i = 1 To colCount
        rowStr = rowStr & " --- |"
    Next i
    rowStr = rowStr & Chr(92) & "n"
    result = result & rowStr
    
    ' Loop through each row in the range
    For Each row In rng.Rows
        rowStr = "|"
        ' Loop through each cell in the row
        For Each cell In row.Cells
            rowStr = rowStr & " " & cell.value & " |"
        Next cell
        result = result & rowStr & Chr(92) & "n"
    Next row
    
    RangeToStr = result
End Function

