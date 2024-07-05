Attribute VB_Name = "FunctionsLLM1"
' License: public domain (https://creativecommons.org/publicdomain/zero/1.0/)
' Kita Toshihiro https://tkita.net 2024
' Mac�� Excel�ł����삵�܂��B���̂��߂ɁA�o�b�N�X���b�V��������
' �_�u���N�H�[�e�[�V���������̒u���������s���ď������Ă��܂��B

Option Explicit

Const GPT_API_URL As String = "https://api.openai.com/v1/chat/completions"
Const DEFAULT_MODEL As String = "gpt-4o"
'Const DEFAULT_MODEL As String = "gpt-3.5-turbo"

Const GEMINI_API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/"
Const GEMINI_DEFAULT_MODEL As String = "gemini-1.5-pro-latest"
'Const GEMINI_DEFAULT_MODEL As String = "gemini-1.5-flash-latest"

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
    
    'API�L�[��api�V�[�g��A2����ǂݍ���
    apiKey = Sheets("api").Range("A2").value

    ' JSON���N�G�X�g���쐬
    json = "{""model"":""" & model & """,""messages"":[{""role"":""user"",""content"":"""
    json = json & EscapeJsonString(prompt) & """}],""max_tokens"":250,""temperature"":0.7}"

#If Mac Then
    Dim http As String
    Dim command As String
    Dim command1 As String
    json = Replace(json, """", DQ_ALT) ' �_�u���N�H�[�g��u������
    command1 = "echo '" & json & "' | perl -pe '$a=chr(34);s/" & DQ_ALT & "/$a/g' | "
    command1 = command1 & "curl '" & GPT_API_URL & "' --header 'Content-Type: application/json' "
    command1 = command1 & " --header 'Authorization: Bearer " & apiKey & "' --data @- -X POST"
    command = "do shell script "" " & command1 & " "" "
    http = MacScript(command)
    response = http
#Else
    Dim http As Object
    ' HTTP�I�u�W�F�N�g (Windows��Excel)
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", GPT_API_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send json
    response = http.responseText
#End If
    If InStr(response, """error""") > 0 Then
        GPT = response '�G���[�̏ꍇ��JSON�����̂܂ܕ\��
    Else
        gptResponse = ParseGPTResponse(response)
        GPT = gptResponse
    End If
End Function

' ���X�|���XJSON���p�[�X����֐�
Private Function ParseGPTResponse(response As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim content As String
    Dim searchString As String
    Dim tmp1 As String
    
    searchString = """content"": """
    'searchString = """content"":"""
    
    ' "content":" �̈ʒu��T��
    startPos = InStr(response, searchString)
    If startPos > 0 Then
        startPos = startPos + Len(searchString)
        endPos = FindUnescapedQuote(response, startPos)
        'endPos = InStr(startPos, response, """")
        If endPos > 0 Then
            ' �R���e���c�𒊏o
            content = Mid(response, startPos, endPos - startPos)
        End If
    End If
   
    tmp1 = UnescapeJsonString(content)
    '�����̉��s���폜
    If Right(tmp1, 1) = Chr(10) Then
        tmp1 = Left(tmp1, Len(tmp1) - 1)
    End If
    ParseGPTResponse = tmp1
End Function

Function GPTrange(rng As Range, Optional model As String = DEFAULT_MODEL) As String
    GPTrange = GPT(RangeToStr(rng), model)
End Function

Function GPTtranslate(prompt As String, Optional model As String = DEFAULT_MODEL, Optional lang As String = "English") As String
    GPTtranslate = GPT(prompt & Chr(92) & "n" & Chr(92) & "n ���̕��͂�" & lang & "�ɖ|�󂵂����� : ", model)
End Function

Function GPTsummary(prompt As String, Optional model As String = DEFAULT_MODEL, Optional length As Integer = 200) As String
    GPTsummary = GPT(prompt & Chr(92) & "n" & Chr(92) & "n ���̕��͂�" & length & "�����ŗv�񂵂����� : ", model)
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
    
    ' API�L�[��api�V�[�g��A3����ǂݍ���
    apiKey = Sheets("api").Range("A3").value
    apiURL = GEMINI_API_URL & model & ":generateContent?key=" & apiKey

    ' JSON���N�G�X�g���쐬
    json = "{""contents"":[{""parts"":[{""text"":""" & EscapeJsonString(prompt) & """}]}]}"

#If Mac Then
    Dim http As String
    Dim command1 As String
    Dim command  As String
    json = Replace(json, """", DQ_ALT) ' �_�u���N�H�[�g��u������
    command1 = "echo '" & json & "' | perl -pe '$a=chr(34);s/" & DQ_ALT & "/$a/g' | "
    command1 = command1 & " curl '" & apiURL & "' --header 'Content-Type: application/json' --data @- -X POST"
    command = "do shell script "" " & command1 & " "" "
    http = MacScript(command)
    response = http
#Else
    Dim http As Object
    ' HTTP�I�u�W�F�N�g���쐬
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", apiURL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send json
    response = http.responseText
#End If
    If InStr(response, """error""") > 0 Then
        Gemini = response '�G���[�̏ꍇ��JSON�����̂܂ܕ\��
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

    ' "text: "�̈ʒu��T��
    startPos = InStr(response, searchString)
    If startPos > 0 Then
        startPos = startPos + Len(searchString)
        endPos = FindUnescapedQuote(response, startPos)
        'endPos = InStr(startPos, response, """")
        If endPos > 0 Then
            ' �R���e���c�𒊏o
            content = Mid(response, startPos, endPos - startPos)
        End If
    End If

    tmp1 = UnescapeJsonString(content)
    tmp1 = Replace(tmp1, Chr(92) & "n", Chr(10))
    '�����̉��s���폜
    If Right(tmp1, 1) = Chr(10) Then
        tmp1 = Left(tmp1, Len(tmp1) - 1)
    End If
    ParseGeminiResponse = tmp1
End Function

Function GeminiRange(rng As Range, Optional model As String = GEMINI_DEFAULT_MODEL) As String
    GeminiRange = Gemini(RangeToStr(rng), model)
End Function

Function GeminiTranslate(prompt As String, Optional model As String = GEMINI_DEFAULT_MODEL, Optional lang As String = "English") As String
    GeminiTranslate = Gemini(prompt & Chr(92) & "n" & Chr(92) & "n ���̕��͂�" & lang & "�ɖ|�󂵂����� : ", model)
End Function

Function GeminiSummary(prompt As String, Optional model As String = GEMINI_DEFAULT_MODEL, Optional length As Integer = 200) As String
    GeminiSummary = Gemini(prompt & Chr(92) & "n" & Chr(92) & "n ���̕��͂�" & length & "�����ŗv�񂵂����� : ", model)
End Function

' --------------------------------------------------------------------------------------

Private Function FindUnescapedQuote(response As String, startPos As Long) As Long
    Dim endP As Long
    Dim startP As Long
    
    startP = startPos
    Do
        ' �R���e���c�̏I���ʒu��T��
        endP = InStr(startP, response, """")
        
        ' �G�X�P�[�v����Ă��Ȃ��_�u���N�H�[�g���������ꍇ�̓��[�v���I��
        If endP = 0 Or Mid(response, endP - 1, 1) <> "\" Then
            Exit Do
        End If
        
        ' ���̈ʒu�Ɉړ����čēx����
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


'�w�肵���͈͂̃Z�����e���}�[�N�_�E���̕\�`���ŕ�����Ƃ��ĕԂ��֐�
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

