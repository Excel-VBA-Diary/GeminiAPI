Attribute VB_Name = "Gemini_API"
Option Explicit

'-------------------------------------------------------------------------------
' Gemini APIを経由して問い合わせた結果を返す.
' Returns the results of a query made via the Gemini API.
'
' Syntax:
'   Gemini(PromptText, [ModelName])
'
' Arguments:
'   PromptText: Required, String
'               Text that describes the task or question for the AI model.
'               PromptText must be a single prompt.
'
'   ModelName: Optional, String
'              Specify the Gemini AI model. If omitted,
'              the default is "emini-3.5-flash-lite".
'
' Return Value:
'   Response (String)
'
' Usage:
'   Dim ans As String
'   ans = Gemini(prompt)
'
' 前提条件
'   コード内の Gemini_API_KEY はご自身で取得したGemini APIキーを割り当てること。
'   APIキーは Google AI Studio で取得する。
'
' Prerequisites
'   Replace “Gemini_API_KEY” in the code with the Gemini API key you obtained.
'   You can obtain the API key from Google AI Studio.
'
'
' Author: Excel VBA Diary (@excelvba_diary)
' Created: September 7, 2026
' Last Updated: September 14, 2026
' Version: 1.001
' License: MIT
'-------------------------------------------------------------------------------

Private Const DefaultModel As String = "gemini-3.5-flash-lite"

Public Function Gemini(PromptText As String, Optional ModelName As String = "") As String
    
    On Error GoTo ErrHandler
    
    If Trim(PromptText) = "" Then
        Gemini = "#ERROR: EmptyPrompt"
        Exit Function
    End If
    
    If Len(Gemini_API_KEY) = 0 Then
        Gemini = "#ERROR: NoApiKey"
        Exit Function
    End If
    
    Dim model As String
    model = IIf(ModelName = "", DefaultModel, ModelName)
    
    ' プロンプト内の特殊文字（\, ", CRLF）をJSON用にエスケープする
    ' Escape special characters (\, ", line breaks) in the prompt for JSON
    
    Dim safePrompt As String
    safePrompt = PromptText
    safePrompt = Replace(safePrompt, "\", "\\")
    safePrompt = Replace(safePrompt, """", "\""")
    safePrompt = Replace(safePrompt, vbCrLf, "\n")
    safePrompt = Replace(safePrompt, vbCr, "\n")
    safePrompt = Replace(safePrompt, vbLf, "\n")
    
    ' JSONペイロードの作成
    ' Creating a JSON Payload
    
    Dim jsonPpayload As String
    jsonPpayload = "{""contents"": [{""parts"":[{""text"": """ & safePrompt & """}]}]}"
    
    Dim api_url As String
    api_url = "https://generativelanguage.googleapis.com/v1beta/models/" & model & ":generateContent"
    
    Dim objHttp As Object
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    With objHttp
        .Open "POST", api_url, False
        .setTimeouts 5000, 5000, 10000, 60000
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "x-goog-api-key", Gemini_API_KEY
        .send StrToUtf8Bytes(jsonPpayload)
        
        If .Status <> 200 Then
            Gemini = "#ERROR:" & .Status & ":" & .responseText
            Debug.Print "POST Error occurred: Status="; .Status
            Debug.Print .responseText
            Exit Function
        End If
            
        Dim jsonResponse As String
        jsonResponse = DecodeUtf8Response(.responseBody)
    End With
    
    Gemini = ExtractJsonValue(jsonResponse, "text")
    Exit Function

ErrHandler:
    Gemini = "#ERROR:Exception:" & Err.Number & ":" & Err.Description
    Debug.Print "Runtime Error occurred: Number="; Err.Number
    Debug.Print Err.Description

End Function
    
    
' JSONのテキストデーターをバイナリーデーターに変換する
' Convert JSON text data to binary data

Private Function StrToUtf8Bytes(ByVal JsonText As String) As Variant
    
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    With objStream
        .Type = 2                   ' adTypeText (Text Data)
        .Charset = "UTF-8"
        .Open
        .WriteText JsonText
        .Position = 0
        .Type = 1                   ' adTypeBinary (Binary Data)
        .Position = 3               ' Since a UTF-8 BOM (EF BB BF) is appended at the beginning, skip 3 bytes.
        StrToUtf8Bytes = .Read      ' An array is returned as a Variant
        .Close
    End With

End Function


' JSONのバイナリーデータをテキストデーターに変換する
' Convert JSON binary data to text data

Private Function DecodeUtf8Response(ByVal JsonBinary As Variant) As String
    
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    With objStream
        .Type = 1                   ' adTypeBinary (Binary Data)
        .Open
        .Write JsonBinary
        .Position = 0
        .Type = 2                   ' adTypeText (Text Data)
        .Charset = "UTF-8"
        DecodeUtf8Response = .ReadText
        .Close
    End With

End Function


' JSONテキストの中からtextキーの値（本文）を抽出する
' ここでは正規表現で抽出しているがJSONパーサーを使ってもよい

' Retrieve the string for a specified key from JSON
' Although this code is using regular expressions for extraction here,
' you can also use a JSON parser.

Function ExtractJsonValue(ByVal JsonText As String, _
                          ByVal keyName As String) As String
    
    Dim matches As Object, strTemp As String
    
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    
    With objRegExp
        .Pattern = """" & keyName & """:\s*?""(.*?)"""
        Set matches = .Execute(JsonText)
        If matches.Count > 0 Then
            strTemp = matches(0).SubMatches(0)
            strTemp = Replace(strTemp, "\""", """")
            strTemp = Replace(strTemp, "\\", "\")
            strTemp = Replace(strTemp, "\n", vbCrLf)
            ExtractJsonValue = strTemp
        Else
            ExtractJsonValue = "#ERROR: TextNotFound"
        End If
    End With

End Function
