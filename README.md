# GeminiAPI
Gemini APIを経由して問い合わせた結果を返す. (Returns the results of a query made via the Gemini API.)

2026年9月14日以降、COPILOT機能が使用できなくなるというので代替手段としてGemini関数を作ってみた。  
  
前提条件  
・コード内のAPI_KEYはご自身で取得したGemini APIキーを割り当てること。  
・APIキーはGoogle AI Studioで取得する。  

使用例
```
Dim ans As String
ans = Gemini("日本で一番高い山は？")
```

VBAコード
```
Option Explicit

'-------------------------------------------------------------------------------
' Gemini APIを経由して問い合わせた結果を返す.
' Returns the results of a query made via the Gemini API.
'
' Arguments:
'   PromptText: Required, String
'               Text that describes the task or question for the AI model.
'               PromptText must be a single prompt.
'
'   ModelName: Optional, String
'              Specify the Gemini AI model. If omitted,
'              the default is “gemini-3.5-flash-lite.”
'
' Return Value:
'   Response (String)
'
' Usage:
'   Dim ans As String
'   ans = Gemini(prompt)
'
' 前提条件
'   コード内のAPI_KEYはご自身で取得したGemini APIキーを割り当てること。
'   APIキーはGoogle AI Studioで取得する。
'
' Prerequisites
'   Replace “API_KEY” in the code with the Gemini API key you obtained.
'   You can obtain the API key from Google AI Studio.
'
'
' Author: Excel VBA Diary (@excelvba_diary)
' Created: September 7, 2026
' Last Updated: September 7, 2026
' Version: 1.000
' License: MIT
'-------------------------------------------------------------------------------

Private Const DefaultModel As String = "gemini-3.5-flash-lite"

Public Function Gemini(PromptText As String, Optional ModelName As String = "") As String
    
    On Error GoTo ErrHandler
    
    If Trim(PromptText) = "" Then
        Gemini = "#ERROR:EmptyPrompt"
        Exit Function
    End If
    
    If Len(API_KEY) = 0 Then
        Gemini = "#ERROR:NoApiKey"
        Exit Function
    End If
    
    Dim model As String
    model = IIf(ModelName = "", DefaultModel, ModelName)
    
    
    ' プロンプト内の特殊文字（\, ", 改行）をJSON用にエスケープ処理
    Dim safePrompt As String
    safePrompt = PromptText
    safePrompt = Replace(safePrompt, "\", "\\")
    safePrompt = Replace(safePrompt, """", "\""")
    safePrompt = Replace(safePrompt, vbCrLf, "\n")
    safePrompt = Replace(safePrompt, vbCr, "\n")
    safePrompt = Replace(safePrompt, vbLf, "\n")
    
    ' JSONペイロードの作成
    Dim jsonPpayload As String
    jsonPpayload = "{""contents"": [{""parts"":[{""text"": """ & safePrompt & """}]}]}"
    
    Dim api_url As String
    api_url = "https://generativelanguage.googleapis.com/v1beta/models/" & model & ":generateContent"
    
    Dim objHttp As Object
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    With objHttp
        .Open "POST", api_url, False
        .setTimeouts 5000, 5000, 10000, 30000
        .SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        .SetRequestHeader "x-goog-api-key", API_KEY
        .Send StrToUtf8Bytes(jsonPpayload)
        
        If .Status <> 200 Then
            Gemini = "#ERROR:" & .Status & ":" & .responseText
            Debug.Print "エラーが発生しました: Status="; .Status
            Debug.Print .responseText
            Exit Function
        End If
            
        Dim jsonResponse As String
        jsonResponse = DecodeUtf8Response(.responseBody)
    End With
    
    Gemini = ExtractTextFromJson(jsonResponse)
    Exit Function

ErrHandler:
    Gemini = "#ERROR:Exception:" & Err.Number & ":" & Err.Description
    Debug.Print "実行時エラーが発生しました: Number="; Err.Number
    Debug.Print Err.Description

End Function
    
    
' JSONのテキストデーターをバイナリーデーターに変換する
Private Function StrToUtf8Bytes(ByVal JsonText As String) As Variant
    
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    With objStream
        .Type = 2                   ' adTypeText (テキストデーター)
        .Charset = "UTF-8"
        .Open
        .WriteText JsonText
        .Position = 0
        .Type = 1                   ' adTypeBinary (バイナリデーター)
        .Position = 3               ' UTF-8 BOM (EF BB BF)が付くので3バイト分スキップ
        StrToUtf8Bytes = .Read      ' Variant型として配列が返る
        .Close
    End With

End Function


' JSONのバイナリーデータをテキストデーターに変換する
Private Function DecodeUtf8Response(ByVal JsonBinary As Variant) As String
    
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    With objStream
        .Type = 1                   ' adTypeBinary (バイナリデーター)
        .Open
        .Write JsonBinary
        .Position = 0
        .Type = 2                   ' adTypeText (テキストデーター)
        .Charset = "UTF-8"
        DecodeUtf8Response = .ReadText
        .Close
    End With

End Function


' JSONテキストの中からtextキーの値（本文）を抽出する
' ここでは正規表現で抽出しているがJSONパーサーを使ってもよい
Private Function ExtractTextFromJson(ByVal jsonResponse As String) As String
    
    Dim matches As Object, strTemp As String
    With CreateObject("VBScript.RegExp")
        .Pattern = """text"":\s*?""(.*?)"""
        Set matches = .Execute(jsonResponse)
        If matches.Count > 0 Then
            strTemp = matches(0).SubMatches(0)
            strTemp = Replace(strTemp, "\""", """")
            strTemp = Replace(strTemp, "\\", "\")
            strTemp = Replace(strTemp, "\n", vbCrLf)

            ExtractTextFromJson = strTemp
        Else
            ExtractTextFromJson = "#ERROR:TextNotFound"
        End If
    End With

End Function
```
