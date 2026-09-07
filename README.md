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
