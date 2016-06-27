VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private fso As Object
Private LOG As Object
Private filePathU As String
Public Sub Init()
   Set fso = CreateObject("ADODB.Stream")
End Sub
   
Public Sub OpenFile(filePath As String, mode As Integer)
    filePathU = filePath
    With fso
        .Charset = "UTF-8"
        .Type = 2 '文字列型
        .Open
   End With
End Sub

Public Sub WriteEx(str As Variant)
    fso.WriteText str
End Sub

Public Sub WriteLine(str As Variant)
    fso.WriteText str & vbCrLf
End Sub

Public Sub CloseEx()
   fso.SaveToFile (filePathU), 2
   
   ' バイナリモードにするためにPositionを一度0に戻す
   ' Readするためにはバイナリタイプでないといけない
   
   fso.Position = 0
   fso.Type = 1
   ' Positionを3にしてから読み込むことで最初の3バイトをスキップする
   ' つまりBOMをスキップします
   On Error GoTo Exit0:
   fso.Position = 3
   Dim bin: bin = fso.Read()
   fso.Close
  
   ' 読み込んだバイナリデータをバイナリデータとしてファイルに出力する
   ' ここは一般的な書き方なので説明を省略
   Dim stm As Object
   Set stm = CreateObject("ADODB.Stream")
   stm.Type = 1
   stm.Open
   stm.Write (bin)
   stm.SaveToFile filePathU, 2 ' force overwrite
   stm.Close
   
Exit0:
   
   Set LOG = Nothing
   Set fso = Nothing
End Sub

