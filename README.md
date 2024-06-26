# base64-encode-ls

# overview
入力ストリームのバイナリをBase64にエンコードし、出力ストリームに書き出す

# version
Notes/Domino 9.0.1 over

# use
```
	Dim session As New NotesSession
	
	Dim fileIn As NotesStream
	Set fileIn = session.Createstream()

  '---
  'エンコードしたいデータ(バイナリファイルでも可)
  '---
  Call fileIn.Open(Timer() & ".tmp", "UTF-8")
	Call fileIn.Writetext("a", EOL_NONE)
	Call fileIn.Writetext("あ", EOL_NONE)
	fileIn.Position = 0

  '---
  'Base64データの受け皿ストリーム
  '---
	Dim fileOut As NotesStream
	Set fileOut = session.Createstream()

  'Base64エンコード
	If Not EncodeBase64(fileIn, fileOut) = 0 Then
		Error Err, Error
	End If

  '不要なので削除
	Call fileIn.Truncate()
	Call fileIn.Close()

  '---
  '以下はデバッグ用
  '---
	Dim base64String As String
	
	fileOut.Position = 0
	base64String = fileOut.Readtext
	
	Print Now ":" base64String '=> YeOBgg==
```

# src
```vb:base64-encode.lss
%REM
	入力ストリームのバイナリをBase64にエンコードし、出力ストリームに書き出す
	@param {NotesStream} fileIn 入力ストリーム。 NotesStream.Position は 0 に明示的に割り当てする必要あり
	@param {NotesStream} fileOut 出力ストリーム。 NotesStream.Position は last に設定される
  @return {Integer} エラーの場合、そのエラー番号
%END REM
Function EncodeBase64(fileIn As NotesStream, fileOut As NotesStream) As Integer
	On Error GoTo ErrorHandle
	
	Do Until fileIn.IsEOS = True
		Dim bytes As Variant
		bytes = fileIn.Read(255)
		
		ForAll b In bytes
			Dim bit8 As String
			
			'byte -> bit
			bit8 = Bin(b)
			
			Dim bitLen As Byte
			Dim bit6 As String
			
			'bit length
			bitLen = Len(bit8)
			
			' If it is less than 8 bits, adjust it to fit
			If bitLen < 8 Then
				bit8 = String(8-bitLen, "0") & bit8
				bitLen = Len(bit8)
			End If
			
			Dim p As Byte
			Dim base64Pos As Byte
			Dim base64Char As String
			
			For p = 1 To bitLen
				bit6 = bit6 & Mid(bit8, p, 1)
				
				'Process until 6 bits are complete
				If Len(bit6) = 6 Then
					'bit -> byte
					base64Pos = CByte("&B" & bit6) + 1
					
					'byte -> char
					base64Char = Mid(BASE64_CHARS, base64Pos, 1)
					
					'add stream
					Call fileOut.Writetext(base64Char, EOL_NONE)
					
					'Clear after processing
					bit6 = ""
				End If
			Next
			
		End ForAll
	Loop
	
	'Not cleared
	If Not bit6 = "" Then
		'Less than 6 bits
		If Len(bit6) < 6 Then
			bit6 = bit6 & String(6-Len(bit6), "0")
		End If
		
		'bit -> byte
		base64Pos = CByte("&B" & bit6) + 1
		
		'byte -> char
		base64Char = Mid(BASE64_CHARS, base64Pos, 1)
		
		
		'add stream
		Call fileOut.Writetext(base64Char, EOL_NONE)
		
		'Clear after processing
		bit6 = ""
	End If
	
	'The output string is less than four characters long
	Do until (fileOut.Bytes /2) Mod 4 = 0
		'Fill with "="
		Call fileOut.Writetext("=", EOL_NONE)
	Loop 
	
	EncodeBase64 = 0
	
	Exit Function
ErrorHandle:
	EncodeBase64 = Err
	Exit Function
End Function
```
