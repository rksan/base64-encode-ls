# base64-encode-ls

# overview
* Encode : 入力ストリームのバイナリをBase64にエンコードし、出力ストリームに書き出す
* Decode : 入力ストリームのBase64テキストを、バイナリストリームとして出力する

# version
Notes/Domino 9.0.1 over

# use
## Encode
```vbs
Sub Initialize
	
	Dim session As New NotesSession
	
	Dim fileIn As NotesStream
	Set fileIn = session.Createstream()
	
	'---
	'エンコードしたいデータ(バイナリファイルでも可)
	'Call fileIn.Open("doclink.gif", "binary")
	'---
	Call fileIn.Open(Timer() & ".txt", "UTF-8")
	Call fileIn.Writetext("a", EOL_NONE)
	fileIn.Position = 0
	
	'Base64へエンコード
	Dim base64 As New CBase64
	If Not base64.Encode(fileIn) = 0 Then
		Error Err, Error
	End If
	
	'ファイルへエクスポート
	Dim fileOut As NotesStream
	Set fileOut = session.Createstream()
	
	Call fileOut.Open("***.txt", "ASCII")
	
	If Not base64.Export(Fileout) = 0 Then
		Error Err, Error
	End If
	
	Call fileOut.Close()
	
	'以下はデバッグ用
	Dim base64String As String
	base64String = base64.ToString()
	
	Print Now ":" base64String '->YQ==
	
End Sub
```

## Decode

```vba
Sub Initialize
	
	Dim ss As New NotesSession
	Dim stream As NotesStream
	
	'デコードしたいBase64テキスト
	Set stream = ss.Createstream()
	Call stream.open("***.txt", "ASCII")
	Call stream.Writetext("YQ==", EOL_NONE) 'YQ== -> a
	
	'デコード
	Dim base64 As New CBase64
	If Not base64.Decode(stream) = 0 Then
		Error Err, Error
	End If
	
	'エクスポート
	Set stream = ss.Createstream()
	Call stream.open("***.txt", "UTF-8")
	
	If not base64.Export(stream) = 0 Then
		Error Err, Error
	End If
	
	Call stream.close()
	
	'以下はデバッグ用
	Dim base64String As String
	base64String = base64.ToString()
	
	Print Now ":" & base64String ' empty
	
End Sub
```

# src
```vbs
Class CBase64
	
	Sub New()
		Me.zBufferLlength = 255
	End Sub
	
	Sub Delete()
		If Not Me.zFileOut Is Nothing Then
			Call Me.zFileOut.Truncate()
			Call Me.zFileOut.Close()
		End If
	End Sub
	
	'Returns the stored data as text.
	'@return {String}
	'	Encoded or decoded binary data.
	'	The result of encoding is stored as ASCII code,
	'	but the result of decoding is stored as binary data,
	'	so it may not meet your expectations.
	Function ToString() As String
		Me.zFileOut.Position = 0
		ToString = Me.zFileOut.Readtext
	End Function
	
	'Exporting retained data.
	'If you want to export text,
	'you need to specify the character encoding for the Stream.
	'@param {NotesStream} fileOut
	'	stream to write the encoded or decoded binary data to.
	'@return {Integer}
	'	If there is an error, the error number.
	Function Export(fileOut As NotesStream) As Integer
		On Error GoTo ErrorHandle
		
		If Me.zFileOut Is Nothing Then
			Exit Function
		End If
		
		Me.zFileOut.Position = 0
		
		Do Until Me.zFileOut.Iseos = True
			Call fileOut.Write(Me.zFileOut.Read(Me.zBufferLlength))
		Loop
		
		Exit Function
ErrorHandle:
		Export = Err
		Exit Function
	End Function
	
	' Text or Binary -> Base64
	'@param {NotesStream} fileIn
	'	Data stream to encode. text or binary.
	'@return {Integer}
	'	If there is an error, the error number.
	Function Encode(fileIn As NotesStream) As Integer
		On Error GoTo ErrorHandle
		
		Dim ss As New NotesSession
		Dim base64 As NotesStream
		
		Set base64 = ss.Createstream()
		
		Call base64.Open(Timer() & ".base64", "ASCII")
		
		Dim bit6 As String

		fileIn.Position = 0

		Do Until fileIn.Iseos = True
			
			Dim chunk As Variant
			chunk = fileIn.Read(Me.zBufferLlength)
			
			ForAll c In chunk
				Dim piese As Byte
				piese = c
				
				'byte -> bit8
				Dim bit8 As String
				If Not Me.zByteToBit(piese, bit8) = 0 Then
					Error Err, Error
				End If
				
				'bit8 -> bit6
				Dim p As Byte
				Dim char As String
				
				For p = 1 To 8
					bit6 = bit6 & Mid(bit8, p, 1)
					
					'Process until 6 bits are complete
					If Len(bit6) = 6 Then
						'bit6 -> char
						If Not Me.zBit6ToChar(bit6, char) = 0 Then
							Error Err, Error
						End If
						
						'add stream
						Call base64.Writetext(char, EOL_NONE)
						
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
			
			'bit -> char
			If Not Me.zBit6ToChar(bit6, char) = 0 Then
				Error Err, Error
			End If
			
			'add stream
			Call base64.Writetext(char, EOL_NONE)
			
			'Clear after processing
			bit6 = ""
		End If
		
		'The output string is less than four characters long
		Do Until (base64.Bytes) Mod 4 = 0
			'Fill with "="
			Call base64.Writetext("=", EOL_NONE)
		Loop
		
		base64.Position = 0
		
		Set Me.zFileOut = base64
		
		Exit Function
ErrorHandle:
		Encode = Err
		Exit Function
	End Function
	
	'Base64 -> Binary
	'@param {NotesStream} base64
	'	A base64 encoded text stream.
	'@return {Integer}
	'	If there is an error, the error number.
	Function Decode(base64 As NotesStream) As Integer
		On Error GoTo ErrorHandle
		
		Dim ss As New NotesSession
		Dim stream As NotesStream
		
		'All processed as binary data
		Set stream = ss.Createstream()
		Call stream.Open(Timer() & ".tmp", "binary")
		
		base64.Position = 0
		
		Do Until base64.Iseos = True
			Dim buffer As String
			buffer = base64.Readtext
			
			Do Until buffer = ""
				'buffer -> char
				Dim char As String
				char = Mid(buffer, 1, 1)
				buffer = Mid(buffer, 2)
				
				If char = "=" Then
					GoTo NextChar
				End If
				
				'char -> byte
				Dim p As Byte
				p = InStr(1, Me.zBase64Char, char, 0)
				
				If p = 0 Then
					GoTo NextChar
				End If
				
				'byte -> bit8
				Dim bit6 As String
				
				If Not Me.zByteToBit(p-1, bit6) = 0 Then
					Error Err, Error
				End If
				
				'bit8 -> bit6
				bit6 = Right(bit6, 6)
				
				Dim bit8 As String
				
				For p = 1 To Len(bit6)
					bit8 = bit8 & Mid(bit6, p, 1)
					
					If Len(bit8) = 8 Then
						Dim chunk(0) As Byte
						chunk(0) = CByte("&B" & bit8)
						
						Call stream.Write(chunk)
						
						bit8 = ""
					End If
				Next
				
NextChar:
			Loop
			
		Loop
		
		Set Me.zFileOut = stream
		
		Exit Function
ErrorHandle:
		Decode = Err
		Exit Function
	End Function
	
	' private ---
	
	Private zFileOut As NotesStream
	Private zBufferLlength As Long
	
	Private Property Get zBase64Char As String
		zBase64Char = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	End Property
	
	Private Function zByteToBit(piese As Byte, bit As String) As Integer
		On Error GoTo ErrorHandle
		
		'byte -> bit
		bit = Bin$(piese)
		
		' If it is less than 8 bits, adjust it to fit
		bit = Right("00000000" & bit, 8)
		
		Exit Function
ErrorHandle:
		zByteToBit = Err
		Exit Function
	End Function
	
	Private Function zBit6ToChar(bit6 As String, char As String) As Integer
		On Error GoTo ErrorHandle
		
		'bit -> byte
		Dim pos As Byte
		pos = CByte("&B" & bit6) + 1
		
		'byte -> char
		char = Mid(Me.zBase64Char, pos, 1)
		
		Exit Function
ErrorHandle:
		zBit6ToChar = Err
		Exit Function
	End Function
	
End Class
```
