<%
Class Base64

	Private B64_RAW_CHAR_DICT
	Private B64_PAD_CHAR

	Private Sub Class_Initialize
		B64_RAW_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
		B64_PAD_CHAR = "="
	End Sub

	'//-------------------------------------------------------------------- --
	'|| Procedure:      Encode
	'||                 (
	'||                 TextStream As String
	'||                 ) As String
	'||
	'|| Description:    Base64 encodes input
	'||
	'|| Notes:          If an error occurs, the input string will be
	'||                 returned to the calling procedure.  There is no
	'||                 other error handling.
	'||
	'||-------------------------------------------------------------------- --
	'|| Date        Eng     Ver     Description
	'|| 20000823    JKF     1.0     Initial version
	'||
	'\\-------------------------------------------------------------------- --
	'Public Function Encode( TextStream As String ) As String
	Public Function Encode( TextStream )

	    Dim intLoopA ' As Integer
	    Dim intLoopB ' As Integer
	    Dim bytArray() ' As Byte
	    Dim bitArray() ' As Boolean
	    Dim intPadFactor ' As Integer
	    Dim strBuffer ' As String

	    If Len(TextStream) = 0 Then Exit Function

	    '// Poor man's error handling.  If anything bad happens, the procedure
	    '|| will return the input string to the caller, signaling an error.
	    '|| This is a reliable method becuase successful encoding will never
	    '|| result in Encode = TextStream
	    Encode = TextStream

	    '// Size the byte array to recieve the incoming text stream
	    ReDim bytArray(Len(TextStream) - 1)

	    '// Put incoming text stream into byte array
	    For intLoopA = 0 To UBound(bytArray)
	        bytArray(intLoopA) = CByte(Asc(Mid(TextStream, intLoopA + 1)))
	    Next

	    '// Now build our bit array, one byte at a time
	    ReDim bitArray(((UBound(bytArray) + 1) * 8) - 1)
	    For intLoopA = 0 To UBound(bytArray)
	        For intLoopB = 7 To 0 Step -1
	            '// Do a most-to-least bitwise assignment into the bit array
	            '|| from the current byte.
	            bitArray((intLoopA * 8) + (7 - intLoopB)) = _
	                CBool(((bytArray(intLoopA) And (2 ^ intLoopB)) > 0))
	        Next
	    Next

	    '// Check to make sure that the bitArray is integral number of 6-bit
	    '|| parts.
	    intPadFactor = 0
	    Select Case ((UBound(bitArray) + 1) Mod 6)
	        '// N.B. There's no case else here.  The value may be Mod 6 = 0,
	        '|| in which case  the final quantum of encoding input is an
	        '|| integral multiple of 24 bits.  In this case, the final unit
	        '|| of encoded output will be an integral multiple of 4 characters
	        '\\ with no "=" padding
	        Case 2
	            '// The final quantum of encoding input is exactly 8 bits
	            '|| In this case, the final unit of encoded output will be
	            '|| two characters followed by two "=" padding characters.
	            '|| Hence the bitArray must be padded with 4 zeros yielding:
	            '||     bb0000
	            ReDim Preserve bitArray(UBound(bitArray) + 4)
	            intPadFactor = 2
	        Case 4
	            '// The final quantum of encoding input is exactly 16 bits
	            '|| In this case, the final unit of encoded output will be
	            '|| three characters followed by one "=" padding character.
	            '|| Hence the bitArray must be padded with 2 zeros yielding:
	            '||     bbbb00
	            ReDim Preserve bitArray(UBound(bitArray) + 2)
	            intPadFactor = 1
	    End Select

	    '// Now we create a new output byte array composed of sextets pulled
	    '|| from our bit array.
	    ReDim bytArray((UBound(bitArray) / 6) - 1)
	    For intLoopA = 0 To UBound(bytArray)
	        '// Assign the bit sextets into the six lowest bits of each new byte
	        '|| resulting in 00bbbbbb, so that the range of possible values is now
	        '|| 0 - 63 inclusive (or 64 discreet values.)
	        For intLoopB = 0 To 5
	            If bitArray((intLoopA * 6) + intLoopB) Then
	                bytArray(intLoopA) = (bytArray(intLoopA) Or 2 ^ (5 - intLoopB))
	            End If
	        Next
	    Next

	    '// Map the new byte values to the base64 character set
	    For intLoopA = 0 To UBound(bytArray)
	        strBuffer = strBuffer & Mid(B64_RAW_CHAR_DICT, CLng(bytArray(intLoopA)) + 1, 1)
	    Next

	    '// Pad if neccessary
	    strBuffer = strBuffer & String(intPadFactor, B64_PAD_CHAR)

	    Encode = strBuffer
	End Function

	'//-------------------------------------------------------------------- --
	'|| Procedure:      Decode
	'||                 (
	'||                 TextStream As String
	'||                 ) As String
	'||
	'|| Description:    Decodes Base64 input
	'||
	'|| Notes:          If an error occurs, the input string will be
	'||                 returned to the calling procedure.  There is no
	'||                 other error handling.
	'||
	'||-------------------------------------------------------------------- --
	'|| Date        Eng     Ver     Description
	'|| 20000823    JKF     1.0     Initial version
	'||
	'\\-------------------------------------------------------------------- --
	'Public Function Decode(TextStream As String) As String
	Public Function Decode( TextStream )
	    Dim intLoopA ' As Integer
	    Dim intLoopB ' As Integer
	    Dim intPadFactor ' As Integer
	    Dim bytArray() ' As Byte
	    Dim bitArray() ' As Boolean
	    Dim strBuffer ' As String

	    If (Len(TextStream) & "") = 0 Then Exit Function

	    '// Poor man's error handling.  If anything bad happens, the procedure
	    '|| will return the input string to the caller, signaling an error.
	    '|| This is a reliable method becuase successful decoding will never
	    '|| result in Decode = TextStream
	    Decode = TextStream

	    '// Validate input as Base64 encoded text stream
	    For intLoopA = 1 To Len(TextStream)
	        '// Does TextStream conatain any invalid (i.e. non-Base64) characters,
	        '|| either encodings or pad ("=" equals sign)?
	        If (InStr(1, B64_RAW_CHAR_DICT, Mid(TextStream, intLoopA, 1), vbBinaryCompare) = 0) And _
	                (Mid(TextStream, intLoopA, 1) <B64_PAD_CHAR) Then
	            Decode = TextStream
	            Exit Function
	        End If
	    Next

	    '// Determine the 'pad factor'.  Will be 0,1 or 2 equals ("=") signs tacked onto
	    '|| the end of the Base64 encoded text stream.  So we have one of the following
	    '|| three possibilities as the last two characters at the end of the stream:
	    '||     "XX" = 0 pad factor (where the Xs are normal, valid Base64 characters)
	    '||     "X=" = 1 pad factor (the X is a normal, valid Base64 character)
	    '||     "==" = 2 pad factor
	    '|| The padding does not decode, but simply acts as a flag to indicate that the
	    '|| final quantum of the Base64 binary stream was not an intergal multiple of 24
	    '|| bits (pad factor 0) but instead was either exactly 8 or 16 bits (pad factor
	    '|| 2 or 1 respectively) to which we appended the correct number of zeros to complete
	    '|| the 24 bit quantum.  The pad factor just lets us know how many zeros to strip
	    '|| off the end of the resolved binary stream (because they're padding!)
	    '|| I'll leave it up to you to explore the technique I'm using here to do the work
	    '|| in a single line of code (who says VB can't be elegant?!)
	    intPadFactor = ((CByte(InStr(1, Right(TextStream, 2), B64_PAD_CHAR, vbBinaryCompare)) And (2 ^ 0)) * 2) + _
	        ((CByte(InStr(1, Right(TextStream, 2), B64_PAD_CHAR, vbBinaryCompare)) And (2 ^ 1)) / 2)

	    '// Strip any pad characters
	    TextStream = Mid(TextStream, 1, Len(TextStream) - intPadFactor)

	    '// "Unmap" the TextStream  from the Base64 encodings into a byte array
	    ReDim bytArray(Len(TextStream) - 1)
	    For intLoopA = 0 To UBound(bytArray)
	        bytArray(intLoopA) = CByte(InStr(1, B64_RAW_CHAR_DICT, Mid(TextStream, intLoopA + 1, 1), vbBinaryCompare) - 1)
	    Next

	    '// Now build our bit array, one "six-bit byte" at a time
	    ReDim bitArray(((UBound(bytArray) + 1) * 6) - 1)
	    For intLoopA = 0 To UBound(bytArray)
	        For intLoopB = 5 To 0 Step -1
	            '// Do a most-to-least bitwise assignment into the six
	            '|| right-hand bits from the current byte.
	            bitArray((intLoopA * 6) + (5 - intLoopB)) = _
	                CBool(((bytArray(intLoopA) And (2 ^ intLoopB)) > 0))
	        Next
	    Next

	    '// Remove zero padding
	    ReDim Preserve bitArray(UBound(bitArray) - (intPadFactor * 2))

	    '// Load the bit array into the byte array
	    ReDim bytArray((UBound(bitArray) / 8) - 1)
	    For intLoopA = 0 To UBound(bytArray)
	        '// Set the appropriate bits in each byte
	        For intLoopB = 0 To 7
	            If bitArray(intLoopA * 8 + intLoopB) Then
	                bytArray(intLoopA) = (bytArray(intLoopA) Or 2 ^ (7 - intLoopB))
	            End If
	        Next
	    Next

	    '// Load the bytes into the output string
	    For intLoopA = 0 To UBound(bytArray)
	        strBuffer = strBuffer & Chr(CLng(bytArray(intLoopA)))
	    Next

	    Decode = strBuffer

	End Function
	
End Class
%>