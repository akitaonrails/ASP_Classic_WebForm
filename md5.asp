<%
' Derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm,
' as set out in the memo RFC1321.
'
' ASP VBScript code for generating an MD5 'digest' or 'signature' of a string. The
' MD5 algorithm is one of the industry standard methods for generating digital
' signatures. It is generically known as a digest, digital signature, one-way
' encryption, hash or checksum algorithm. A common use for MD5 is for password
' encryption as it is one-way in nature, that does not mean that your passwords
' are not free from a dictionary attack.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on this code provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk
'
' ------ 
' Code remanaged to fit in a VBScript Class for further modularity
' This code is offered "as is" and the requirements are the same as of the original
' author as written above
' E-mail:	akita@ieg.com.br	(do not contact to ask for support)

Class MD5

	Private BITS_TO_A_BYTE 
	Private BYTES_TO_A_WORD 
	Private BITS_TO_A_WORD 

	Private m_lOnBits(30)
	Private m_l2Power(30)
	
	Private Sub Class_Initialize
		BITS_TO_A_BYTE = 8
		BYTES_TO_A_WORD = 4
		BITS_TO_A_WORD = 32

		Dim i, j
		j = 0
		For i = 0 To 30
			j = ( j * 2 ) + 1
			m_lOnBits( i ) = CLng( j )
		Next
 
		j = 1
		m_l2Power(0) = CLng(1)
		For i = 1 To 30    
			j = j * 2
			m_l2Power( i ) = CLng( j )
		Next
	End Sub
	    
	Private Function LShift(lValue, iShiftBits)
	    If iShiftBits = 0 Then
	        LShift = lValue
	        Exit Function
	    ElseIf iShiftBits = 31 Then
	        If lValue And 1 Then
	            LShift = &H80000000
	        Else
	            LShift = 0
	        End If
	        Exit Function
	    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
	        Err.Raise 6
	    End If

	    If (lValue And m_l2Power(31 - iShiftBits)) Then
	        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
	    Else
	        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
	    End If
	End Function

	Private Function RShift(lValue, iShiftBits)
	    If iShiftBits = 0 Then
	        RShift = lValue
	        Exit Function
	    ElseIf iShiftBits = 31 Then
	        If lValue And &H80000000 Then
	            RShift = 1
	        Else
	            RShift = 0
	        End If
	        Exit Function
	    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
	        Err.Raise 6
	    End If
	    
	    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

	    If (lValue And &H80000000) Then
	        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
	    End If
	End Function

	Private Function RotateLeft(lValue, iShiftBits)
	    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
	End Function

	Private Function AddUnsigned(lX, lY)
	    Dim lX4
	    Dim lY4
	    Dim lX8
	    Dim lY8
	    Dim lResult
	 
	    lX8 = lX And &H80000000
	    lY8 = lY And &H80000000
	    lX4 = lX And &H40000000
	    lY4 = lY And &H40000000
	 
	    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
	 
	    If lX4 And lY4 Then
	        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
	    ElseIf lX4 Or lY4 Then
	        If lResult And &H40000000 Then
	            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
	        Else
	            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
	        End If
	    Else
	        lResult = lResult Xor lX8 Xor lY8
	    End If
	 
	    AddUnsigned = lResult
	End Function

	Private Function F(x, y, z)
	    F = (x And y) Or ((Not x) And z)
	End Function

	Private Function G(x, y, z)
	    G = (x And z) Or (y And (Not z))
	End Function

	Private Function H(x, y, z)
	    H = (x Xor y Xor z)
	End Function

	Private Function I(x, y, z)
	    I = (y Xor (x Or (Not z)))
	End Function

	Private Sub FF(a, b, c, d, x, s, ac)
	    a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac))
	    a = RotateLeft(a, s)
	    a = AddUnsigned(a, b)
	End Sub

	Private Sub GG(a, b, c, d, x, s, ac)
	    a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac))
	    a = RotateLeft(a, s)
	    a = AddUnsigned(a, b)
	End Sub

	Private Sub HH(a, b, c, d, x, s, ac)
	    a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac))
	    a = RotateLeft(a, s)
	    a = AddUnsigned(a, b)
	End Sub

	Private Sub II(a, b, c, d, x, s, ac)
	    a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac))
	    a = RotateLeft(a, s)
	    a = AddUnsigned(a, b)
	End Sub

	Private Function ConvertToWordArray(sMessage)
	    Dim lMessageLength
	    Dim lNumberOfWords
	    Dim lWordArray()
	    Dim lBytePosition
	    Dim lByteCount
	    Dim lWordCount
	    
	    Const MODULUS_BITS = 512
	    Const CONGRUENT_BITS = 448
	    
	    lMessageLength = Len(sMessage)
	    
	    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
	    ReDim lWordArray(lNumberOfWords - 1)
	    
	    lBytePosition = 0
	    lByteCount = 0
	    Do Until lByteCount >= lMessageLength
	        lWordCount = lByteCount \ BYTES_TO_A_WORD
	        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
	        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
	        lByteCount = lByteCount + 1
	    Loop

	    lWordCount = lByteCount \ BYTES_TO_A_WORD
	    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

	    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

	    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
	    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
	    
	    ConvertToWordArray = lWordArray
	End Function

	Private Function WordToHex(lValue)
	    Dim lByte
	    Dim lCount
	    
	    For lCount = 0 To 3
	        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
	        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
	    Next
	End Function

	Public Function MakeDigest(sMessage)
	    Dim x
	    Dim k
	    Dim AA
	    Dim BB
	    Dim CC
	    Dim DD
	    Dim a
	    Dim b
	    Dim c
	    Dim d
	    
	    Const S11 = 7
	    Const S12 = 12
	    Const S13 = 17
	    Const S14 = 22
	    Const S21 = 5
	    Const S22 = 9
	    Const S23 = 14
	    Const S24 = 20
	    Const S31 = 4
	    Const S32 = 11
	    Const S33 = 16
	    Const S34 = 23
	    Const S41 = 6
	    Const S42 = 10
	    Const S43 = 15
	    Const S44 = 21

	    x = ConvertToWordArray(sMessage)
	    
	    a = &H67452301
	    b = &HEFCDAB89
	    c = &H98BADCFE
	    d = &H10325476

	    For k = 0 To UBound(x) Step 16
	        AA = a
	        BB = b
	        CC = c
	        DD = d
	    
	        FF a, b, c, d, x(k + 0), S11, &HD76AA478
	        FF d, a, b, c, x(k + 1), S12, &HE8C7B756
	        FF c, d, a, b, x(k + 2), S13, &H242070DB
	        FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
	        FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
	        FF d, a, b, c, x(k + 5), S12, &H4787C62A
	        FF c, d, a, b, x(k + 6), S13, &HA8304613
	        FF b, c, d, a, x(k + 7), S14, &HFD469501
	        FF a, b, c, d, x(k + 8), S11, &H698098D8
	        FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
	        FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
	        FF b, c, d, a, x(k + 11), S14, &H895CD7BE
	        FF a, b, c, d, x(k + 12), S11, &H6B901122
	        FF d, a, b, c, x(k + 13), S12, &HFD987193
	        FF c, d, a, b, x(k + 14), S13, &HA679438E
	        FF b, c, d, a, x(k + 15), S14, &H49B40821
	    
	        GG a, b, c, d, x(k + 1), S21, &HF61E2562
	        GG d, a, b, c, x(k + 6), S22, &HC040B340
	        GG c, d, a, b, x(k + 11), S23, &H265E5A51
	        GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
	        GG a, b, c, d, x(k + 5), S21, &HD62F105D
	        GG d, a, b, c, x(k + 10), S22, &H2441453
	        GG c, d, a, b, x(k + 15), S23, &HD8A1E681
	        GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
	        GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
	        GG d, a, b, c, x(k + 14), S22, &HC33707D6
	        GG c, d, a, b, x(k + 3), S23, &HF4D50D87
	        GG b, c, d, a, x(k + 8), S24, &H455A14ED
	        GG a, b, c, d, x(k + 13), S21, &HA9E3E905
	        GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
	        GG c, d, a, b, x(k + 7), S23, &H676F02D9
	        GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
	            
	        HH a, b, c, d, x(k + 5), S31, &HFFFA3942
	        HH d, a, b, c, x(k + 8), S32, &H8771F681
	        HH c, d, a, b, x(k + 11), S33, &H6D9D6122
	        HH b, c, d, a, x(k + 14), S34, &HFDE5380C
	        HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
	        HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
	        HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
	        HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
	        HH a, b, c, d, x(k + 13), S31, &H289B7EC6
	        HH d, a, b, c, x(k + 0), S32, &HEAA127FA
	        HH c, d, a, b, x(k + 3), S33, &HD4EF3085
	        HH b, c, d, a, x(k + 6), S34, &H4881D05
	        HH a, b, c, d, x(k + 9), S31, &HD9D4D039
	        HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
	        HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
	        HH b, c, d, a, x(k + 2), S34, &HC4AC5665
	    
	        II a, b, c, d, x(k + 0), S41, &HF4292244
	        II d, a, b, c, x(k + 7), S42, &H432AFF97
	        II c, d, a, b, x(k + 14), S43, &HAB9423A7
	        II b, c, d, a, x(k + 5), S44, &HFC93A039
	        II a, b, c, d, x(k + 12), S41, &H655B59C3
	        II d, a, b, c, x(k + 3), S42, &H8F0CCC92
	        II c, d, a, b, x(k + 10), S43, &HFFEFF47D
	        II b, c, d, a, x(k + 1), S44, &H85845DD1
	        II a, b, c, d, x(k + 8), S41, &H6FA87E4F
	        II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
	        II c, d, a, b, x(k + 6), S43, &HA3014314
	        II b, c, d, a, x(k + 13), S44, &H4E0811A1
	        II a, b, c, d, x(k + 4), S41, &HF7537E82
	        II d, a, b, c, x(k + 11), S42, &HBD3AF235
	        II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
	        II b, c, d, a, x(k + 9), S44, &HEB86D391
	    
	        a = AddUnsigned(a, AA)
	        b = AddUnsigned(b, BB)
	        c = AddUnsigned(c, CC)
	        d = AddUnsigned(d, DD)
	    Next
	    
	    MakeDigest = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
	End Function
End Class
%>
