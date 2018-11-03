<SCRIPT LANGUAGE=vbscript RUNAT=Server>
Class clsUpload
'========================================================='
'	This class will parse the binary contents of the 	  '
'	request, and populate the Form and Files collections. '
'========================================================='
	Private m_objFiles
	Private m_objForm
	
	Public Property Get Form()
		Set Form = m_objForm
	End Property
	
	Public Property Get Files()
		Set Files = m_objFiles
	End Property
	
	Private Sub Class_Initialize()
		Set m_objFiles = New VBollection
		Set m_objForm = New VBCollection
		ParseRequest
	End Sub
	
	Private Sub ParseRequest()
		Dim lngTotalBytes, lngPosBeg, lngPosEnd, lngPosBoundary, lngPosTmp, lngPosFileName
		Dim strBRequest, strBBoundary, strBContent
		Dim strName, strFileName, strContentType, strValue, strTemp
		Dim objFile

		'Grab the entire contents of the Request as a Byte string
		lngTotalBytes = Request.TotalBytes
		strBRequest = Request.BinaryRead(lngTotalBytes)
		
		'Find the first Boundary
		lngPosBeg = 1
		lngPosEnd = InStrB(lngPosBeg, strBRequest, UStr2Bstr(Chr(13)))
		If lngPosEnd > 0 Then
			strBBoundary = MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg)
			lngPosBoundary = InStrB(1, strBRequest, strBBoundary)
		End If

		If strBBoundary = "" Then
		'The form must have been submitted *without* ENCTYPE="multipart/form-data"
		'But since we already called Request.BinaryRead, we can no longer access
		'the Request.Form collection, so we need to parse the request and populate
		'our own form collection.
			lngPosBeg = 1
			lngPosEnd = InStrB(lngPosBeg, strBRequest, UStr2BStr("&"))
			Do While lngPosBeg < LenB(strBRequest)
				'Parse the element and add it to the collection
				strTemp = BStr2UStr(MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg))
				lngPosTmp = InStr(1, strTemp, "=")
				strName = URLDecode(Left(strTemp, lngPosTmp - 1))
				strValue = URLDecode(Right(strTemp, Len(strTemp) - lngPosTmp))
' --- 10/01/2001 - merge list results (radio lists) in a single comma separated value
				If m_objForm.Exists( strName ) Then
					m_objForm.setItem strName, m_objForm.getValue( strName ) & "," & strValue
				Else
					m_objForm.Add strName, strValue
				End If
				
				'Find the next element
				lngPosBeg = lngPosEnd + 1
				lngPosEnd = InStrB(lngPosBeg, strBRequest, UStr2BStr("&"))
				If lngPosEnd = 0 Then lngPosEnd = LenB(strBRequest) + 1
			Loop
		Else
		'The form was submitted with ENCTYPE="multipart/form-data"
		'Loop through all the boundaries, and parse them into either the
		'Form or Files collections.
			Do Until (lngPosBoundary = InStrB(strBRequest, strBBoundary & UStr2Bstr("--")))
				'Get the element name
				lngPosTmp = InStrB(lngPosBoundary, strBRequest, UStr2BStr("Content-Disposition"))
				lngPosTmp = InStrB(lngPosTmp, strBRequest, UStr2BStr("name="))
				lngPosBeg = lngPosTmp + 6
				lngPosEnd = InStrB(lngPosBeg, strBRequest, UStr2BStr(Chr(34)))
				strName = BStr2UStr(MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg))
				'Look for an element named 'filename'
				lngPosFileName = InStrB(lngPosBoundary, strBRequest, UStr2BStr("filename="))
				'If found, we have a file, otherwise it is a normal form element
				If lngPosFileName <> 0 And lngPosFileName < InStrB(lngPosEnd, strBRequest, strBBoundary) Then 'It is a file
					'Get the FileName
					lngPosBeg = lngPosFileName + 10
					lngPosEnd = InStrB(lngPosBeg, strBRequest, UStr2BStr(chr(34)))
					strFileName = BStr2UStr(MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg))
					'Get the ContentType
					lngPosTmp = InStrB(lngPosEnd, strBRequest, UStr2BStr("Content-Type:"))
					lngPosBeg = lngPosTmp + 14
					lngPosEnd = InstrB(lngPosBeg, strBRequest, UStr2BStr(chr(13)))
					strContentType = BStr2UStr(MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg))
					'Get the Content
					lngPosBeg = lngPosEnd + 4
					lngPosEnd = InStrB(lngPosBeg, strBRequest, strBBoundary) - 2
					strBContent = MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg)
					If strFileName <> "" And strBContent <> "" Then
						'Create the File object, and add it to the Files collection
						Set objFile = New clsFile
						objFile.Name = strName
						objFile.FileName = Right(strFileName, Len(strFileName) - InStrRev(strFileName, "\"))
						objFile.ContentType = strContentType
						objFile.Blob = strBContent
						m_objFiles.Add strName, objFile
					End If
				Else 'It is a form element
					'Get the value of the form element
					lngPosTmp = InStrB(lngPosTmp, strBRequest, UStr2BStr(chr(13)))
					lngPosBeg = lngPosTmp + 4
					lngPosEnd = InStrB(lngPosBeg, strBRequest, strBBoundary) - 2
					strValue = BStr2UStr(MidB(strBRequest, lngPosBeg, lngPosEnd - lngPosBeg))
					'Add the element to the collection
' --- 10/01/2001 - merge list results (radio lists) in a single comma separated value
				If m_objForm.Exists( strName ) Then
					m_objForm.setItem strName, m_objForm.getItem( strName ) & "," & strValue
				Else
					m_objForm.Add strName, strValue
				End If

				End If
				'Move to Next Element
				lngPosBoundary = InStrB(lngPosBoundary + LenB(strBBoundary), strBRequest, strBBoundary)
			Loop
		End If
	End Sub
	
	Private Function BStr2UStr(BStr)
	'Byte string to Unicode string conversion
		Dim lngLoop
		BStr2UStr = ""
		For lngLoop = 1 to LenB(BStr)
			BStr2UStr = BStr2UStr & Chr(AscB(MidB(BStr,lngLoop,1))) 
		Next
	End Function
	
	Private Function UStr2Bstr(UStr)
	'Unicode string to Byte string conversion
		Dim lngLoop
		Dim strChar
		UStr2Bstr = ""
		For lngLoop = 1 to Len(UStr)
			strChar = Mid(UStr, lngLoop, 1)
			UStr2Bstr = UStr2Bstr & ChrB(AscB(strChar))
		Next
	End Function

	' the old deprecated function. replace is below
	Private Function URLDecode_OLD(Expression)
	'Why doesn't ASP provide this functionality for us?
		Dim strSource, strTemp, strResult
		Dim lngPos
		strSource = Replace(Expression, "+", " ")
		For lngPos = 1 To Len(strSource)
			strTemp = Mid(strSource, lngPos, 1)
			If strTemp = "%" Then
				If lngPos + 2 < Len(strSource) Then
					strResult = strResult & Chr(CInt("&H" & Mid(strSource, lngPos + 1, 2)))
					lngPos = lngPos + 2
				End If
			Else
				strResult = strResult & strTemp
			End If
		Next
		URLDecode = strResult
	End Function	

	' new URLDecode function. Made public for general purpose uses outside this object
	' - there´s no problem because it doesn´t affect any internal state
	Public Function URLDecode(SSdecode)
		dim SStemp(1,1)

		SSin = SSdecode
		SSout = ""
		SSin = replace(SSin, "+", " " ,1 ,-1 ,1)
		SSpos = instr(SSin, "%")
		do while SSpos
			SSlen = len(SSin)
			if SSpos > 1 then 
				SSout = SSout & left(SSin, SSpos - 1)
			End If
			SStemp(0,0) = mid(SSin, SSpos + 1, 1)
			SStemp(1,0) = mid(SSin, SSpos + 2, 1)
			for SSi = 0 to 1
				if asc(SStemp(SSi,0)) > 47 and asc(SStemp(SSi,0)) < 58 then
					SStemp(SSi,1) = asc(SStemp(SSi,0)) - 48
				else
					SStemp(SSi,1) = asc(SStemp(SSi,0)) - 55
				end if
			next
			SSout = SSout & chr((SStemp(0,1) * 16) + SStemp(1,1))
			SSin = right(SSin, (SSlen - (SSpos + 2)))
			SSpos = instr(SSin, "%")
		loop
		URLDecode = SSout & SSin
	end function

End Class


'========================================================='
'	This class is a pseudo-collection. It is not a real   '
'	collection, because there is no way that I am aware   '
'   of to implement an enumerator to support the		  '
'	For..Each syntax using VBScript classes.			  '
'========================================================='
Class clsCollection
	Private m_objDicItems
	
	Private Sub Class_Initialize()
		Set m_objDicItems = Server.CreateObject("Scripting.Dictionary")
		m_objDicItems.CompareMode = vbTextCompare
	End Sub
	
	Private Sub Class_Terminate
		Set m_objDicItems = Nothing
	End Sub
	
	Public Property Get Count()
		Count = m_objDicItems.Count
	End Property
	
	Public Function Exists( sKey )
		Exists = m_objDicItems.Exists( skey )
	End Function
	
	Public Function GetValue( sKey )
		GetValue = m_objDicItems.Item( sKey )
	End Function
	
	Public Sub SetValue( sKey, sValue )
		m_objDicItems.Item( sKey ) = sValue
	End Sub
	
	Public Default Function Item(Index)
		Dim arrItems
		If IsNumeric(Index) Then
			arrItems = m_objDicItems.Items
			If IsObject(arrItems(Index)) Then
				Set Item = arrItems(Index)
			Else
				Item = arrItems(Index)
			End If
		Else
			If m_objDicItems.Exists(Index) Then
				If IsObject(m_objDicItems.Item(Index)) Then
					Set Item = m_objDicItems.Item(Index)
				Else
					Item = m_objDicItems.Item(Index)
				End If
			End If
		End If
	End Function
	
	Public Function Key(Index)
		Dim arrKeys
		If IsNumeric(Index) Then
			arrKeys = m_objDicItems.Keys
			Key = arrKeys(Index)
		End If
	End Function
	
	Public Sub Add(Name, Value)
		If m_objDicItems.Exists(Name) Then
			m_objDicItems.Item(Name) = Value
		Else
			m_objDicItems.Add Name, Value
		End If
	End Sub
End Class


'========================================================='
'	This class is used as a container for a file sent via '
'	an http multipart/form-data post.					  '
'========================================================='
Class clsFile
	Private m_strName
	Private m_strContentType
	Private m_strFileName
	Private m_Blob
	
	Public Property Get Name() : Name = m_strName : End Property
	Public Property Let Name(vIn) : m_strName = vIn : End Property
	Public Property Get ContentType() : ContentType = m_strContentType : End Property
	Public Property Let ContentType(vIn) : m_strContentType = vIn : End Property
	Public Property Get FileName() : FileName = m_strFileName : End Property
	Public Property Let FileName(vIn) : m_strFileName = vIn : End Property
	Public Property Get Blob() : Blob = m_Blob : End Property
	Public Property Let Blob(vIn) : m_Blob = vIn : End Property

	Public Sub Save(Path)
		Dim objFSO, objFSOFile
		Dim lngLoop
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set objFSOFile = objFSO.CreateTextFile(objFSO.BuildPath(Path, m_strFileName))
		For lngLoop = 1 to LenB(m_Blob)
			objFSOFile.Write Chr(AscB(MidB(m_Blob, lngLoop, 1)))
		Next
		objFSOFile.Close
		Set objFSOFile = Nothing
		Set objFSO = Nothing
	End Sub
End Class
</SCRIPT>