<%
'
'	Encapsulates all the functionallity of a Page/Form
'	- wraps the Request object and Form elements handling
'
Class Page
	Private objQueue
	Private objEncoders
	Private objBase64
	Private objForm
	Private objValidator
	Private strPost
	Private strFormName
	Private strSavePath
	Private strJSMD5Path
	Private boolEncode
	Private boolHashPassword
	Private childDelimiter
	Private childNodeDelimiter
	Private childValueDelimiter
	
	Private evtLoad
	Private evtInit
	Private evtTerminate
	Private evtFileUpload	' must expects a reference to the files collection
	Private boolEvtLoad
	Private boolEvtInit
	Private boolEvtTerminate
	Private boolEvtFileUpload	' must expects a reference to the files collection

	' constructor - gets the submitted form elements and initialize vital objects
	Private Sub Class_Initialize
		childDelimiter = "<br />"
		childNodeDelimiter = "&"
		childValueDelimiter = "="

		strSavePath = "undefined"
		strJSMD5Path = "md5.js"
		boolHashPassword = True

		boolEvtLoad = False
		boolEvtInit = False
		boolEvtTerminate = False
		boolEvtFileUpload = False

		Set objQueue = Server.CreateObject( "Scripting.Dictionary" )
		Set objEncoders = Server.CreateObject( "Encoders" )
		Set objBase64 = objEncoders.CreateInstance( "base64" )
		strPost = ""
		If Request.TotalBytes > 0 Then
			Set objForm = New clsUpload
			' correct the last character broken by the clsUpload class
			strPost = objForm.Form.Item( "_VIEWSTATE" )

			If Not IsEmpty( strPost ) Then
				'strPost = Left( strPost, Len( strPost ) - 2 ) & "="
				strPost = objBase64.Decode( strPost )
			End If
		End If
	End Sub
	
	' destructor - clean up
	Private Sub Class_Terminate
		Set objBase64 = Nothing
		Set objQueue = Nothing
		Set objForm = Nothing
	End Sub

	' set the OnLoad event handler
	Public Property Let OnLoad( ByRef subRef ) : Set evtLoad = subRef : boolEvtLoad = True : End Property

	' set the OnInit event handler
	Public Property Let OnInit( ByRef subRef ) : Set evtInit = subRef : boolEvtInit = True : End Property

	' set the OnTerminate event handler
	Public Property Let OnTerminate( ByRef subRef ) : Set evtTerminate = subRef : boolEvtTerminate = True : End Property

	' set the OnTerminate event handler
	Public Property Let OnFileUpload( ByRef subRef ) : Set evtFileUpload = subRef : boolEvtFileUpload = True : End Property
	
	' get/set the form name
	Public Property Get FormName : FormName = strFormName : End Property
	Public Property Let FormName( str ) : strFormName = str : End Property

	' get the form collection
	Public Property Get Form : Set Form = objForm : End Property

	' get the postedchallenge collection
	Public Property Get ChallengeSeed
		ChallengeSeed = ""
		If IsPostBack Then
			ChallengeSeed = objForm.Form.Item( "_CHALLENGE" )
		End If
	End Property

	' get/set the path of the client side md5 script
	Public Property Get MD5Path : MD5Path = strJSMD5Path : End Property
	Public Property Let MD5Path( str ) : strJSMD5Path = str : End Property

	' set validator container
	Public Property Get Validator : Set Validator = objValidator : End Property
	Public Property Let Validator( ByRef obj ) 
		Set objValidator = obj 
		objValidator.Parent = Me
	End Property
	
	' check all the registered validators
	Public Property Get IsValid
		IsValid = False
		If Not IsNull( objValidator ) Then
			IsValid = objValidator.IsValid
		End If
	End Property

	' choose to transport hashed version of the password instead of plain-text
	Public Property Get HidePassword : HidePassword = boolHashPassword : End Property
	Public Property Let HidePassword( bool )
		If bool Then
			boolHashPassword = True
		Else
			boolHashPassword = False
		End If
	End Property
	
	' choose to make a binary encoded form for uploads
	Public Property Get AllowUpload : AllowUpload = boolEncode : End Property
	Public Property Let AllowUpload( bool )
		If bool and LCase( bool ) <> "false" Then
			boolEncode = True
		Else
			boolEncode = False
		End If
	End Property

	' get/set the uploaded files saving location
	Public Property Get UploadPath : UploadPath = strSavePath : End Property
	Public Property Let UploadPath( str ) : strSavePath = str : End Property

	' check if this is a form post back (result of a previous submit)
	Public Property Get IsPostBack
		If strPost <> "" Then
			IsPostBack = True
		Else
			IsPostBack = False
		End If
	End Property

	' cancel deserialization and other post back event operation
	Public Sub CancelPostBack
		strPost = ""
	End Sub

	' function called by the child objects
	Public Function ChildSerializeNode( sName, sValue )
		ChildSerializeNode = childNodeDelimiter & sName & childValueDelimiter & Server.URLEncode( sValue )
	End Function

	' persists all the elements currently registered objects
	Private Function Serialize
		Dim strResult, objTmp
		For Each sElement in objQueue
			Set objTmp = objQueue.Item( sElement )
			strResult = strResult & objTmp.Serialize() & childDelimiter
		Next
		
		' serialize the form큦 own properties
		strResult = strResult & _
		ChildSerializeNode( "name", strFormName ) & _
		ChildSerializeNode( "path", strSavePath ) & _
		ChildSerializeNode( "hidepwd", boolHashPassword ) & _
		ChildSerializeNode( "encode", boolEncode )
		
		Serialize = objBase64.Encode( strResult )
	End Function
	
	' wraps the object큦 properties
	Public Sub SetProperty( strKey, strArg )
		Select Case strKey
		Case "name"
			FormName = strArg
		Case "path"
			UploadPath = strArg
		Case "hidepwd"
			HidePassword = strArg
		Case "encode"
			AllowUpload = strArg
		End Select
	End Sub

	' adds a new form elements
	Public Sub Add( ByRef oElement, strName )
		If Not objQueue.Exists( strName ) and strName <> "" Then
			objQueue.Add strName, oElement
			oElement.Name = strName
			oElement.Parent = Me
		End If
	End Sub
	
	' returns an element based on it큦 name
	Public Function Item( strName )
		If objQueue.Exists( strName ) Then
			Set Item = objQueue( strName )
		End If
	End Function

	' loads the elements with the previous stored persistent data
	' and trigger the current event
	Public Sub Load
		' call the event handler
		If boolEvtInit Then
			on error resume next
			Call evtInit
			on error goto 0
		End If

		Dim oList, oElement, intCount, intSub
		oList = Split( strPost, childDelimiter )
		If Not IsArray( oList ) Then
			Exit Sub
		End If

		' iterate thru all the returned objects and deserialize them
		For intCount = 0 To UBound( oList )
			oList( intCount ) = Split( oList( intCount ), "&" )
			' check if the element had persisted data posted back
			If IsArray( oList( intCount ) ) Then
				If objQueue.Exists( oList( intCount )( 0 ) ) Then
					' if yes then iterates thru all it큦 properties
					For intSub = 1 To UBound( oList( intCount ) )
						oElement = Split( oList( intCount )( intSub ), "=" )
						If IsArray( oElement ) Then
							If UBound( oElement ) = 1 Then
								Call objQueue.Item( oList( intCount )( 0 ) ).SetProperty( oElement( 0 ), objForm.URLDecode( oElement( 1 ) ) )
							End If
						End If
					Next
				ElseIf oList( intCount )( 0 ) = strFormName Then
					' it큦 the forms own properties
					For intSub = 1 To UBound( oList( intCount ) )
						oElement = Split( oList( intCount )( intSub ), "=" )
						If IsArray( oElement ) Then
							If UBound( oElement ) = 1 Then
								Call Me.SetProperty( oElement( 0 ), objForm.URLDecode( oElement( 1 ) ) )
							End If
						End If
					Next
				End If
			End If
		Next
		Set oList = Nothing

		If IsPostBack Then
			' update the correct values submitted by the users
			For Each sElement in objQueue
				If objForm.Form.Exists( sElement ) Then
					objQueue.Item( sElement ).Value = objForm.Form.GetValue( sElement )
				End If
			Next
			
			' check for triggered events
			If objForm.Form.Item( "_EVENTTARGET" ) <> "" Then
				For Each sElement in objQueue
					If sElement = objForm.Form.Item( "_EVENTTARGET" ) Then
						Call objQueue.Item( sElement ).EventHandler( objForm.Form.Item( "_EVENTTARGET" ), objForm.Form.Item( "_EVENTARGS" ) )
						Exit For
					End If
				Next
			End If		

			' call the event handler
			If boolEvtFileUpload Then
				on error resume next
				Call evtFileUpload( objForm.Files )
				on error goto 0
			ElseIf strSavePath <> "undefined" Then
				If objForm.Files.Count > 1 Then
					For intCount = 0 To objForm.Files.Count - 1
						Call objForm.Files.Item( intCount ).Save( strSavePath )
					Next
				ElseIf objForm.Files.Count = 1 Then
					Call objForm.Files.Item( 0 ).Save( strSavePath )
				End If
			End If

		End If
		
		' call the event handler
		If boolEvtLoad Then
			on error resume next
			Call evtLoad
			on error goto 0
		End If

	End Sub
	
	' prints the <form> element and the main javascript behaviors
	Public Sub RenderBegin
		Dim strHTML
		
		If boolHashPassword Then
			strHTML = strHTML & vbCRLF & _		
			"<script language=""javascript"" src=""" & strJSMD5Path & """></script>" & vbCRLF
		End If
		
		strHTML = strHTML & vbCRLF & _		
		"<script language=""javascript""><" & "!--" & vbCRLF & _
		"var __IsPosting = false; // flag to avoid double-posting commands" & vbCRLF & _
		"var __arrValidators = new Array(); " & vbCRLF & _
		"function _doSubmit() { " & vbCRLF & _
		"	if ( __IsPosting ) { return false } " & vbCRLF & _
		"	__IsPosting = true; " & vbCRLF & _
		"	var boolReturn = true; " & vbCRLF & _
		"	if ( __arrValidators.length > 0 ) { " & vbCRLF & _
		"		for ( var i = 0; i < __arrValidators.length; i++ ) { " & vbCRLF & _
		"			boolReturn = boolReturn && eval( __arrValidators[ i ] ); " & vbCRLF & _
		"		} " & vbCRLF & _
		"	}" & vbCRLF & _
		"	if ( boolReturn ) { " & vbCRLF & _
		"		_hashPassword(); " & vbCRLF & _
		"		%formname.submit(); " & vbCRLF & _
		"	} else { " & vbCRLF & _
		"		__IsPosting = false;" & vbCRLF & _
		"	} " & vbCRLF & _
		"	return false; " & vbCRLF & _
		"} " & vbCRLF
		
		If boolHashPassword Then
			strHTML = strHTML & vbCRLF & _
			"function _hashPassword() { " & vbCRLF & _
			"	if ( ! MD5 ) { " & vbCRLF & _
			"		return;" & vbCRLF & _
			"	} " & vbCRLF & _
			"	var strUsername = ''; " & vbCRLF & _
			"	if ( %formname.username ) { " & vbCRLF & _
			"		strUsername = %formname.username.value + ':'; " & vbCRLF & _
			"	} " & vbCRLF & _
			"	for( var i = 0; i < %formname.elements.length; i++ ) { " & vbCRLF & _
			"		if ( %formname.elements[ i ].type == 'password' ) { " & vbCRLF & _
			"			strHash = MD5( strUsername + MD5( %formname.elements[ i ].value ) + ':' + %formname._CHALLENGE.value ); " & vbCRLF & _
			"			%formname.elements[ i ].value = strHash; " & vbCRLF & _
			"		} " & vbCRLF & _
			"	} " & vbCRLF & _
			"} " & vbCRLF
		Else
			strHTML = strHTML & vbCRLF & _
			"function _hashPassword() {}" & vbCRLF
		End If
		
		strHTML = strHTML & vbCRLF & _		
		"function _doPostBack( element, args ) { " & vbCRLF & _
		"	%formname._EVENTTARGET.value = element; " & vbCRLF & _
		"	%formname._EVENTARGS.value = args; " & vbCRLF & _
		"	_doSubmit(); " & vbCRLF & _
		"} " & vbCRLF & _
		"//--" & "></script>" & vbCRLF & _
		"<form id=""%formname"" name=""%formname"" method=""POST"" action=""%url"" %encode onsubmit=""return _doSubmit()"">" & vbCRLF & _
		"<input type=""hidden"" id=""_VIEWSTATE"" name=""_VIEWSTATE"" value=""%viewstate"">" & vbCRLF & _
		"<input type=""hidden"" id=""_EVENTTARGET"" name=""_EVENTTARGET"">" & vbCRLF & _
		"<input type=""hidden"" id=""_EVENTARGS"" name=""_EVENTARGS"">" & vbCRLF & _
		"<input type=""hidden"" id=""_CHALLENGE"" name=""_CHALLENGE"" value=""%challenge"">" & vbCRLF 

		' add upload capability		
		If boolEncode Then
			strHTML = Replace( strHTML, "%encode", "ENCTYPE=""multipart/form-data""" )
		Else
			strHTML = Replace( strHTML, "%encode", "" )
		End If
		
		Randomize
		strHTML = Replace( strHTML, "%challenge", Round( Rnd * 30000 ) + 1 )
		strHTML = Replace( strHTML, "%formname", strFormName, 1, -1 )
		strHTML = Replace( strHTML, "%url", Request.ServerVariables( "PATH_INFO" ) )
		strHTML = Replace( strHTML, "%viewstate", Serialize() )
		
		Response.Write strHTML
	End Sub
	
	' print the end of the <form> element
	Public Sub RenderEnd
		Response.Write vbCRLF & "</form>" & vbCRLF 

		' call the event handler
		If boolEvtTerminate Then
			on error resume next
			Call evtTerminate
			on error goto 0
		End If
	End Sub
End Class

'
'	Encapsulates a Button object
'
Class Button
	Private objParent
	Private strName
	Private strValue
	Private strStyle
	Private strClass
	Private boolVisible
	Private boolSerialize

	Private evtClick
	Private boolEvtClick

	' constructor
	Private Sub Class_Initialize
		strName = ""
		strValue = ""
		strStyle = ""
		boolVisible = True
		boolSerialize = True
		boolEvtClick = False
	End Sub

	' global event handler for this element
	Public Sub EventHandler( sender, arguments )
		If arguments = "click" Then
			If boolEvtClick Then
				Call evtClick( sender, arguments )
			End If
		End If
	End Sub
	
	' set the click event handler for this object
	Public Property Let OnClick( ByRef subRef ) : Set evtClick = subRef : boolEvtClick = True : End Property

	' type of this object
	Public Property Get IsType : IsType = "button" : End Property

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property

	' the required name of the object
	Public Property Get Name : Name = strName : End Property
	Public Property Let Name( str ) : strName = str : End Property

	' the current value of the object
	Public Property Get Value : Value = strValue : End Property	
	Public Property Let Value( str ) : strValue = str : End Property

	' the required name of the object
	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style( str ) : strStyle = str : End Property
	
	' get/set the CSS class name
	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName ( str ) : strClass = str : End Property

	' get/set visibility
	Public Property Get Visible : Visible = boolVisible : End Property
	Public Property Let Visible( bool )
		If bool and LCase( bool ) <> "false" Then
			boolVisible = True
		Else
			boolVisible = False
		End If
	End Property
	
	' get/set serialization behavior
	Public Property Get CanSerialize : CanSerialize = boolSerialize : End Property
	Public Property Let CanSerialize( bool )
		If bool and LCase( bool ) <> "false" Then
			boolSerialize = True
		Else
			boolSerialize = False
		End If
	End Property
	
	' persists it큦 current state in a XML stream
	Public Function Serialize
		Dim strXML
		strXML = strName

		If Not boolSerialize Then
			strXML = strXML & _
			Parent.ChildSerializeNode( "serialize", "false" )
		Else
			strXML = strXML & _
			Parent.ChildSerializeNode( "name", strName ) & _
			Parent.ChildSerializeNode( "value", strValue ) & _
			Parent.ChildSerializeNode( "style", strStyle ) & _
			Parent.ChildSerializeNode( "class", strClass ) & _
			Parent.ChildSerializeNode( "visible", boolVisible )
		End If

		Serialize = strXML
	End Function
	
	' wraps the object큦 properties
	Public Sub SetProperty( strKey, strArg )
		Select Case strKey
		Case "name"
			Name = strArg
		Case "value"
			Value = strArg
		Case "style"
			Style = strArg
		Case "class"
			ClassName = strArg
		Case "visible"
			Visible = strArg
		End Select
	End Sub
	
	' prints it큦 current HTML element
	Public Sub Render
		If Not Visible Then
			Exit Sub
		End If
		
		Dim strHTML, strComplement
		
		If boolEvtClick Then
			strComplement = "onclick=""_doPostBack( '%name', 'click' )"" "
		End If
		
		strHTML = "<input type=""button"" id=""%name"" name=""%name""%style%class value=""%value"" %complement/>"
		strHTML = Replace( strHTML, "%complement", strComplement )
		strHTML = Replace( strHTML, "%value", Replace( strValue, """", "\""", 1, -1 ) )
		strHTML = Replace( strHTML, "%name", strName, 1, -1 )
		If strStyle <> "" Then
			strHTML = Replace( strHTML, "%style", " style=""" & strStyle & """" )
		Else
			strHTML = Replace( strHTML, "%style", "" )
		End If
		If strClass <> "" Then
			strHTML = Replace( strHTML, "%class", " class=""" & strClass & """" )
		Else
			strHTML = Replace( strHTML, "%class", "" )
		End If
		
		Response.Write strHTML
	End Sub
	
End Class

'
' Encapsulates a Label object
'
Class Label
	Private objParent
	Private strName
	Private strValue
	Private strStyle
	Private strClass
	Private boolVisible
	Private boolSerialize

	' constructor
	Private Sub Class_Initialize
		strName = ""
		strValue = ""
		strStyle = ""
		boolVisible = True
		boolSerialize = True
	End Sub

	' dummy event handler
	Public Sub EventHandler( sender, arguments ) : End Sub

	' type of this object
	Public Property Get IsType : IsType = "label" : End Property

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property

	' get/set required name for this element
	Public Property Get Name : Name = strName : End Property
	Public Property Let Name ( str ) : strName = str : End Property

	' get/set the value
	Public Property Get Value : Value = strValue : End Property
	Public Property Let Value ( str ) : strValue = str : End Property

	' get/set the style
	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style ( str ) : strStyle = str : End Property

	' get/set the CSS class name
	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName ( str ) : strClass = str : End Property

	' get/set visibility
	Public Property Get Visible : Visible = boolVisible : End Property
	Public Property Let Visible( bool )
		If bool and LCase( bool ) <> "false" Then
			boolVisible = True
		Else
			boolVisible = False
		End If
	End Property

	' get/set serialization behavior
	Public Property Get CanSerialize : CanSerialize = boolSerialize : End Property
	Public Property Let CanSerialize( bool )
		If bool and LCase( bool ) <> "false" Then
			boolSerialize = True
		Else
			boolSerialize = False
		End If
	End Property

	' serializes
	Public Function Serialize
		Dim strXML
		strXML = strName

		If Not boolSerialize Then
			strXML = strXML & _
			Parent.ChildSerializeNode( "serialize", "false" )
		Else
			strXML = strXML & _
			Parent.ChildSerializeNode( "name", strName ) & _
			Parent.ChildSerializeNode( "value", strValue ) & _
			Parent.ChildSerializeNode( "style", strStyle ) & _
			Parent.ChildSerializeNode( "class", strClass ) & _
			Parent.ChildSerializeNode( "visible", boolVisible )
		End If

		Serialize = strXML
	End Function
	
	' called by the parent Page class and must be able to set
	' all of it큦 own properties
	Public Sub SetProperty( strKey, strArg )
		Select Case strKey
		Case "name"
			Name = strArg
		Case "value"
			Value = strArg
		Case "style"
			Style = strArg
		Case "class"
			ClassName = strArg
		Case "visible"
			Visible = strArg
		End Select
	End Sub
	
	' renders the HTML control in the page
	Public Sub Render
		If Not Visible Then
			Exit Sub
		End If
		
		Dim strHTML
		strHTML = "<span id=""%name""%style%class>%value</span>"
		strHTML = Replace( strHTML, "%value", strValue )
		strHTML = Replace( strHTML, "%name", strName, 1, -1 )
		If strStyle <> "" Then
			strHTML = Replace( strHTML, "%style", " style=""" & strStyle & """" )
		Else
			strHTML = Replace( strHTML, "%style", "" )
		End If
		If strClass <> "" Then
			strHTML = Replace( strHTML, "%class", " class=""" & strClass & """" )
		Else
			strHTML = Replace( strHTML, "%class", "" )
		End If
		
		Response.Write strHTML
	End Sub
	
End Class

'
' Encapsulates a TextBox object
'
Class TextBox
	Private objParent
	Private strName
	Private strValue
	Private strStyle
	Private strClass
	Private intMaxLength
	Private intRows
	Private intCols
	Private boolVisible
	Private boolSerialize
	Private boolPassword
	
	Private evtChange
	Private boolEvtChange

	' constructor
	Private Sub Class_Initialize
		strName = ""
		strValue = ""
		strStyle = ""
		intMaxLength = 0
		intRows = 1
		intCols = 15
		boolVisible = True
		boolSerialize = True
		boolPassword = False
		boolEvtChange = False
	End Sub

	' global event handler for this element
	Public Sub EventHandler( sender, arguments )
		If arguments = "change" Then
			If boolEvtChange Then
				Call evtChange( sender, arguments )
			End If
		End If
	End Sub
	
	' set the click event handler for this object
	Public Property Let OnChange( ByRef subRef ) : Set evtChange = subRef : boolEvtChange = True : End Property

	' type of this object
	Public Property Get IsType : IsType = "textbox" : End Property

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property

	' get/set required name for this element
	Public Property Get Name : Name = strName : End Property
	Public Property Let Name ( str ) : strName = str : End Property

	' get/set the value
	Public Property Get Value : Value = strValue : End Property
	Public Property Let Value ( str ) : strValue = str : End Property

	' get/set the style
	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style ( str ) : strStyle = str : End Property

	' get/set the CSS class name
	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName ( str ) : strClass = str : End Property

	' get/set the field max char length
	Public Property Get MaxLength : MaxLength = intMaxLength : End Property
	Public Property Let MaxLength ( intTmp ) 
		If IsNumeric( intTmp ) Then
			on error resume next
			intMaxLength = CInt( intTmp )
			If Err.number <> 0 Then
				intMaxLength = 0
			End If
			on error goto 0
		End If
	End Property
	
	' get/set the number of rows
	Public Property Get Rows : Rows = intRows : End Property
	Public Property Let Rows ( intCount ) 
		If IsNumeric( intCount ) Then
			on error resume next
			intRows = CInt( intCount )
			If Err.number <> 0 Then
				intRows = 1
			End If
			on error goto 0
		End If
	End Property

	' get/set the number of columns
	Public Property Get Columns : Columns = intCols : End Property
	Public Property Let Columns ( intCount ) 
		If IsNumeric( intCount ) Then
			on error resume next
			intCols = CInt( intCount )
			If Err.number <> 0 Then
				intCols = 15
			End If
			on error goto 0
		End If
	End Property

	' get/set visibility
	Public Property Get Visible : Visible = boolVisible : End Property
	Public Property Let Visible( bool )
		If bool and LCase( bool ) <> "false" Then
			boolVisible = True
		Else
			boolVisible = False
		End If
	End Property

	' get/set serialization behavior
	Public Property Get CanSerialize : CanSerialize = boolSerialize : End Property
	Public Property Let CanSerialize( bool )
		If bool and LCase( bool ) <> "false" Then
			boolSerialize = True
		Else
			boolSerialize = False
		End If
	End Property

	' get/set if it큦 a password input box
	Public Property Get IsPassword : IsPassword = boolPassword : End Property
	Public Property Let IsPassword( bool )
		If bool and LCase( bool ) <> "false" Then
			boolPassword = True
		Else
			boolPassword = False
		End If
	End Property
			
	' serializes
	Public Function Serialize
		Dim strXML
		strXML = strName

		If Not boolSerialize Then
			strXML = strXML & _
			Parent.ChildSerializeNode( "serialize", "false" )
		Else
			strXML = strXML & _
			Parent.ChildSerializeNode( "name", strName ) & _
			Parent.ChildSerializeNode( "value", strValue ) & _
			Parent.ChildSerializeNode( "maxlength", intMaxLength ) & _
			Parent.ChildSerializeNode( "style", strStyle ) & _
			Parent.ChildSerializeNode( "class", strClass ) & _
			Parent.ChildSerializeNode( "rows", intRows ) & _
			Parent.ChildSerializeNode( "cols", intCols ) & _
			Parent.ChildSerializeNode( "visible", boolVisible )
		End If
		
		Serialize = strXML
	End Function
	
	' called by the parent Page class and must be able to set
	' all of it큦 own properties
	Public Sub SetProperty( strKey, strArg )
		Select Case strKey
		Case "name"
			Name = strArg
		Case "value"
			Value = strArg
		Case "maxlength"
			MaxLength = strArg
		Case "style"
			Style = strArg
		Case "visible"
			Visible = strArg
		Case "rows"
			Rows = strArg
		Case "cols"
			Columns = strArg
		End Select
	End Sub
	
	' renders the HTML control in the page
	Public Sub Render
		If Not Visible Then
			Exit Sub
		End If

		' password fields can큧 be text based		
		If IsPassword Then
			intRows = 1
		End If
		
		Dim strHTML, sComplement
		If intRows = "1" Then
			strHTML = "<input type=""%type"" id=""%name"" name=""%name"" size=""%cols"" value=""%value""%style%class%complement />"
		Else
			strHTML = "<textarea id=""%name"" name=""%name"" rows=""%rows"" cols=""%cols""%style%class%complement>%value</textarea>"
		End If
		sComplement = ""
		If boolEvtChange Then
			sComplement = " onchange=""_doPostBack( '%name', 'change' )"" "
		End If
		If intMaxLength > 0 Then
			sComplement = " maxlength=""" & intMaxLength & """ " & sComplement
		End If
		strHTML = Replace( strHTML, "%complement", sComplement )
		strHTML = Replace( strHTML, "%rows", intRows )
		strHTML = Replace( strHTML, "%cols", intCols )
		If Not IsPassword Then
			strHTML = Replace( strHTML, "%value", Replace( strValue, """", "\""", 1, -1 ) )
			strHTML = Replace( strHTML, "%type", "text" )
		Else
			strHTML = Replace( strHTML, "%value", "" )
			strHTML = Replace( strHTML, "%type", "password" )
		End If
		strHTML = Replace( strHTML, "%name", strName, 1, -1 )
		If strStyle <> "" Then
			strHTML = Replace( strHTML, "%style", " style=""" & strStyle & """ " )
		Else
			strHTML = Replace( strHTML, "%style", "" )
		End If
		If strClass <> "" Then
			strHTML = Replace( strHTML, "%class", " class=""" & strClass & """ " )
		Else
			strHTML = Replace( strHTML, "%class", "" )
		End If
		
		Response.Write strHTML
	End Sub
	
End Class

'
'	Encapsulates the functionallity of a drop down list
'
Class DropDownList
	Private objParent
	Private objList
	Private arrSelected
	Private strName
	Private strValue
	Private strStyle
	Private strClass
	Private intRows
	Private boolMultiple
	Private boolVisible
	Private boolSerialize
	Private constElementDivider
	Private constTokenDivider

	Private evtChange
	Private boolEvtChange

	' constructor
	Private Sub Class_Initialize
		strName = ""
		strValue = ""
		strStyle = ""
		intRows = 1
		boolMultiple = False
		boolVisible = True
		boolSerialize = True

		constElementDivider = "#@#"
		constTokenDivider = "@#@"

		boolEvtChange = False
		
		Set objList = Server.CreateObject( "Scripting.Dictionary" )
		objList.RemoveAll
	End Sub
	
	Private Sub Class_Terminate
		Set objList = Nothing
	End Sub

	' global event handler for this element
	Public Sub EventHandler( sender, arguments )
		If arguments = "change" Then
			If boolEvtChange Then
				Call evtChange( sender, arguments )
			End If
		End If
	End Sub
	
	' set the click event handler for this object
	Public Property Let OnChange( ByRef subRef ) : Set evtChange = subRef : boolEvtChange = True : End Property

	' type of this object
	Public Property Get IsType : IsType = "dropdown" : End Property

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property

	' the required name of the object
	Public Property Get Name : Name = strName : End Property
	Public Property Let Name( str ) : strName = str : End Property

	' the required name of the object
	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style( str ) : strStyle = str : End Property
	
	' get/set the CSS class name
	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName ( str ) : strClass = str : End Property

	' get/set the Items (Dictionary) object
	' any Dictionary can be connected here, but if there큦 no dictionary the
	' object will categorically fail without warning
	Public Property Get Items : Set Items = objList : End Property
	Public Property Let Items ( ByRef obj ) : Set objList = obj : End Property

	' get/set visibility
	Public Property Get Visible : Visible = boolVisible : End Property
	Public Property Let Visible( bool )
		If bool and LCase( bool ) <> "false" Then
			boolVisible = True
		Else
			boolVisible = False
		End If
	End Property
	
	' get/set serialization behavior
	Public Property Get CanSerialize : CanSerialize = boolSerialize : End Property
	Public Property Let CanSerialize( bool )
		If bool and LCase( bool ) <> "false" Then
			boolSerialize = True
		Else
			boolSerialize = False
		End If
	End Property

	' get/set multiple select
	Public Property Get AllowMultiple : AllowMultiple = boolMultiple : End Property
	Public Property Let AllowMultiple( bool )
		If bool and LCase( bool ) <> "false" Then
			boolMultiple = True
		Else
			boolMultiple = False
		End If
	End Property

	' the current selected indexes
	Public Property Get Value 
		If IsArray( arrSelected ) Then
			Value = arrSelected
		End If
	End Property
	Public Property Let Value( str ) 
		str = Replace( str, ", ", ",", 1, -1 )
		If InStr( str, "," ) Then
			arrSelected = Split( str, "," )
		Else
			arrSelected = Array( str )
		End If
		Dim intCount
		on error resume next
		For intCount = 0 To UBound( arrSelected )
			arrSelected( intCount ) = CInt( arrSelected( intCount ) )
		Next
		on error goto 0
	End Property
	
	' retrieves a value for it큦 Index
	Public Function GetValue( index )
		If IsNumeric( index ) and index < objList.Count Then
			Dim arrTmp
			arrTmp = objList.Items
			GetValue = arrTmp( index )
		Else
			If objList.Exists( index ) Then
				GetValue = objList( index )
			End If
		End If
	End Function

	' retrieves a key for it큦 index
	Public Function GetKey( index )
		If IsNumeric( index ) and index < objList.Count Then
			Dim arrTmp
			arrTmp = objList.Keys
			GetKey = arrTmp( index )
		End If
	End Function

	' total number of returned selected elements	
	Public Property Get SelectedCount 
		SelectedCount = 0
		If IsArray( arrSelected ) Then
			SelectedCount = CInt( UBound( arrSelected ) ) + 1
		End If
	End Property
	
	' total number of elements
	Public Property Get Count : Count = objList.Count : End Property

	' get/set number of rows
	Public Property Get Rows : Rows = intRows : End Property
	Public Property Let Rows( intNumber )
		If IsNumeric( intNumber ) Then
			on error resume next
			intNumber = CInt( intNumber )
			If Err.number <> 0 or intNumber < 1 Then
				intNumber = 1
			End If
			on error goto 0
			
			intRows = intNumber
		End If
	End Property
	
	' persists it큦 current state in a XML stream
	Public Function Serialize
		Dim strXML
		strXML = strName

		If Not boolSerialize Then
			strXML = strXML & _
			Parent.ChildSerializeNode( "serialize", "false" )
		Else
			strXML = strXML & _
			Parent.ChildSerializeNode( "name", strName ) & _
			Parent.ChildSerializeNode( "value", strValue ) & _
			Parent.ChildSerializeNode( "style", strStyle ) & _
			Parent.ChildSerializeNode( "class", strClass ) & _
			Parent.ChildSerializeNode( "visible", boolVisible ) & _
			Parent.ChildSerializeNode( "multiple", boolMultiple ) & _
			Parent.ChildSerializeNode( "rows", intRows ) 
			
			' serialize the list values
			Dim strKey, strList
			For Each strKey in objList
				strList = strList & strKey & constTokenDivider & objList( strKey ) & constElementDivider
			Next
			If strList <> "" Then
				strList = Left( strList, Len( strList ) - Len( constElementDivider ) )
				strXML = strXML & Parent.ChildSerializeNode( "list", strList )
			End If
		End If

		Serialize = strXML
	End Function
	
	' wraps the object큦 properties
	Public Sub SetProperty( strKey, strArg )
		Select Case strKey
		Case "name"
			Name = strArg
		Case "value"
			Value = strArg
		Case "style"
			Style = strArg
		Case "class"
			ClassName = strArg
		Case "visible"
			Visible = strArg
		Case "multiple"
			AllowMultiple = strArg
		Case "rows"
			Rows = strArg
		Case "list"
			' deserialize list
			Dim arrList, intCount, arrElement
			arrList = Split( strArg, constElementDivider )
			objList.RemoveAll
			For intCount = 0 To UBound( arrList )
				arrElement = Split( arrList( intCount ), constTokenDivider )
				If IsArray( arrElement ) Then
					If UBound( arrElement ) = 1 Then
						objList.Add arrElement( 0 ), arrElement( 1 )
					End If
				End If
			Next
		End Select
	End Sub
	
	' prints it큦 current HTML element
	Public Sub Render
		If Not Visible Then
			Exit Sub
		End If
		
		Dim strHTML, strComplement
		
		If intRows > 1 Then
			strComplement = " size=""" & intRows & """ "
		End If
		If boolEvtChange Then
			strComplement = strComplement & " onchange=""_doPostBack( '%name', 'change' )"" "
		End If
		
		If boolMultiple Then
			strComplement = strComplement & " multiple=""multiple"" "
		End If
		
		strHTML = "<select id=""%name"" name=""%name""%style%class %complement>" & vbCRLF
		strHTML = Replace( strHTML, "%complement", strComplement )
		strHTML = Replace( strHTML, "%name", strName, 1, -1 )
		If strStyle <> "" Then
			strHTML = Replace( strHTML, "%style", " style=""" & strStyle & """" )
		Else
			strHTML = Replace( strHTML, "%style", "" )
		End If
		If strClass <> "" Then
			strHTML = Replace( strHTML, "%class", " class=""" & strClass & """" )
		Else
			strHTML = Replace( strHTML, "%class", "" )
		End If
		
		' print the items
		Dim strSelected, intTmp, intCount
		intCount = 0
		If objList.Count > 0 Then
			For Each strKey in objList
				strSelected = ""
				If IsArray( arrSelected ) Then
					For Each intTmp in arrSelected
						If intTmp = intCount Then
							strSelected = " selected=""selected"""
							Exit For
						End If
					Next
				End If
				strHTML = strHTML & "<option value=""" & intCount & """" & strSelected & ">" & strKey & "</option>" & vbCRLF
				intCount = intCount + 1
			Next
		End If
		
		strHTML = strHTML & "</select>" & vbCRLF
		Response.Write strHTML
	End Sub

End Class

'
'	Encapsulates the functionallity of a group set of radio or check boxes
'
Class List
	Private objParent
	Private objList
	Private arrSelected
	Private strName
	Private strValue
	Private strStyle
	Private strClass
	Private boolVertical
	Private boolMultiple
	Private boolVisible
	Private boolSerialize
	Private constElementDivider
	Private constTokenDivider

	Private evtChange
	Private boolEvtChange

	' constructor
	Private Sub Class_Initialize
		strName = ""
		strValue = ""
		strStyle = ""
		boolMultiple = False
		boolVertical = True
		boolVisible = True
		boolSerialize = True

		constElementDivider = "#@#"
		constTokenDivider = "@#@"

		boolEvtChange = False
		
		Set objList = Server.CreateObject( "Scripting.Dictionary" )
		objList.RemoveAll
	End Sub
	
	Private Sub Class_Terminate
		Set objList = Nothing
	End Sub

	' global event handler for this element
	Public Sub EventHandler( sender, arguments )
		If arguments = "change" Then
			If boolEvtChange Then
				Call evtChange( sender, arguments )
			End If
		End If
	End Sub
	
	' set the click event handler for this object
	Public Property Let OnChange( ByRef subRef ) : Set evtChange = subRef : boolEvtChange = True : End Property

	' type of this object
	Public Property Get IsType 
		If boolMultiple Then
			IsType = "checklist"
		Else
			IsType = "radiolist"
		End If
	End Property

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property

	' the required name of the object
	Public Property Get Name : Name = strName : End Property
	Public Property Let Name( str ) : strName = str : End Property

	' the required name of the object
	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style( str ) : strStyle = str : End Property
	
	' get/set the CSS class name
	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName ( str ) : strClass = str : End Property

	' get/set the Items (Dictionary) object
	' any Dictionary can be connected here, but if there큦 no dictionary the
	' object will categorically fail without warning
	Public Property Get Items : Set Items = objList : End Property
	Public Property Let Items ( ByRef obj ) : Set objList = obj : End Property

	' get/set visibility
	Public Property Get Visible : Visible = boolVisible : End Property
	Public Property Let Visible( bool )
		If bool and LCase( bool ) <> "false" Then
			boolVisible = True
		Else
			boolVisible = False
		End If
	End Property
	
	' get/set serialization behavior
	Public Property Get CanSerialize : CanSerialize = boolSerialize : End Property
	Public Property Let CanSerialize( bool )
		If bool and LCase( bool ) <> "false" Then
			boolSerialize = True
		Else
			boolSerialize = False
		End If
	End Property

	' get/set multiple select
	Public Property Get AllowMultiple : AllowMultiple = boolMultiple : End Property
	Public Property Let AllowMultiple( bool )
		If bool and LCase( bool ) <> "false" Then
			boolMultiple = True
		Else
			boolMultiple = False
		End If
	End Property

	' get/set direction of the rendering
	Public Property Get Vertical : Vertical = boolVertical : End Property
	Public Property Let Vertical( bool )
		If bool and LCase( bool ) <> "false" Then
			boolVertical = True
		Else
			boolVertical = False
		End If
	End Property
	
	' the current selected values
	Public Property Get Value 
		If IsArray( arrSelected ) Then
			Value = arrSelected
		End If
	End Property
	Public Property Let Value( str ) 
		str = Replace( str, ", ", ",", 1, -1 )
		If InStr( str, "," ) Then
			arrSelected = Split( str, "," )
		Else
			arrSelected = Array( str )
		End If
		Dim intCount
		on error resume next
		For intCount = 0 To UBound( arrSelected )
			arrSelected( intCount ) = CInt( arrSelected( intCount ) )
		Next
		on error goto 0
	End Property

	' retrieves a value for it큦 Index
	Public Function GetValue( index )
		If IsNumeric( index ) and index < objList.Count Then
			Dim arrTmp
			arrTmp = objList.Items
			GetValue = arrTmp( index )
		Else
			If objList.Exists( index ) Then
				GetValue = objList( index )
			End If
		End If
	End Function

	' retrieves a key for it큦 index
	Public Function GetKey( index )
		If IsNumeric( index ) and index < objList.Count Then
			Dim arrTmp
			arrTmp = objList.Keys
			GetKey = arrTmp( index )
		End If
	End Function

	' total number of returned selected elements	
	Public Property Get SelectedCount 
		SelectedCount = 0
		If IsArray( arrSelected ) Then
			SelectedCount = CInt( UBound( arrSelected ) ) + 1
		End If
	End Property
	
	' total number of elements
	Public Property Get Count : Count = objList.Count : End Property

	' persists it큦 current state in a XML stream
	Public Function Serialize
		Dim strXML
		strXML = strName

		If Not boolSerialize Then
			strXML = strXML & _
			Parent.ChildSerializeNode( "serialize", "false" )
		Else
			strXML = strXML & _
			Parent.ChildSerializeNode( "name", strName ) & _
			Parent.ChildSerializeNode( "value", strValue ) & _
			Parent.ChildSerializeNode( "style", strStyle ) & _
			Parent.ChildSerializeNode( "class", strClass ) & _
			Parent.ChildSerializeNode( "visible", boolVisible ) & _
			Parent.ChildSerializeNode( "multiple", boolMultiple ) & _
			Parent.ChildSerializeNode( "vertical", boolVertical )
			
			' serialize the list values
			Dim strKey, strList
			For Each strKey in objList
				strList = strList & strKey & constTokenDivider & objList( strKey ) & constElementDivider
			Next
			If strList <> "" Then
				strList = Left( strList, Len( strList ) - Len( constElementDivider ) )
				strXML = strXML & Parent.ChildSerializeNode( "list", strList )
			End If
		End If

		Serialize = strXML
	End Function
	
	' wraps the object큦 properties
	Public Sub SetProperty( strKey, strArg )
		Select Case strKey
		Case "name"
			Name = strArg
		Case "value"
			Value = strArg
		Case "style"
			Style = strArg
		Case "class"
			ClassName = strArg
		Case "visible"
			Visible = strArg
		Case "multiple"
			AllowMultiple = strArg
		Case "vertical"
			Vertical = strArg
		Case "list"
			' deserialize list
			Dim arrList, intCount, arrElement
			arrList = Split( strArg, constElementDivider )
			objList.RemoveAll
			For intCount = 0 To UBound( arrList )
				arrElement = Split( arrList( intCount ), constTokenDivider )
				If IsArray( arrElement ) Then
					If UBound( arrElement ) = 1 Then
						objList.Add arrElement( 0 ), arrElement( 1 )
					End If
				End If
			Next
		End Select
	End Sub
	
	' prints it큦 current HTML element
	Public Sub Render
		If Not Visible Then
			Exit Sub
		End If
		
		Dim strHTML, strElement, strComplement
		If boolEvtChange Then
			strComplement = strComplement & " onchange=""_doPostBack( '%name', 'change' )"" "
		End If
		
		strElement = "<input type=""%type"" id=""%name"" name=""%name"" value=""%value"" %style%class%complement%selected/> %key"
		strElement = Replace( strElement, "%complement", strComplement )
		strElement = Replace( strElement, "%name", strName, 1, -1 )
		If strStyle <> "" Then
			strElement = Replace( strElement, "%style", " style=""" & strStyle & """" )
		Else
			strElement = Replace( strElement, "%style", "" )
		End If
		If strClass <> "" Then
			strElement = Replace( strElement, "%class", " class=""" & strClass & """" )
		Else
			strElement = Replace( strElement, "%class", "" )
		End If
		If boolMultiple Then
			strElement = Replace( strElement, "%type", "checkbox" )
		Else
			strElement = Replace( strElement, "%type", "radio" )
		End If
		If boolVertical Then
			strElement = strElement & "<br />"
		Else
			strElement = strElement & "&nbsp;"
		End If
		strElement = strElement & vbCRLF
		
		' print the items
		Dim strSelected, intTmp, intCount
		intCount = 0
		If objList.Count > 0 Then
			For Each strKey in objList
				strSelected = ""
				If IsArray( arrSelected ) Then
					For Each intTmp in arrSelected
						If intCount = intTmp Then
							strSelected = " checked=""checked"""
							Exit For
						End If
					Next
				End If
				strTmp = Replace( strElement, "%value", intCount )
				strTmp = Replace( strTmp, "%key", strKey )
				strTmp = Replace( strTmp, "%selected", strSelected )
				strHTML = strHTML & strTmp
				intCount = intCount + 1
			Next
		End If
		
		Response.Write strHTML
	End Sub
End Class

'
'	Encapsulates a Button object
'
Class FileUpload
	Private objParent
	Private strName
	Private strValue
	Private strStyle
	Private strClass
	Private intSize
	Private intMaxLength
	Private boolVisible
	Private boolSerialize

	Private evtChange
	Private boolEvtChange

	' constructor
	Private Sub Class_Initialize
		strName = ""
		strValue = ""
		strStyle = ""
		intSize = 0
		intMaxLength = 0
		boolVisible = True
		boolSerialize = True
		boolEvtChange = False
	End Sub

	' global event handler for this element
	Public Sub EventHandler( sender, arguments )
		If arguments = "change" Then
			If boolEvtChange Then
				Call evtChange( sender, arguments )
			End If
		End If
	End Sub
	
	' set the click event handler for this object
	Public Property Let OnChange( ByRef subRef ) : Set evtChange = subRef : boolEvtChange = True : End Property

	' type of this object
	Public Property Get IsType : IsType = "fileupload" : End Property

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property

	' the required name of the object
	Public Property Get Name : Name = strName : End Property
	Public Property Let Name( str ) : strName = str : End Property

	' the current value of the object
	Public Property Get Value : Value = strValue : End Property	
	Public Property Let Value( str ) : strValue = str : End Property

	' the required name of the object
	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style( str ) : strStyle = str : End Property
	
	' get/set the CSS class name
	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName ( str ) : strClass = str : End Property

	' get/set visibility
	Public Property Get Visible : Visible = boolVisible : End Property
	Public Property Let Visible( bool )
		If bool and LCase( bool ) <> "false" Then
			boolVisible = True
		Else
			boolVisible = False
		End If
	End Property

	' get/set the char size
	Public Property Get Size : Size = intSize : End Property
	Public Property Let Size ( intTmp ) 
		If IsNumeric( intTmp ) Then
			on error resume next
			intSize = CInt( intTmp )
			If Err.number <> 0 Then
				intSize = 0
			End If
			on error goto 0
		End If
	End Property

	' get/set the field max char length
	Public Property Get MaxLength : MaxLength = intMaxLength : End Property
	Public Property Let MaxLength ( intTmp ) 
		If IsNumeric( intTmp ) Then
			on error resume next
			intMaxLength = CInt( intTmp )
			If Err.number <> 0 Then
				intMaxLength = 0
			End If
			on error goto 0
		End If
	End Property
	
	' get/set serialization behavior
	Public Property Get CanSerialize : CanSerialize = boolSerialize : End Property
	Public Property Let CanSerialize( bool )
		If bool and LCase( bool ) <> "false" Then
			boolSerialize = True
		Else
			boolSerialize = False
		End If
	End Property
	
	' persists it큦 current state in a XML stream
	Public Function Serialize
		Dim strXML
		strXML = strName

		If Not boolSerialize Then
			strXML = strXML & _
			Parent.ChildSerializeNode( "serialize", "false" )
		Else
			strXML = strXML & _
			Parent.ChildSerializeNode( "name", strName ) & _
			Parent.ChildSerializeNode( "value", strValue ) & _
			Parent.ChildSerializeNode( "size", intSize ) & _
			Parent.ChildSerializeNode( "maxlength", intMaxLength ) & _
			Parent.ChildSerializeNode( "style", strStyle ) & _
			Parent.ChildSerializeNode( "class", strClass ) & _
			Parent.ChildSerializeNode( "visible", boolVisible )
		End If

		Serialize = strXML
	End Function
	
	' wraps the object큦 properties
	Public Sub SetProperty( strKey, strArg )
		Select Case strKey
		Case "name"
			Name = strArg
		Case "value"
			Value = strArg
		Case "size"
			Size = strArg
		Case "maxlength"
			MaxLength = strArg
		Case "style"
			Style = strArg
		Case "class"
			ClassName = strArg
		Case "visible"
			Visible = strArg
		End Select
	End Sub
	
	' prints it큦 current HTML element
	Public Sub Render
		If Not Visible Then
			Exit Sub
		End If
		
		' only allow uploads if the form allows it too
		If Not objParent.AllowUpload Then
			Exit Sub
		End If
		
		Dim strHTML, strComplement
		
		If boolEvtChange Then
			strComplement = "onchange=""_doPostBack( '%name', 'change' )"" "
		End If
		If intMaxLength > 0 Then
			strComplement = "maxlength=""" & intMaxLength & """ " & sComplement
		End If
		If intSize > 0 Then
			strComplement = "size=""" & intSize & """ " & sComplement
		End If
		
		strHTML = "<input type=""file"" id=""%name"" name=""%name""%style%class value=""%value"" %complement/>"
		strHTML = Replace( strHTML, "%complement", strComplement )
		strHTML = Replace( strHTML, "%value", Replace( strValue, """", "\""", 1, -1 ) )
		strHTML = Replace( strHTML, "%name", strName, 1, -1 )
		If strStyle <> "" Then
			strHTML = Replace( strHTML, "%style", " style=""" & strStyle & """" )
		Else
			strHTML = Replace( strHTML, "%style", "" )
		End If
		If strClass <> "" Then
			strHTML = Replace( strHTML, "%class", " class=""" & strClass & """" )
		Else
			strHTML = Replace( strHTML, "%class", "" )
		End If
		
		Response.Write strHTML
	End Sub

End Class

'
'	Encapsulates an Image object
'
Class Image
	Private objParent
	Private strName
	Private strValue
	Private strStyle
	Private strClass
	Private strToopTip
	Private boolVisible
	Private boolSerialize

	Private evtClick
	Private boolEvtClick

	' constructor
	Private Sub Class_Initialize
		strName = ""
		strValue = ""
		strStyle = ""
		strToolTip = ""
		boolVisible = True
		boolSerialize = True
		boolEvtClick = False
	End Sub

	' global event handler for this element
	Public Sub EventHandler( sender, arguments )
		If arguments = "click" Then
			If boolEvtClick Then
				Call evtClick( sender, arguments )
			End If
		End If
	End Sub
	
	' set the click event handler for this object
	Public Property Let OnClick( ByRef subRef ) : Set evtClick = subRef : boolEvtClick = True : End Property

	' type of this object
	Public Property Get IsType : IsType = "image" : End Property

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property

	' the required name of the object
	Public Property Get Name : Name = strName : End Property
	Public Property Let Name( str ) : strName = str : End Property

	' the current value of the object
	Public Property Get Value : Value = strValue : End Property	
	Public Property Let Value( str ) : strValue = str : End Property

	' tool tip
	Public Property Get ToolTip : ToolTip = strToolTip : End Property	
	Public Property Let ToolTip( str ) : strToolTip = str : End Property

	' wraps the "Value" as a "Path"
	Public Property Get Path : Path = strValue : End Property	
	Public Property Let Path( str ) : strValue = str : End Property

	' the required name of the object
	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style( str ) : strStyle = str : End Property
	
	' get/set the CSS class name
	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName ( str ) : strClass = str : End Property

	' get/set visibility
	Public Property Get Visible : Visible = boolVisible : End Property
	Public Property Let Visible( bool )
		If bool and LCase( bool ) <> "false" Then
			boolVisible = True
		Else
			boolVisible = False
		End If
	End Property
	
	' get/set serialization behavior
	Public Property Get CanSerialize : CanSerialize = boolSerialize : End Property
	Public Property Let CanSerialize( bool )
		If bool and LCase( bool ) <> "false" Then
			boolSerialize = True
		Else
			boolSerialize = False
		End If
	End Property
	
	' persists it큦 current state in a XML stream
	Public Function Serialize
		Dim strXML
		strXML = strName

		If Not boolSerialize Then
			strXML = strXML & _
			Parent.ChildSerializeNode( "serialize", "false" )
		Else
			strXML = strXML & _
			Parent.ChildSerializeNode( "name", strName ) & _
			Parent.ChildSerializeNode( "value", strValue ) & _
			Parent.ChildSerializeNode( "tip", strToolTip ) & _
			Parent.ChildSerializeNode( "style", strStyle ) & _
			Parent.ChildSerializeNode( "class", strClass ) & _
			Parent.ChildSerializeNode( "visible", boolVisible )
		End If

		Serialize = strXML
	End Function
	
	' wraps the object큦 properties
	Public Sub SetProperty( strKey, strArg )
		Select Case strKey
		Case "name"
			Name = strArg
		Case "value"
			Value = strArg
		Case "tip"
			ToolTip = strArg
		Case "style"
			Style = strArg
		Case "class"
			ClassName = strArg
		Case "visible"
			Visible = strArg
		End Select
	End Sub
	
	' prints it큦 current HTML element
	Public Sub Render
		If Not Visible Then
			Exit Sub
		End If
		
		Dim strHTML, strComplement
		
		If boolEvtClick Then
			strComplement = "onclick=""_doPostBack( '%name', 'click' )"" "
		else
			strComplement = "onclick=""return false;"" "
		End If
		If strToolTip <> "" Then
			strComplement = "alt=""" & strToolTip & """ " & strComplement
		End If
		
		strHTML = "<input type=""image"" id=""%name"" name=""%name""%style%class src=""%value"" %complement/>"
		strHTML = Replace( strHTML, "%complement", strComplement )
		strHTML = Replace( strHTML, "%value", Replace( strValue, """", "\""", 1, -1 ) )
		strHTML = Replace( strHTML, "%name", strName, 1, -1 )
		If strStyle <> "" Then
			strHTML = Replace( strHTML, "%style", " style=""" & strStyle & """" )
		Else
			strHTML = Replace( strHTML, "%style", "" )
		End If
		If strClass <> "" Then
			strHTML = Replace( strHTML, "%class", " class=""" & strClass & """" )
		Else
			strHTML = Replace( strHTML, "%class", "" )
		End If
		
		Response.Write strHTML
	End Sub
	
End Class

'
'	Encapsulates a Hidden data-only object
'
Class Hidden
	Private objParent
	Private strName
	Private strValue
	Private boolVisible
	Private boolSerialize

	' constructor
	Private Sub Class_Initialize
		strName = ""
		strValue = ""
		boolVisible = True
		boolSerialize = True
	End Sub

	' dummy handler
	Public Sub EventHandler( sender, arguments ) : End Sub
	
	' type of this object
	Public Property Get IsType : IsType = "hidden" : End Property

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property

	' the required name of the object
	Public Property Get Name : Name = strName : End Property
	Public Property Let Name( str ) : strName = str : End Property

	' the current value of the object
	Public Property Get Value : Value = strValue : End Property	
	Public Property Let Value( str ) : strValue = str : End Property

	' get/set visibility
	Public Property Get Visible : Visible = boolVisible : End Property
	Public Property Let Visible( bool )
		If bool and LCase( bool ) <> "false" Then
			boolVisible = True
		Else
			boolVisible = False
		End If
	End Property
	
	' get/set serialization behavior
	Public Property Get CanSerialize : CanSerialize = boolSerialize : End Property
	Public Property Let CanSerialize( bool )
		If bool and LCase( bool ) <> "false" Then
			boolSerialize = True
		Else
			boolSerialize = False
		End If
	End Property
	
	' persists it큦 current state in a XML stream
	Public Function Serialize
		Dim strXML
		strXML = strName

		If Not boolSerialize Then
			strXML = strXML & _
			Parent.ChildSerializeNode( "serialize", "false" )
		Else
			strXML = strXML & _
			Parent.ChildSerializeNode( "name", strName ) & _
			Parent.ChildSerializeNode( "value", strValue ) & _
			Parent.ChildSerializeNode( "visible", boolVisible )
		End If

		Serialize = strXML
	End Function
	
	' wraps the object큦 properties
	Public Sub SetProperty( strKey, strArg )
		Select Case strKey
		Case "name"
			Name = strArg
		Case "value"
			Value = strArg
		Case "visible"
			Visible = strArg
		End Select
	End Sub
	
	' prints it큦 current HTML element
	Public Sub Render
		If Not Visible Then
			Exit Sub
		End If
		
		Dim strHTML
		strHTML = "<input type=""hidden"" id=""%name"" name=""%name"" value=""%value"" />"
		strHTML = Replace( strHTML, "%value", Replace( strValue, """", "\""", 1, -1 ) )
		strHTML = Replace( strHTML, "%name", strName, 1, -1 )
		
		Response.Write strHTML
	End Sub
	
End Class

'
'	Encapsulates an HTML link
'
Class Anchor
	Private objParent
	Private strName
	Private strValue
	Private strStyle
	Private strClass
	Private boolVisible
	Private boolSerialize

	Private evtClick
	Private boolEvtClick

	' constructor
	Private Sub Class_Initialize
		strName = ""
		strValue = ""
		strStyle = ""
		boolVisible = True
		boolSerialize = True
		boolEvtClick = False
	End Sub

	' global event handler for this element
	Public Sub EventHandler( sender, arguments )
		If arguments = "click" Then
			If boolEvtClick Then
				Call evtClick( sender, arguments )
			End If
		End If
	End Sub
	
	' set the click event handler for this object
	Public Property Let OnClick( ByRef subRef ) : Set evtClick = subRef : boolEvtClick = True : End Property

	' type of this object
	Public Property Get IsType : IsType = "anchor" : End Property

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property

	' the required name of the object
	Public Property Get Name : Name = strName : End Property
	Public Property Let Name( str ) : strName = str : End Property

	' the current value of the object
	Public Property Get Value : Value = strValue : End Property	
	Public Property Let Value( str ) : strValue = str : End Property

	' the required name of the object
	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style( str ) : strStyle = str : End Property
	
	' get/set the CSS class name
	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName ( str ) : strClass = str : End Property

	' get/set visibility
	Public Property Get Visible : Visible = boolVisible : End Property
	Public Property Let Visible( bool )
		If bool and LCase( bool ) <> "false" Then
			boolVisible = True
		Else
			boolVisible = False
		End If
	End Property
	
	' get/set serialization behavior
	Public Property Get CanSerialize : CanSerialize = boolSerialize : End Property
	Public Property Let CanSerialize( bool )
		If bool and LCase( bool ) <> "false" Then
			boolSerialize = True
		Else
			boolSerialize = False
		End If
	End Property
	
	' persists it큦 current state in a XML stream
	Public Function Serialize
		Dim strXML
		strXML = strName

		If Not boolSerialize Then
			strXML = strXML & _
			Parent.ChildSerializeNode( "serialize", "false" )
		Else
			strXML = strXML & _
			Parent.ChildSerializeNode( "name", strName ) & _
			Parent.ChildSerializeNode( "value", strValue ) & _
			Parent.ChildSerializeNode( "style", strStyle ) & _
			Parent.ChildSerializeNode( "class", strClass ) & _
			Parent.ChildSerializeNode( "visible", boolVisible )
		End If

		Serialize = strXML
	End Function
	
	' wraps the object큦 properties
	Public Sub SetProperty( strKey, strArg )
		Select Case strKey
		Case "name"
			Name = strArg
		Case "value"
			Value = strArg
		Case "style"
			Style = strArg
		Case "class"
			ClassName = strArg
		Case "visible"
			Visible = strArg
		End Select
	End Sub
	
	' prints it큦 current HTML element
	Public Sub Render
		If Not Visible Then
			Exit Sub
		End If
		
		Dim strHTML, strComplement
		
		If boolEvtClick Then
			strComplement = "href=""javascript:_doPostBack( '%name', 'click' )"" "
		Else
			strComplement = "href=""#"" "
		End If
		
		strHTML = "<a id=""%name"" name=""%name""%style%class %complement>%value</a>"
		strHTML = Replace( strHTML, "%complement", strComplement )
		strHTML = Replace( strHTML, "%value", Replace( strValue, """", "\""", 1, -1 ) )
		strHTML = Replace( strHTML, "%name", strName, 1, -1 )
		If strStyle <> "" Then
			strHTML = Replace( strHTML, "%style", " style=""" & strStyle & """" )
		Else
			strHTML = Replace( strHTML, "%style", "" )
		End If
		If strClass <> "" Then
			strHTML = Replace( strHTML, "%class", " class=""" & strClass & """" )
		Else
			strHTML = Replace( strHTML, "%class", "" )
		End If
		
		Response.Write strHTML
	End Sub
	
End Class

%>