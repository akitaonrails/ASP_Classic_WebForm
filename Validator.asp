<%
'
'	Every Validator Class expects the Value property of it´s binded element
'	and also must implement a IsOk, IsPostBack and ErrorMessage properties that an 
'	external manager as the Form class will access
'

' 
'	This is a Container for all validators and it can be added 
'	in a Form.ValidatorContainer property
'

Class ValidatorContainer
	Private arrControls()
	Private objParent
	Private intPointer
	Private constIncrement
	Private boolClientSided

	Private Sub Class_Initialize
		intPointer = 0
		constIncrement = 20
		boolClientSided = False
		ReDim arrControls(50)
	End Sub

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property

	' verifies postback from a form
	Public Property Get IsPostBack 
		IsPostBack = False
		If Not IsNull( objParent ) Then
			IsPostBack = objParent.IsPostBack
		End If
	End Property

	' return the form name
	Public Property Get FormName 
		FormName = False
		If Not IsNull( objParent ) Then
			FormName = objParent.FormName
		End If
	End Property

	' dive thru all the validators and join the boolean results
	Public Property Get IsValid
		Dim boolOK
		boolOK = True
		If intPointer > 0 Then
			Dim intCount
			For intCount = 0 To intPointer
				boolOK = boolOK and arrControls( intPointer ).IsOK
			Next
		End If
		isValid = boolOK
	End Property

	' allows the form to be checked prior to a submit
	Public Property Get IsClientSided
		IsClientSided = boolClientSided
	End Property
	Public Property Let IsClientSided( bool )	
		If bool and LCase( CStr( bool ) ) <> "false" Then
			boolClientSided = True
		Else
			boolClientSided = False
		End If
	End Property
		
	' add validators to the container
	Public Sub Add( ByRef obj )
		If intPointer > UBound( arrControls ) Then
			ReDim Preserve arrControls( intPointer + constIncrement )
		End If
		
		If Not IsNull( obj ) Then
			Set arrControls( intPointer ) = obj
			obj.Parent = Me
			intPointer = intPointer + 1
		End If
	End Sub

	' renders the error message if applicable
	Public Sub ChildRender( boolInvalid, strStyle, strClass, strError )
		If Not Me.IsPostBack or boolInvalid Then
			Exit Sub
		End If
		
		Dim strHTML
		If strStyle <> "" or strClass <> "" Then
			strHTML = "<span%style%class>" & strError & "</span>"
			If strStyle <> "" Then
				strHTML = Replace( strHTML, "%style", " style=""" & strStyle & """ " )
			End If
			If strClass <> "" Then
				strHTML = Replace( strHTML, "%class", " class=""" & strClass & """ " )
			End If
		Else
			strHTML = strError
		End If
		
		Response.Write strHTML
	End Sub
	
End Class

'
'	Encapsulates a "required field" validation pattern
'
Class ValidatorRequired
	Private objParent
	Private objControl
	Private boolOK
	Private strError
	Private strStyle
	Private strClass

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property
	
	Public Property Let Control( ByRef obj ) : Set objControl = obj : End Property
	Public Property Get IsOK : IsOK = CheckValue() : End Property
	
	Public Property Get ErrorMessage : ErrorMessage = strError : End Property
	Public Property Let ErrorMessage( str ) : strError = str : End Property
	
	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style( str ) : strStyle = str : End Property

	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName( str ) : strClass = str : End Property

	Private Sub Class_Initialize
		boolOK = False
		strError = ""
		strStyle = ""
		strClass = ""
	End Sub
	
	' executes the checking
	Private Function CheckValue
		CheckValue = False
		If Not IsNull( objControl ) Then
			Dim strValue
			strValue = objControl.Value
			If Not IsNull( strValue ) and Not IsEmpty( strValue ) and strValue <> "" Then
				CheckValue = True
			End If
		End If
	End Function
	
	' uses the parent render method
	Public Sub Render
		If Not IsNull( objParent ) and Not IsNull( objControl ) Then
			If objParent.IsClientSided Then
				Dim strHTML, strMessage
				strMessage = Replace( strError, "'", "\'", 1, -1 )
				strMessage = Replace( strMessage, "<br>", "\n", 1, -1 )
				strHTML = "<script language=""javascript"">" & vbCRLF & _
				"function __ValidatorRequired_" & objControl.Name & "() " & vbCRLF & _
				"{ " & vbCRLF & _
				"	boolResult = ( document." & objParent.FormName & "." & objControl.Name & ".value != '' );" & vbCRLF & _
				"	if ( ! boolResult ) " & vbCRLF & _
				"	{" & vbCRLF & _
				"		alert( '" & strMessage  & "' );" & vbCRLF & _
				"	}" & vbCRLF & _
				"	return boolResult;" & vbCRLF & _
				"} " & vbCRLF & _
				"__arrValidators.length ++; " & vbCRLF & _
				"__arrValidators[ __arrValidators.length - 1 ] = '__ValidatorRequired_" & objControl.Name & "()';" & vbCRLF & _
				"</script>" & vbCRLF
				Response.Write strHTML
			End If
			Call objParent.ChildRender( CheckValue(), strStyle, strClass, strError )
		End If
	End Sub
End Class

'
'	Makes a comparison between different objects that implements, at least,
'	a "Value" property
'
Class ValidatorCompare
	Private objParent
	Private objControl
	Private objControlValidate
	Private strValueToCompare
	Private boolOK
	Private strOperator
	Private strError
	Private strStyle
	Private strClass

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property
	
	Public Property Let Control( ByRef obj ) : Set objControl = obj : End Property
	Public Property Let ControlToCompare( ByRef obj ) : Set objControlValidate = obj : End Property
	Public Property Get ValueToCompare : ValueToCompare = strValueToCompare : End Property
	Public Property Let ValueToCompare( str ) : strValueToCompare = str : End Property
	Public Property Get IsOK : IsOK = CheckValue() : End Property
	
	Public Property Get ErrorMessage : ErrorMessage = strError : End Property
	Public Property Let ErrorMessage( str ) : strError = str : End Property

	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style( str ) : strStyle = str : End Property

	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName( str ) : strClass = str : End Property
	
	' set which operator to use for comparison
	Public Property Get Operator : Operator = strOperator : End Property
	Public Property Let Operator( str )
		Dim strAll
		strAll = "equal | not equal | greater than | less than | greater than equal | less than equal"
		str = Trim( LCase( str ) )
		If InStr( strAll, str ) Then
			strOperator = str
		End If
	End Property
	
	Private Sub Class_Initialize
		boolOK = False
		strError = ""
		strOperator = "equal"
		strValueToCompare = ""
	End Sub

	Private Function CheckValue
		CheckValue = False
		If Not IsNull( objControl ) and Not IsNull( objControlValidate ) Then
			on error resume next
			Dim strValue1, strValue2
			strValue1 = Trim( objControl.Value )
			If objControlValidate Is Nothing Then
				strValue2 = Trim( strValueToCompare )
			Else
				strValue2 = Trim( objControlValidate.Value )
			End If
			
			Select Case strOperator
			Case "equal"
				CheckValue = ( strValue1 = strValue2 )
			Case "not equal"
				CheckValue = ( strValue1 <> strValue2 )
			Case "greater than"
				CheckValue = ( strValue1 > strValue2 )
			Case "smaller than"
				CheckValue = ( strValue1 < strValue2 )
			Case "greater than equal"
				CheckValue = ( strValue1 >= strValue2 )
			Case "smaller than equal"
				CheckValue = ( strValue1 <= strValue2 )
			End Select
			If Err.number <> 0 Then
				CheckValue = False
			End If
			on error goto 0
		End If
	End Function

	' uses the parent render method
	Public Sub Render
		If Not IsNull( objParent ) Then
			If objParent.IsClientSided Then
				Dim strOperator, strCompare
				strOperator = "=="
				Select Case strOperator
				Case "not equal"
					strOperator = "!="
				Case "greater than"
					strOperator = ">"
				Case "smaller than"
					strOperator = "<"
				Case "greater than equal"
					strOperator = ">="
				Case "smaller than equal"
					strOperator = "<="
				End Select

				If objControlValidate Is Nothing Then
					strCompare = "'" & Replace( Trim( strValueToCompare ), "'", "\'", 1, -1 ) & "'"
				Else
					strCompare = "document." & objParent.FormName & "." & objControlValidate.Name & ".value"
				End If

				Dim strHTML, strMessage
				strMessage = Replace( strError, "'", "\'", 1, -1 )
				strMessage = Replace( strMessage, "<br>", "\n", 1, -1 )

				strHTML = "<script language=""javascript"">" & vbCRLF & _
				"function __ValidatorCompare_" & objControl.Name & "() " & vbCRLF & _
				"{ " & vbCRLF & _
				"	boolResult = ( document." & objParent.FormName & "." & objControl.Name & ".value " & strOperator & " " & strCompare & " );" & vbCRLF & _
				"	if ( ! boolResult ) " & vbCRLF & _
				"	{" & vbCRLF & _
				"		alert( '" & strMessage & "' );" & vbCRLF & _
				"	}" & vbCRLF & _
				"	return boolResult;" & vbCRLF & _
				"} " & vbCRLF & _
				"__arrValidators.length ++; " & vbCRLF & _
				"__arrValidators[ __arrValidators.length - 1 ] = '__ValidatorCompare_" & objControl.Name & "()';" & vbCRLF & _
				"</script>" & vbCRLF
				Response.Write strHTML
			End If
			Call objParent.ChildRender( CheckValue(), strStyle, strClass, strError )
		End If
	End Sub
End Class

'
'	Encapsulates a validation of a value agains a regular expression
'
Class ValidatorRegularExpression
	Private objParent
	Private objControl
	Private boolOK
	Private strError
	Private strRegex
	Private strStyle
	Private strClass

	' the required name of the object
	Public Property Get Parent : Set Parent = objParent : End Property
	Public Property Let Parent( ByRef refObj ) : Set objParent = refObj : End Property
	
	Public Property Let Control( ByRef obj ) : Set objControl = obj : End Property
	Public Property Get IsOK : IsOK = CheckValue() : End Property
	
	Public Property Get ErrorMessage : ErrorMessage = strError : End Property
	Public Property Let ErrorMessage( str ) : strError = str : End Property

	Public Property Get Style : Style = strStyle : End Property
	Public Property Let Style( str ) : strStyle = str : End Property

	Public Property Get ClassName : ClassName = strClass : End Property
	Public Property Let ClassName( str ) : strClass = str : End Property
	
	Public Property Get RegularExpression : RegularExpression = strRegex : End Property
	Public Property Let RegularExpression( str ) : strRegex = str : End Property
	
	Private Sub Class_Initialize
		boolOK = False
		strError = ""
		strRegex = ""
	End Sub

	Private Function CheckValue
		CheckValue = False
		If Not IsNull( objControl ) and strRegex <> "" Then
			' create the regex object (dependes on VBScript 5.1+)
			Dim objRegex, arrResults
			Set objRegex = new RegExp
			objRegex.Pattern = strRegex
			objRegex.IgnoreCase = True
			objRegex.Global = True	
			on error resume next
				Set arrResults = objRegex.Execute( objControl.Value )
				If arrResults.Count > 0 Then
					CheckValue = True
				End If
				If Err.number <> 0 Then
					CheckValue = False
				End If
			on error goto 0
			Set arrResults = Nothing
			Set objReges = Nothing
		End If
	End Function

	' uses the parent render method
	Public Sub Render
		If Not IsNull( objParent ) Then
			If objParent.IsClientSided Then
				Dim strHTML, strMessage
				strMessage = Replace( strError, "'", "\'", 1, -1 )
				strMessage = Replace( strMessage, "<br>", "\n", 1, -1 )
				strHTML = "<script language=""JScript"">" & vbCRLF & _
				"function __ValidatorRequired_" & objControl.Name & "() " & vbCRLF & _
				"{ " & vbCRLF & _
				"	var __regex = new RegExp( '" & strRegex & "' );" & vbCRLF & _
				"	var __regexResult = __regex.exec( document." & objParent.FormName & "." & objControl.Name & ".value );" & vbCRLF & _
				"	boolResult = ( __regexResult != null );" & vbCRLF & _
				"	if ( ! boolResult ) " & vbCRLF & _
				"	{" & vbCRLF & _
				"		alert( '" & strMessage  & "' );" & vbCRLF & _
				"	}" & vbCRLF & _
				"	return boolResult;" & vbCRLF & _
				"} " & vbCRLF & _
				"__arrValidators.length ++; " & vbCRLF & _
				"__arrValidators[ __arrValidators.length - 1 ] = '__ValidatorRequired_" & objControl.Name & "()';" & vbCRLF & _
				"</script>" & vbCRLF
				Response.Write strHTML
			End If

			Call objParent.ChildRender( CheckValue(), strStyle, strClass, strError )
		End If
	End Sub
	
	' return default regex
	Public Function Pattern( strName )
		on error resume next
		strName = LCase( Trim( strName ) )
		If Err.number <> 0 Then
			strName = ""
		End If
		on error goto 0
		Select Case strName
		Case "email"
			Pattern = ".*\@.*\..*"
		Case "integer"
			Pattern = "^\s*[-\+]?\d+\s*$"
		Case "double"
			' must replace decimalchat
			Pattern = "^\s*([-\+])?(\d+)?(\<decimalchar/>(\d+))?\s*$"
		Case "currencyabs"
			' must replace groupchar
			Pattern = "^\s*([-\+])?(((\d+)\<groupchar/>)*)(\d+)\s*$"
		Case "currency"
			' must replace groupchar, decimalchar and digits
			Pattern = "^\s*([-\+])?(((\d+)\<groupchar/>)*)(\d+)(\<decimalchar/>(\d{1,<digits/>}))?\s*$"
		Case "dateymd"
			Pattern = "^\s*((\d{4})|(\d{2}))([-./])(\d{1,2})\4(\d{1,2})\s*$"
		Case "datemdy"
			Pattern = "^\s*(\d{1,2})([-./])(\d{1,2})\2((\d{4})|(\d{2}))\s*$"
		End Select
	End Function
End Class
%>