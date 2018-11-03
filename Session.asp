<%
'
'	Replaces the built-in ASP Session object
'	Advantages:
'	- does not consume local resources
'	- allows true Web Farms because the user doesn´t have to get back to the
'		same server all the time, which means that it allows for true load balancing
'		with users accessing different servers but still having a solid session
'	- expiration customizable in the level of the page
'
'	Requirements
'	- MD5.asp (Class MD5)
'	- RDBMS database access and the following table structure
'	- Base64.asp (Class Base64)
'	- a Data Adaptor Class (more on this at the end of the SessionPlus implementation)
'
'	create table tbSessionPlus (
'		userid	varchar2( 40 ) not null,
'		keyname	varchar2( 100 ) not null,
'		value	varchar2( 8000 ) not null,
'		expiration	date	not null
'	)
'
Class SessionPlus
	Private strUserID
	Private objBase64
	Private objItems
	Private dtExpire
	Private objDataAdaptor
	Private boolCommit

	' get/set an MD5 unique identifier (pseudo-unique but very close to the 128 bit GUID)	
	Property Get UserID : UserID = strUserID : End Property
	Property Let UserID( str ) : strUserID = str : End Property

	' get/set expiration date for the current values (must be set before Commit)
	Property Get Expire : Expire = dtExpire : End Property
	Property Let Expire( dtTmp )
		If IsDate( dtTmp ) Then
			dtExpire = dtTmp
		End If
	End Property

	' get/set an external data adaptor that exposes a Save method that receives the
	' current data
	Property Get DataAdaptor : Set DataAdaptor = objDataAdaptor : End Property
	Property Let DataAdaptor( ByRef obj )
		If IsObject( obj ) Then
			Set objDataAdaptor = obj
		End If
	End Property

	' total items
	Property Get Count : Count = objItems.Count : End Property
	
	' constructor
	Private Sub Class_Initialize
		Dim objMD5, strSeed
		
		boolCommit = False
		dtExpire = DateAdd( "h", 1, Now )	' expires in an hour
		Set objDataAdaptor = Nothing
		
		Set objMD5 = new MD5
		Set objBase64 = new Base64
		Set objItems = Server.CreateObject( "Scripting.Dictionary" )
		
		' check if the user already has an ID, otherwise generate a new one based on MD5
		strUserID = Trim( Request.Cookies( "SESSIONKEYPLUS" ) )
		If strUserID = "" Then
			Randomize
			strSeed = Request.ServerVariables( "REMOTE_ADDR" ) & _
				Day( Now ) & Month( Now ) & Year( Now ) & Hour( Now ) & Minute( Now ) & Second( Now ) & Hex( Rnd * 30000 )
			strUserID = objMD5.MakeDigest( strSeed )
		End If
		Set objMD5 = Nothing
		
		' try to write in the cookie header
		on error resume next
		Response.Cookies( "SESSIONKEYPLUS" ) = strUserID
		Response.Cookies( "SESSIONKEYPLUS" ).Path = "/"
		on error goto 0
	End Sub
	
	' destructor - clean up
	Private Sub Class_Terminate
		Set objBase64 = Nothing
		Set objItems = Nothing
		Set objDataAdaptor = Nothing
	End Sub

	' get an item
	Public Function getItem( strKey )
		getItem = objBase64.Decode( objItems( strKey ) )
	End Function
	
	' set/edit a new item
	Public Sub setItem( strKey, strValue )
		If Not objItems.Exists( strKey ) Then
			objItems.Add strKey,  objBase64.Encode( strValue )
		Else
			objItems( strKey ) = objBase64.Encode( strValue )
		End If
		boolCommit = True
	End Sub
	
	' check a key
	Public Function hasItem( strKey )
		hasItem = objItems.Exists( strKey )
	End Function
	
	' expires all items
	Public Sub ExpireAll
		objItems.RemoveAll
	End Sub
	
	' get the item in the given index
	Public Function getItemIndex( intIndex )
		If intIndex >= objItems.Count Then
			Exit Function
		End If
		
		Dim arrTokens
		arrTokens = objItems.Items
		getItemIndex = arrTokens( intIndex )
	End Function
	
	' set the item in the given index
	Public Sub setItemIndex( intIndex, strValue )
		If intIndex >= objItems.Count Then
			Exit Sub
		End If
		
		Dim arrkeys
		arrkeys = objItems.Keys
		objItems( arrKeys( intIndex ) ) = strValue
		boolCommit = True
	End Sub
	
	' get the key in the given index position
	Public Function getKey( intIndex )
		If intIndex >= objItems.Count Then
			Exit Function
		End If
		
		Dim arrkeys
		arrkeys = objItems.Keys
		getKey = arrKeys( intIndex )
	End Function
		
	' load previously saved user data
	Public Function Load
		Load = False
		If Not IsObject( objDataAdaptor ) or objDataAdaptor Is Nothing Then
			Exit Function
		End If

		With objDataAdaptor
			.UserID = strUserID
			.Items = objItems
			
			Load = .Retrieve()
		End With
	End Function
	
	' saves the current user data
	Public Function Commit
		' only run a commit transaction if some data were modified, otherwise do nothing
		If Not boolCommit Then
			Commit = True
			Exit Function
		End If

		Commit = False
		If Not IsObject( objDataAdaptor ) or objDataAdaptor Is Nothing Then
			Exit Function
		End If

		With objDataAdaptor
			.UserID = strUserID
			.Expire = dtExpire
			.Items = objItems
			
			Commit = .Fetch()
		End With
	End Function
End Class

'
'	Implements a default Data Adaptor for SQL
'	Any Class can be an adaptor as far as it implements the following interface:
'
'	Interface IDataAdaptor
'		Property String UserID (get/set)
'		Property Scripting.Dictionary Items (get/set) 
'		Property Date Expire (get/set)
'
'		Public Bool Retrieve
'		Public Bool Fetch
'	End Interface
'
Class SQLDataAdaptor
	Private strUserID
	Private dtExpire
	Private objItems
	Private strTableName
	Private strConnection
	Private objConnection
	Private boolOracle

	' get/set an MD5 unique identifier (pseudo-unique but very close to the 128 bit GUID)	
	Property Get UserID : UserID = strUserID : End Property
	Property Let UserID( str ) : strUserID = str : End Property

	' get/set a Dictionary object from the caller object
	Property Get Items : Set Items = objItems : End Property
	Property Let Items( ByRef obj ) 
		If IsObject( obj ) Then
			Set objItems = obj
		End If
	End Property

	' get/set expiration date for the current values (must be set before Commit)
	Property Get Expire : Expire = dtExpire : End Property
	Property Let Expire( dtTmp )
		If IsDate( dtTmp ) Then
			dtExpire = dtTmp
		End If
	End Property

	' get/set the table name
	Property Get TableName : TableName = strTableName : End Property
	Property Let TableName( str ) : strTableName = str : End Property

	' get/set the connection string
	Property Get ConnectionString : ConnectionString = strConnection : End Property
	Property Let ConnectionString( str ) : strConnection = str : End Property

	' get/set a connection object
	Property Get Connection : Set Connection = objConnection : End Property
	Property Let Connection( ByRef obj ) : Set objConnection = obj : End Property

	' get/set whether it´s Oracle or SQL Server
	Property Get IsOracle : IsOracle = boolOracle : End Property
	Property Let IsOracle( bool ) 
		If bool Then
			boolOracle = True
		Else
			boolOracle = False
		End If
	End Property

	' constructor
	Private Sub Class_Initialize
		boolOracle = True
		strTableName = "tbSessionPlus"
		strConnection = ""
		Set objConnection = Nothing
	End Sub

	' go down to the database and query the saved user data
	Public Function Retrieve
		Retrieve = False
		If strConnection = "" and objConnection Is Nothing THen
			Exit Function
		End If

		Dim strSQL, oConn, oRS
		
		strSQL = "select userid, keyname, value from " & strTableName & " where expiration >= "
		If boolOracle Then
			strSQL = strSQL & "SYSDATE"
		Else
			strSQL = strSQL & "GetDate()"
		End If

		If objConnection Is Nothing Then
			Set oConn = Server.CreateObject( "ADODB.Connection" )
			on error resume next
			oConn.Open strConnection
			If Err.number <> 0 Then
				Exit Function
			End If
			on error goto 0
		Else
			Set oConn = objConnection
		End If

		Set oRS = Server.CreateObject( "ADODB.RecordSet" )
		oRS.Open strSQL, oConn, 3, 3

		' retrieve all values
		on error resume next
		Dim strKey, strValue
		objItems.RemoveAll
		While Not oRS.EOF
			strKey = oRS( "keyname" )
			strValue = oRS( "value" )
			Call objItems.Add( strKey, strValue )
			oRS.MoveNext
		Wend
		oRS.Close

		If Err.number = 0 Then
			Retrieve = True
		End If
		on error goto 0

		Set oRS = Nothing
		Set oConn = Nothing
	End Function
	
	' save the users data
	Public Function Fetch
		Fetch = False

		If strConnection = "" and objConnection Is Nothing and objItems.Count > 0 Then
			Exit Function
		End If

		Dim strSQL, strDateFormat, strDate, strTmp, strKey, oConn, strSQLDelete
		Dim strDay, strMonth, strYear, strHour, strMinute, strSecond, strAMPM, intHour

		strSQLDelete = "delete from " & strTableName & " where userid = '" & strUserID & "'"		
		strSQL = "insert into " & strTableName & " ( userid, keyname, value, expiration ) values ( '%userid', '%key', '%value', %expiration )"
		If boolOracle Then
			strDateFormat = "TO_DATE( ""MM/DD/YYYY HH:MI:SS PM"", ""%date"" )"
		Else
			strDateFormat = "'%date'"
		End If
		
		If Not IsDate( dtExpire ) Then
			dtExpire = Now
		End If

		' formats expiration date
		intHour	= Hour( dtExpire )
		If intHour > 12 Then
			intHour = intHour - 12
			strAMPM = "PM"
		ElseIf intHour = 12 Then
			strAMPM = "PM"
		Else
			strAMPM = "AM"
		End If
		strDay		= FormatInt( Day( dtExpire ), 10 )
		strMonth	= FormatInt( Month( dtExpire ), 10 )
		strYear		= FormatInt( Year( dtExpire ), 100 )
		strHour		= FormatInt( intHour, 10 )
		strMinute	= FormatInt( Minute( dtExpire ), 10 )
		strSecond	= FormatInt( Second( dtExpire ), 10 )
			
		strDate = strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond & " " & strAMPM
		
		strSQL = Replace( strSQL, "%userid", strUserID )
		strSQL = Replace( strSQL, "%expiration", Replace( strDateFormat, "%date", strDate ) )
		
		If objConnection Is Nothing Then
			Set oConn = Server.CreateObject( "ADODB.Connection" )
			on error resume next
			oConn.Open strConnection
			If Err.number <> 0 Then
				Exit Function
			End If
			on error goto 0
		Else
			Set oConn = objConnection
		End If
				
		' make all the inserts in a single transaction
		oConn.BeginTrans
		on error resume next
		' delete all data from the database as it will be re-written 
		oConn.Execute strSQLDelete
		For Each strKey in objItems
			strTmp = Replace( strSQL, "%key", strKey )
			strTmp = Replace( strTmp, "%value", objItems( strKey ) )
			oConn.Execute strTmp
		Next
		If oConn.Errors.Count > 0 Then
			oConn.RollBackTrans
		Else
			oConn.CommitTrans
			Fetch = True
		End If
		on error goto 0
		
		Set oConn = Nothing
	End Function
	
	' Fetch method helper
	Private Function FormatInt( strInt, intRange )
		If strInt < intRange Then
			FormatInt = "0" & strInt
		Else
			FormatInt = strInt
		End If
	End Function
End Class
%>