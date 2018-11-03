<%
Set SessionFactory = Server.CreateObject( "SessionPlus" )
Set WebFormFactory = Server.CreateObject( "WebForms" )

' initializes an adaptor, a session and binds them together
Set objSession = SessionFactory.CreateInstance( "Session" )
Set objAdaptor = SessionFactory.CreateInstance( "SQLDataAdaptor" )
'Set objSession = New SessionPlus
'Set objAdaptor = New SQLDataAdaptor
objSession.DataAdaptor = objAdaptor

' instantiate all the elements
With WebFormFactory
	Set Page1		= .CreateInstance( "Page" )
	Set Button1		= .CreateInstance( "Button" )
	Set Button2		= .CreateInstance( "Button" )
	Set Button3		= .CreateInstance( "Button" )
	Set Label1		= .CreateInstance( "Label" )
	Set Label2		= .CreateInstance( "Label" )
	Set Label3		= .CreateInstance( "Label" )
	Set TextBox1	= .CreateInstance( "TextBox" )
	Set TextBox2	= .CreateInstance( "TextBox" )
	Set Pass1		= .CreateInstance( "TextBox" )
	Set Pass2		= .CreateInstance( "TextBox" )
	Set DropDown1	= .CreateInstance( "DropDownList" )
	Set DropDown2	= .CreateInstance( "DropDownList" )
	Set List1		= .CreateInstance( "List" )
	Set List2		= .CreateInstance( "List" )
	Set Upload1		= .CreateInstance( "FileUpload" )
	Set Image1		= .CreateInstance( "Image" )
	Set Hidden1		= .CreateInstance( "Hidden" )
	Set ValidatorRequired1	= .CreateInstance( "ValidatorRequired" )
	Set ValidatorCompare1	= .CreateInstance( "ValidatorCompare" )
	Set ValidatorRegex1		= .CreateInstance( "ValidatorRegularExpression" )
End With

' -- configure page
With Page1
	' add the elements to the form	
	.FormName		= "form1"
	.AllowUpload	= True
	.UploadPath		= Replace( Server.MapPath( "teste.asp" ), "teste.asp", "" )
	.MD5Path		= "../wsc/md5.js"
	.Add Button1,	"button"
	.Add Button2,	"button2"
	.Add Button3,	"button3"
	.Add TextBox1,	"textbox1"
	.Add TextBox2,	"textbox2"
	.Add Pass1,		"pass1"
	.Add Pass2,		"pass2"
	.Add Label1,	"label1"
	.Add Label2,	"label2"
	.Add Label3,	"label3"
	.Add DropDown1,	"dropdown1"
	.Add DropDown2,	"dropdown2"
	.Add List1,		"list1"
	.Add List2,		"list2"
	.Add Upload1,	"upload1"
	.Add Image1,	"image1"
	.Add Hidden1,	"hidden1"

	' add validators
	.Validator = WebFormFactory.CreateInstance( "ValidatorContainer" )
	.Validator.IsClientSided = True
	.Validator.Add ValidatorRequired1
	.Validator.Add ValidatorCompare1
	.Validator.Add ValidatorRegex1

	.OnInit = GetRef( "initComponents" )
	.OnLoad = GetRef( "Page1_OnLoad" )
	
	' initialize form
	.Load
	.RenderBegin	' only call render here if there´s no other printings after it
End With

'
'	--- default initComponents
'
Sub initComponents
	' configure session/adaptor
	objAdaptor.ConnectionString = "DSN=PMIDB;UID=pmiuser;PWD=pmiuser"
	strHost = Request.ServerVariables( "HTTP_HOST" )
	If strHost = "akita" or strHost = "200.190.31.214" Then
		objAdaptor.IsOracle = False
	Else
		objAdaptor.IsOracle = True
	End If
	objSession.Load

	' set the button 1
	Button1.Value = "Change Label"
	Button1.OnClick = GetRef( "Button1_OnClick" )

	' set the button 2
	Button2.Value = "Add Item"
	Button2.OnClick = GetRef( "Button2_OnClick" )

	' set the button 3
	Button3.Value = "Expire Session"
	Button3.OnClick = GetRef( "Button3_OnClick" )

	' set the label
	Label1.Value = "teste"
	Label1.Style = "font-family: Arial; font-size: 9pt"
	Label1.Visible = False

	' set the label
	Label2.Value = "teste"
	Label2.Style = "font-family: Arial; font-size: 9pt"
	Label2.Visible = True

	' set the label
	Label3.Value = "Ain´t this great!?"
	Label3.Style = "font-family: Tahoma; font-size: 12pt; font-weight: bolder"
	Label3.Visible = True
	
	' set the input box
	TextBox1.Value = "type in anything"

	' set the input box
	TextBox2.Value = "blabla@bla.com"

	' set the password box
	Pass1.IsPassword = True

	' set the password box
	Pass2.IsPassword = True

	' set the drop down
	DropDown2.AllowMultiple = True
	DropDown2.Rows = 3

	' set the radio list
	List1.Name = "list1"
	List1.OnChange = GetRef( "List1_OnChange" )

	' set the check list
	List2.AllowMultiple = True
	List2.Vertical = False

	' set the image
	Image1.Path = "19.jpg"
	Image1.OnClick = GetRef( "Image1_OnClick" )

	Hidden1.Value = "blabla"

	' configure validators
	With ValidatorRequired1
		.Control = TextBox1
		.ErrorMessage = "Não deixe o campo Teste vazio!"
		.Style = "font-family: Arial; color: red; font-weight: bolder"
	End With

	With ValidatorCompare1
		.Control = Pass1
		.ControlToCompare = Pass2
		.ErrorMessage = "Passwords nao conferem!"
		.Style = "font-family: Verdana; font-weight: bolder; color: blue"
	End With

	With ValidatorRegex1
		.Control = TextBox2
		.ErrorMessage = "Este e-mail é invalido"
		.Style = "font-family: Tahoma; color: red; font-weight: bolder"
		.RegularExpression = ValidatorRegex1.Pattern( "email" )
	End With

End Sub

'
'	--- main event handlers
'

Sub Page1_OnLoad
	If Pass1.Value <> "" Then
		Label1.Value = Pass1.Value & " - " & Pass2.Value
		Label1.Visible = True
		
		Hidden1.Value = Pass2.Value
	End If

	' operation with session
	Response.Write "<BR>"
	If objSession.hasItem( "teste" ) Then
		Response.Write objSession.getItem( "teste" )
	Else
		Response.Write "New Session Initialized"
	End If
	Call objSession.setItem( "teste", "this is a teste - " & Now )
	Response.Write "<BR>"
		
	' will populate de drop down only once, then the serialization takes care of the values
	If Not Page1.IsPostBack Then
		With DropDown1.Items
			.Add "item1", "valor1"
			.Add "item2", "valor2"
			.Add "item3", "valor3"
			.Add "item4", "valor4"
		End With

		With DropDown2.Items
			.Add "item1", "valor1"
			.Add "item2", "valor2"
			.Add "item3", "valor3"
			.Add "item4", "valor4"
		End With
				
		With List1.Items
			.Add "item1", "valor1"
			.Add "item2", "valor2"
			.Add "item3", "valor3"
			.Add "item4", "valor4"
		End With

		With List2.Items
			.Add "item1", "valor1"
			.Add "item2", "valor2"
			.Add "item3", "valor3"
			.Add "item4", "valor4"
		End With
	End If
End Sub

Sub Page1_Finished
	' terminates session
	Response.Write "Página encerrada com sucesso: " & objSession.Commit()
	Set objSession = Nothing
	Set objAdaptor = Nothing
End Sub

Sub Button1_OnClick( sender, arguments )
	TextBox1.Value = "It Works!!"
	Label2.Value = ""
	If IsArray( DropDown2.Value ) Then
		For Each intIndex in DropDown2.Value
			Label2.Value = Label2.Value & " " & DropDown2.GetValue( intIndex )
		Next
	End If
	Label2.Value = Label2.Value & " " & Upload1.Value
End Sub

Sub Button2_OnClick( sender, arguments )
	on error resume next
	DropDown1.Items.Add "key" & List1.Items.Count, "teste"
	DropDown2.Items.Add "key" & List1.Items.Count, "teste"
	List1.Items.Add "key" & List1.Items.Count, "teste"
	List2.Items.Add "key" & List1.Items.Count, "teste"
	If Err.number <> 0 Then
		Response.Write Err.Description
	End If
	on error goto 0
End Sub

Sub Button3_OnClick( sender, arguments )
	objSession.ExpireAll
End Sub

Sub List1_OnChange( sender, arguments )
	If IsArray( List1.Value ) Then
		For Each intIndex in List1.Value
			Label2.Value = Label2.Value & " " & List1.GetValue( intIndex )
		Next
	End If
End Sub

Sub Image1_OnClick( sender, arguments )
	If Image1.Path = "19.jpg" Then
		Image1.Path = "2.jpg"
		Label3.Value = "This is some amazing state of the art CG!"
	Else
		Image1.Path = "19.jpg"
		Label3.Value = "Behold!!"
	End If
End Sub

%>
<%Hidden1.Render%>
<p>
<table>
<tr>
<td>

	Password hashes: <%Label1.Render%> <br>
	<%Label2.Render%> <br> 
	 
	<%Button1.Render%>
	<%Button2.Render%>
	<%Button3.Render%> <br>

	Teste: <%TextBox1.Render%> <br>
	<%ValidatorRequired1.Render%> <br>

	E-mail: <%TextBox2.Render%> <br>
	<%ValidatorRegex1.Render%> <br>

	Password: <%Pass1.Render%> <br>
	Confirm password: <%Pass2.Render%> <br>
	<%ValidatorCompare1.Render%> <br>

	Drop down 1: <%DropDown1.Render%> <br>

	Drop down 2: <%DropDown2.Render%> <br>

	Lista 1: <br> <%List1.Render%> <br>

	Lista 2: <%List2.Render%> <br>

	Upload: <%Upload1.Render%> <br>

</td>
<td>
	<%Image1.Render%> <br>
	<%Label3.Render%>
</td>
</tr>
</table>
</p>

<hr>
<%	
' terminate form
Page1.RenderEnd

' clean up
Call Page1_Finished
Set SessionFactory = Nothing
Set WebFormFactory = Nothing
%>

<p><a href="teste.asp">Reset</a></p>