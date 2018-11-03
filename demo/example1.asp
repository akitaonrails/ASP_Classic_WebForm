<%
' (1) Create the factory
Set WebFormsFactory = Server.CreateObject( "WebForms" )

' (2) Create the Form objects
With WebFormsFactory
	Set Page1 			= .CreateInstance( "Page" )
	Set Button1			= .CreateInstance( "Button" )
	Set TextBox1		= .CreateInstance( "TextBox" )
	Set DropDownList1 	= .CreateInstance( "DropDownList" )
	Set Validator1		= .CreateInstance( "ValidatorRequired" )
End With

' (3) Initialize the Page1 container and add the other elements to it
With Page1
	.FormName = "Form1"
	.MD5Path = "../wsc/md5.js"
	.Add Button1, 		"button1"
	.Add TextBox1,		"textbox1"
	.Add DropDownList1,	"dropdownlist1"

	' set event handlers to be raised on the initialization of the Page
	.OnInit = GetRef( "Page1_OnInit" )
	.OnLoad = GetRef( "Page1_OnLoad" )

	.Validator = WebFormsFactory.CreateInstance( "ValidatorContainer" )
	.Validator.IsClientSided = True
	.Validator.Add Validator1

	' loads the form data and trigger events
	.Load

	' render the header of the Form
	.RenderBegin
End With

' (4) event handler for the Page OnInit
Sub Page1_OnInit
	Button1.Value = "OK"
	Button1.OnClick = GetRef( "Button1_OnClick" )

	TextBox1.Value = ""

	Validator1.Control = TextBox1
	Validator1.ErrorMessage = "Don't leave the field empty"
End Sub

' (5) event handler for the Page OnLoad
Sub Page1_OnLoad
	If Not Page1.IsPostBack Then
		With DropDownList1.Items
			.Add "item1", "value1"
			.Add "item2", "value2"
			.Add "item3", "value3"
			.Add "item4", "value4"
		End With
	End If
End Sub

' (6) event handler for the Button OnClick
Sub Button1_OnClick( sender, arguments )
	Randomize
	Button1.Value = "Hello World"
	DropDownList1.Items.Add "item" & Rnd * Second( Now ), "value" & Rnd * Second( Now )
End Sub

' (7) now render the HTML (try not to render nothing between the
' statements above and this part of the framework
%>
<table>
	<tr><td><%Validator1.Render%></td></tr>
	<tr><td><%TextBox1.Render%></td></tr>
	<tr><td><%DropDownList1.Render%></td></tr>
	<tr><td><%Button1.Render%></td></tr>
</table>
<%
' (8) now render the end of the form and clean up memory
Page1.RenderEnd
Set WebFormsFactory = Nothing
%>