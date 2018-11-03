<%
' 
' This is the required interface that every element must implement
' in order to be usable by the global Page Class
'	- Parent, Name, Value, IsType properties and required elements
'	- Serialize, SetProperty and Render are required methods
'
Class Element

	' type of this object
	Property Get IsType : IsType = "object" : End Property

	' the required name of the object
	Property Get Parent : End Property
	Property Let Parent( ByRef refObj ) : End Property

	' get/set required name for this element
	Property Get Name : End Property
	Property Let Name ( str ) : End Property

	' get/set required name for this element
	Property Get Value : End Property
	Property Let Value ( str ) : End Property
	
	' set a particular optional event: gets a function reference
	Property Let OnEvent( ByRef funcRef ) : End Property
	
	' handles all events: called by the Page class
	Public Sub EventHandler( sender, arguments )
		If Not IsNull( evtEvent ) Then
			Call evtEvent( sender, arguments )
		End If
	End Sub
	
	' serializes the properties of this element in a XML structure like:
	' <element id="name">
	' <propertyname>value</propertyname>
	' </element>
	Public Function Serialize
	End Function
	
	' called by the parent Page class and must be able to set
	' all of it´s own properties
	Public Sub SetProperty( name, value )
	End Sub
	
	' renders the HTML control in the page
	Public Sub Render
	End Sub
	
	Private Sub Class_Initialize
	End Sub

	Private Sub Class_Terminate
	End Sub
End Class

%>