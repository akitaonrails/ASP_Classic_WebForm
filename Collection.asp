<%
'
'	Implementation of a Collection Class
'
'	- tries to mimic the Visual Basic 6 version
'	- lacks support for keys
'	- has the original Add( value, [key], [beforeIndex], [afterIndex] ) 
'		split up in 3 methods: Add( value ), AddBefore( value, index ) and
'		AddAfter( value, index )
'	- the array of values is implemented as a traditional non-optimized
'		bi-direction linked list so navigation and operations are faster and
'		with small memory footprint
'
Class VBCollection
    Private objIni
    Private objEnd
    Private objCurrent
    Private intCount        ' tracks total number of elements
    Private intCurrent      ' tracks current objCurrent position

    ' initialize pointers
    Private Sub Class_Initialize()
        intCount = 0
        intCurrent = 0
    End Sub
    
    ' returns total number of elements
    Property Get Count(): Count = intCount: End Property

    ' stack a new element at the end of the list
    Public Sub Add(ByRef objVal)
        If intCount = 0 Then
            ' this is the first element
            Set objIni = New VBCollectionNode
            Set objEnd = objIni
            Set objCurrent = objIni
        Else
            ' add to the end of the list
            Dim objTmp
            Set objTmp = New VBCollectionNode
            objTmp.PreviousNode = objEnd
            objEnd.NextNode = objTmp
            Set objEnd = objTmp
            Set objCurrent = objTmp
        End If
        
        objCurrent.Value = objVal
        intCurrent = intCurrent + 1
        intCount = intCount + 1
    End Sub
    
    ' insert a new element in the middle of the list
    Public Sub AddBefore(ByRef objVal, Index)
        If intCount > 0 And Index > 0 And Index <= intCount Then
            Call Item(Index)
            Dim objTmp
            Set objTmp = New VBCollectionNode
            objTmp.Value = objVal
            If objCurrent.PreviousNode Is Nothing Then
                ' set as the first element
                Set objIni = objTmp
                objTmp.NextNode = objCurrent
                objCurrent.PreviousNode = objTmp
            Else
                ' insert between 2 elements
                objCurrent.PreviousNode.NextNode = objTmp
                objTmp.PreviousNode = objCurrent.PreviousNode
                objCurrent.PreviousNode = objTmp
                objTmp.NextNode = objCurrent
            End If
            Set objCurrent = objTmp
            intCount = intCount + 1
        Else
            Call Add(objVal)
        End If
    End Sub
    
    ' insert a new element in the middle of the list
    Public Sub AddAfter(ByRef objVal, Index)
        If intCount > 0 And Index > 0 And Index <= intCount Then
            Call Item(Index)
            Dim objTmp
            Set objTmp = New VBCollectionNode
            If objCurrent.NextNode Is Nothing Then
                ' set as the last element
                objCurrent.NextNode = objTmp
                objTmp.PreviousNode = objCurrent
            Else
                ' insert between 2 elements
                objTmp.PreviousNode = objCurrent
                objCurrent.NextNode.PreviousNode = objTmp
                objTmp.NextNode = objCurrent.NextNode
                objCurrent.NextNode = objTmp
            End If
            Set objCurrent = objTmp
            objCurrent.Value = objVal
            intCurrent = intCurrent + 1
            intCount = intCount + 1
        Else
            Call Add(objVal)
        End If
    End Sub
    
    ' remove an element thru it´s numeric position
    Public Sub Remove(Index)
        If intCount > 0 And Index > 0 And Index <= intCount Then
            Call Item(Index)
            Dim objTmp
            If objCurrent Is objIni Then
                ' if it´s the first element
                Set objTmp = objCurrent
                Set objIni = objIni.NextNode
                objIni.PreviousNode = Nothing
                Set objCurrent = objIni
                Set objTmp = Nothing
            ElseIf objCurrent Is objEnd Then
                ' if it´s the last element
                Set objTmp = objCurrent
                Set objEnd = objCurrent.PreviousNode
                Set objCurrent = objEnd
                Set objTmp = Nothing
                intCurrent = intCurrent - 1
            Else
                ' any element of the middle of the list
                Set objTmp = objCurrent
                objCurrent.NextNode.PreviousNode = objCurrent.PreviousNode
                objCurrent.PreviousNode.NextNode = objCurrent.NextNode
                Set objCurrent = objCurrent.NextNode
                Set objTmp = Nothing
            End If
            intCount = intCount - 1
        End If
    End Sub
    
    ' returns the value of an element
    Public Function Item(Index)
        If intCount > 0 And Index > 0 And Index <= intCount Then
			' determines the shorter start point
			Dim intStartPoint, intEndPoint, intCurrentPoint
			intStartPoint = Index
			intEndPoint = Abs( intCount - Index )
			intCurrentPoint = Abs( Index - intCurrent )

			If intStartPoint < intCurrentPoint and intStartPoint < intEndPoint Then
				Set objCurrent = objIni
				intCurrent = 1
			ElseIf intEndPoint < intCurrentPoint and intEndPoint < intStartPoint Then
				Set objCurrent = objEnd
				intCurrent = intCount
			End If
			
            ' navigates thru the list
            If intCurrent < Index Then
                While intCurrent < Index And Not (objCurrent.NextNode Is Nothing)
                    Set objCurrent = objCurrent.NextNode
                    intCurrent = intCurrent + 1
                Wend
            ElseIf intCurrent > Index Then
                While intCurrent > Index And Not (objCurrent.PreviousNode Is Nothing)
                    Set objCurrent = objCurrent.PreviousNode
                    intCurrent = intCurrent - 1
                Wend
            End If
            If IsObject(objCurrent.Value) Then
                Set Item = objCurrent.Value
            Else
                Item = objCurrent.Value
            End If
        End If
    End Function

End Class

'
'	-- bi-directional linked list node
'
Class VBCollectionNode
    Private objPreviousNode
    Private objNextNode
    Private objValue
    
    ' get/set previous node hook
    Property Get PreviousNode()
        On Error Resume Next
        Set PreviousNode = objPreviousNode
        If Err.Number <> 0 Then
            Set PreviousNode = Nothing
        End If
        On Error GoTo 0
    End Property
    Property Let PreviousNode(ByRef objVal)
        Set objPreviousNode = objVal
    End Property

    ' get/set next node hook
    Property Get NextNode()
        On Error Resume Next
        Set NextNode = objNextNode
        If Err.Number <> 0 Then
            Set NextNode = Nothing
        End If
        On Error GoTo 0
    End Property
    Property Let NextNode(ByRef objVal)
        Set objNextNode = objVal
    End Property

    ' get/set current value
    Property Get Value()
        On Error Resume Next
        If IsObject(objValue) Then
            Set Value = objValue
        Else
            Value = objValue
        End If
        If Err.Number <> 0 Then
            Value = Null
        End If
        On Error GoTo 0
    End Property
    Property Let Value(ByRef obj)
        On Error Resume Next
        If IsObject(obj) Then
            Set objValue = obj
        Else
            objValue = obj
        End If
        If Err.Number <> 0 Then
            objValue = Null
        End If
        On Error GoTo 0
    End Property

End Class
%>