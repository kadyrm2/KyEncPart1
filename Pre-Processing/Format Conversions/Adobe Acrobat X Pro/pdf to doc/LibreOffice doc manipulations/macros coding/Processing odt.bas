REM  *****  BASIC  *****

Sub Main
	call test1
End Sub

sub test1
	dim doc
	doc = ThisComponent()
	
	
	msgbox (doc.cursor)
	msgbox(chr(0248),, GetDocumentType(doc))
end sub

sub show_code_point
	dim vCursor
	dim vSelection
	dim bIsSelected as Boolean
	
	bIsSelected = true
	vSelection = ThisComponent.getCurrentSelection()
	
	if isNull(vSelection) OR IsEmpty(vSelection) Then
		bIsSelected = False
	elseif vSelection.getCount() = 0 Then
		bIsSelected = False
	end if
	
	if NOT bIsSelected Then 
		Print "Nothing is selected"
		exit sub
	end if
	
	if vSelection.getCount()>1 then
		print "Multiple selection"
	endif
	
	vCursor = ThisComponent.Text.CreateTextCursorByRange(vSelection(0))
	s$ = vCursor.getString()
	
	if Len(s)>0 then
		msgbox ASC(s) & "   " & s,0, "ASCII (unicode) of Selection: "
	else
		pring "Empty string is selected"
	endif
end sub

sub SchoolBoxCTT_to_UTF8
	rem ----------------------------------------------------------------------
	rem define variables
	dim document   as object
	dim dispatcher as object
	rem ----------------------------------------------------------------------
	rem get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:GoToStartOfDoc", "", 0, Array())
	
	call ChangeAllChars
end sub

sub insert_char_after
	rem ----------------------------------------------------------------------
	rem define variables
	dim document   as object
	dim dispatcher as object
	rem ----------------------------------------------------------------------
	rem get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem ----------------------------------------------------------------------
	dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "Text"
	args1(0).Value = chr(65)
	
	dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args1())

end sub


sub select_char
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dim args1(1) as new com.sun.star.beans.PropertyValue
args1(0).Name = "Count"
args1(0).Value = 1
args1(1).Name = "Select"
args1(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args1())

rem ----------------------------------------------------------------------
dim args2(1) as new com.sun.star.beans.PropertyValue
args2(0).Name = "Count"
args2(0).Value = 1
args2(1).Name = "Select"
args2(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args2())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args4(1) as new com.sun.star.beans.PropertyValue
args4(0).Name = "Count"
args4(0).Value = 1
args4(1).Name = "Select"
args4(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args4())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())


end sub
