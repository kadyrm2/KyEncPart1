REM  *****  BASIC  *****

Sub test_findAndReplaceSensitive
	findAndReplaceSensitive(chr$(213), "FOOO", 2)
End Sub

sub test1
	dim doc
	doc = ThisComponent()
	msgbox ("hell", 0, "sdjf")
	
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
		msgbox "Decimal code: " & ASC(s) & " for  " & s & chr(10) & " hex: " & val("AAA") ,0, "ASCII (unicode) of Selection: "
	else
		print "Empty string is selected"
	endif
end sub

sub ConvertRussianCyrillic
	' This function performs convertion of SchoolBookCTT font chars to Unicode Cyrillic group
	' Abbreviations: SchoolBookCTT = SBookCTT
	'
	'
	dim const SBookCTT_start as Integer = 192 , SBookCTT_end as Integer = 255
	dim const UCyrillic_start as Integer = 1040, UCyrillic_end as Integer = 1103
	dim m_find as String, m_replace as String
	dim i as integer, j as integer
	Const MBYES = 6
	Const MBABORT = 2
	Const MBNO = 7
	
	i = 192
	j = 1040
	
	
	for i= SBookCTT_start to SBookCTT_end
		Status = msgbox ("Replacing" & chr(i) & " to " & chr(j),3+32,"Loop")		
		Select Case Status
			Case MBYES
				m_find = chr(i)
				m_replace = chr(j)
				findAndReplaceSensitive(m_find, m_replace, 3)
				j=j+1
			Case MBABORT, MBNO
				End
		End Select
		
	Next
	
End Sub
sub ConvertKyrgyzCyrillic
	' Converts SchoolBookCTT Kyrgyz Cyrillic chars to Unicode Kyrgyz Cyrillic ones
	' Abbreviations: SchoolBookCTT = SB
	' Unicode: Uni
	'
	dim  Barred_O(2,2) as Integer
	dim  Straight_U(2,2) as Integer
	dim  En_With_Tail(2,2) as Integer
	dim ArrayContainer(3) 
	
	dim m_find as String, m_replace as String
	dim i as integer, j as integer
	Const MBYES = 6
	Const MBABORT = 2
	Const MBNO = 7
	
	Barred_O(0,0) = 170  ' Capital SchoolBookCTT Character
	Barred_O(0,1) = 1256 ' Capital Unicode Character
	Barred_O(1,0) = 186  ' Small SchoolBookCTT Character
	Barred_O(1,1) = 1257 ' Small Unicode Character

	Straight_U(0,0) = 175  ' Capital SchoolBookCTT Character, i.e. source
	Straight_U(0,1) = 1198 ' Capital Unicode Character, i.e. destination
	Straight_U(1,0) = 191  ' Small SchoolBookCTT Character, i.e. source
	Straight_U(1,1) = 1199 ' Small Unicode Character, i.e. destination
	
	En_With_Tail(0,0) = 338  ' Capital SchoolBookCTT Character, i.e. source
	En_With_Tail(0,1) = 1225 ' Capital Unicode Character, i.e. destination
	En_With_Tail(1,0) = 339  ' Small SchoolBookCTT Character, i.e. source
	En_With_Tail(1,1) = 1226 ' Small Unicode Character, i.e. destination
	
	ArrayContainer(0) = Barred_O
	ArrayContainer(1) = Straight_U
	ArrayContainer(2) = En_With_Tail
	
	i = 0
	j = 0
	
	for j=0 to 2
		for i= 0 to 1
			m_find = chr(ArrayContainer(j)(i,0))
			m_replace = chr(ArrayContainer(j)(i,1))
			Status = msgbox ("Replacing" & m_find & " to " & m_replace,3+32,"Loop")		
			Select Case Status
				Case MBYES				
					findAndReplaceSensitive(m_find, m_replace, 3)				
				Case MBABORT, MBNO
					End
			End Select			
		Next
	Next
	
End Sub


sub ConvertFontSBCTT2Unicode
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
	
	Call ConvertRussianCyrillic
	Call ConvertKyrgyzCyrillic
end sub

sub findAndReplaceUnsensitive
	rem ----------------------------------------------------------------------
	rem define variables
	dim document   as object
	dim dispatcher as object
	rem ----------------------------------------------------------------------
	rem get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem ----------------------------------------------------------------------
	dim args1(21) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "SearchItem.StyleFamily"
	args1(0).Value = 2
	args1(1).Name = "SearchItem.CellType"
	args1(1).Value = 0
	args1(2).Name = "SearchItem.RowDirection"
	args1(2).Value = true
	args1(3).Name = "SearchItem.AllTables"
	args1(3).Value = false
	args1(4).Name = "SearchItem.SearchFiltered"
	args1(4).Value = false
	args1(5).Name = "SearchItem.Backward"
	args1(5).Value = false
	args1(6).Name = "SearchItem.Pattern"
	args1(6).Value = false
	args1(7).Name = "SearchItem.Content"
	args1(7).Value = false
	args1(8).Name = "SearchItem.AsianOptions"
	args1(8).Value = false
	args1(9).Name = "SearchItem.AlgorithmType"
	args1(9).Value = 1
	args1(10).Name = "SearchItem.SearchFlags"
	args1(10).Value = 65536
	args1(11).Name = "SearchItem.SearchString"
	args1(11).Value = CHR$(213)
	args1(12).Name = "SearchItem.ReplaceString"
	args1(12).Value = "Xelllloooo"
	args1(13).Name = "SearchItem.Locale"
	args1(13).Value = 255
	args1(14).Name = "SearchItem.ChangedChars"
	args1(14).Value = 2
	args1(15).Name = "SearchItem.DeletedChars"
	args1(15).Value = 2
	args1(16).Name = "SearchItem.InsertedChars"
	args1(16).Value = 2
	args1(17).Name = "SearchItem.TransliterateFlags"
	args1(17).Value = 1073742848
	args1(18).Name = "SearchItem.Command"
	args1(18).Value = 0
	args1(19).Name = "SearchItem.SearchFormatted"
	args1(19).Value = false
	args1(20).Name = "SearchItem.AlgorithmType2"
	args1(20).Value = 2
	args1(21).Name = "Quiet"
	args1(21).Value = true
	
	dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())
	
	rem ----------------------------------------------------------------------
	dim args2(21) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "SearchItem.StyleFamily"
	args2(0).Value = 2
	args2(1).Name = "SearchItem.CellType"
	args2(1).Value = 0
	args2(2).Name = "SearchItem.RowDirection"
	args2(2).Value = true
	args2(3).Name = "SearchItem.AllTables"
	args2(3).Value = false
	args2(4).Name = "SearchItem.SearchFiltered"
	args2(4).Value = false
	args2(5).Name = "SearchItem.Backward"
	args2(5).Value = false
	args2(6).Name = "SearchItem.Pattern"
	args2(6).Value = false
	args2(7).Name = "SearchItem.Content"
	args2(7).Value = false
	args2(8).Name = "SearchItem.AsianOptions"
	args2(8).Value = false
	args2(9).Name = "SearchItem.AlgorithmType"
	args2(9).Value = 1
	args2(10).Name = "SearchItem.SearchFlags"
	args2(10).Value = 65536
	args2(11).Name = "SearchItem.SearchString"
	args2(11).Value = CHR$(213)
	args2(12).Name = "SearchItem.ReplaceString"
	args2(12).Value = "Õello"
	args2(13).Name = "SearchItem.Locale"
	args2(13).Value = 255
	args2(14).Name = "SearchItem.ChangedChars"
	args2(14).Value = 2
	args2(15).Name = "SearchItem.DeletedChars"
	args2(15).Value = 2
	args2(16).Name = "SearchItem.InsertedChars"
	args2(16).Value = 2
	args2(17).Name = "SearchItem.TransliterateFlags"
	args2(17).Value = 1073742848
	args2(18).Name = "SearchItem.Command"
	args2(18).Value = 2
	args2(19).Name = "SearchItem.SearchFormatted"
	args2(19).Value = false
	args2(20).Name = "SearchItem.AlgorithmType2"
	args2(20).Value = 2
	args2(21).Name = "Quiet"
	args2(21).Value = true
	
	dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args2())
	
	
end sub

sub findAndReplaceSensitive (_strFind as String, _strReplace as String, _mode as Integer)
	rem ----------------------------------------------------------------------
	' _mode argument meanings:
	' 0 means find and select 1st occurence, 
	' 1 means find and select all occurences,
	' 2 means replace 1st occurence and select next one circling the doc, 
	' 3 means replace all
	
	rem define variables
	dim document   as object
	dim dispatcher as object
	rem ----------------------------------------------------------------------
	rem get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem ----------------------------------------------------------------------
	dim args1(21) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "SearchItem.StyleFamily"
	args1(0).Value = 2
	args1(1).Name = "SearchItem.CellType"
	args1(1).Value = 0
	args1(2).Name = "SearchItem.RowDirection"
	args1(2).Value = true
	args1(3).Name = "SearchItem.AllTables"
	args1(3).Value = false
	args1(4).Name = "SearchItem.SearchFiltered"
	args1(4).Value = false
	args1(5).Name = "SearchItem.Backward"
	args1(5).Value = false
	args1(6).Name = "SearchItem.Pattern"
	args1(6).Value = false
	args1(7).Name = "SearchItem.Content"
	args1(7).Value = false
	args1(8).Name = "SearchItem.AsianOptions"
	args1(8).Value = false
	args1(9).Name = "SearchItem.AlgorithmType"
	args1(9).Value = 1
	args1(10).Name = "SearchItem.SearchFlags"
	args1(10).Value = 65536
	args1(11).Name = "SearchItem.SearchString"
	args1(11).Value = _strFind
	args1(12).Name = "SearchItem.ReplaceString"
	args1(12).Value = _strReplace
	args1(13).Name = "SearchItem.Locale"
	args1(13).Value = 255
	args1(14).Name = "SearchItem.ChangedChars"
	args1(14).Value = 2
	args1(15).Name = "SearchItem.DeletedChars"
	args1(15).Value = 2
	args1(16).Name = "SearchItem.InsertedChars"
	args1(16).Value = 2
	args1(17).Name = "SearchItem.TransliterateFlags"
	' 1024 value: case and diacritics sensitive
	' 1280 value: diacritics sensitive
	args1(17).Value = 1024
	args1(18).Name = "SearchItem.Command"
	args1(18).Value = _mode					
	args1(19).Name = "SearchItem.SearchFormatted"
	args1(19).Value = false
	args1(20).Name = "SearchItem.AlgorithmType2"
	args1(20).Value = 2
	args1(21).Name = "Quiet"
	args1(21).Value = true

	
	dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())
	
	If _mode = 2 then
		dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())
	endif 


end sub


sub findAndReplaceCaseSensitive
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dim args1(21) as new com.sun.star.beans.PropertyValue
args1(0).Name = "SearchItem.StyleFamily"
args1(0).Value = 2
args1(1).Name = "SearchItem.CellType"
args1(1).Value = 0
args1(2).Name = "SearchItem.RowDirection"
args1(2).Value = true
args1(3).Name = "SearchItem.AllTables"
args1(3).Value = false
args1(4).Name = "SearchItem.SearchFiltered"
args1(4).Value = false
args1(5).Name = "SearchItem.Backward"
args1(5).Value = false
args1(6).Name = "SearchItem.Pattern"
args1(6).Value = false
args1(7).Name = "SearchItem.Content"
args1(7).Value = false
args1(8).Name = "SearchItem.AsianOptions"
args1(8).Value = false
args1(9).Name = "SearchItem.AlgorithmType"
args1(9).Value = 1
args1(10).Name = "SearchItem.SearchFlags"
args1(10).Value = 65536
args1(11).Name = "SearchItem.SearchString"
args1(11).Value = "?"
args1(12).Name = "SearchItem.ReplaceString"
args1(12).Value = ""
args1(13).Name = "SearchItem.Locale"
args1(13).Value = 255
args1(14).Name = "SearchItem.ChangedChars"
args1(14).Value = 2
args1(15).Name = "SearchItem.DeletedChars"
args1(15).Value = 2
args1(16).Name = "SearchItem.InsertedChars"
args1(16).Value = 2
args1(17).Name = "SearchItem.TransliterateFlags"
args1(17).Value = 1024
args1(18).Name = "SearchItem.Command"
args1(18).Value = 0
args1(19).Name = "SearchItem.SearchFormatted"
args1(19).Value = false
args1(20).Name = "SearchItem.AlgorithmType2"
args1(20).Value = 2
args1(21).Name = "Quiet"
args1(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())


end sub
