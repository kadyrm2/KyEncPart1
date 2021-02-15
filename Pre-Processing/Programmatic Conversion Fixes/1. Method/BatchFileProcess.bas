REM  *****  BASIC  *****

Sub Main

End Sub

Sub ShowFiles
  Dim NextFile As String
  Dim AllFiles As String
 
  AllFiles = ""
  NextFile = Dir("C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Pre-Processing\Programmatic Conversion Fixes\2. Testing\Vol 1 page 9-71\", 0) 
  While NextFile  <> ""
    AllFiles = AllFiles & Chr(13) &  NextFile 
    NextFile = Dir
  Wend
 
  MsgBox AllFiles
End Sub
sub CloseDoc
	if HasUnoInterfaces(ThisComponent, "com.sun.star.util.XCloseable") then
		thisComponent.close(true)
	else
		ThisComponent.dispose()
	endif
End Sub
sub SaveAsDoc
	dim args(0) as new com.sun.star.beans.PropertyValue
	dim sUrl as String
	dim dirPath as String
	dirPath = "C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Pre-Processing\Programmatic Conversion Fixes\2. Testing\Vol 1 page 9-71"	
	destPath = "C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\test.doc"
	'sUrl = "file:///C:/Users/Kadyr/Documents/GitHub/Kyrgyzstan-Encyclopedia-Part1/test.doc"
	sUrl = ConvertToUrl(destPath)
	args(0).Name = "Overwrite"
	args(0).Value = False
	
	ThisComponent.storeAsURL(sUrl, args())
	
	
	
End Sub
sub URLtoDIRandBack
	MsgBox ConvertToUrl("C:\doc\test.odt") 
  ' supplies file:///C:/doc/test.odt
	MsgBox ConvertFromUrl("file:///C:/doc/test.odt")    
  '  supplies (under Windows) c:\doc\test.odt
end sub
