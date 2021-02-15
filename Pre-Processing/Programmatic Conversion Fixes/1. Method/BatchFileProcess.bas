REM  *****  BASIC  *****

Sub Main

End Sub

Sub ShowFiles
  Dim NextFile As String
  Dim AllFiles As String
 
  AllFiles = ""
  NextFile = Dir("C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Pre-Processing\Programmatic Conversion Fixes\2. Testing\Vol 1 page 9-71\", 0) 
  ' C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Pre-Processing\Programmatic Conversion Fixes\2. Testing\Vol 1 page 9-71
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

