REM  *****  BASIC  *****
Const MBYES = 6
Const MBABORT = 2
Const MBNO = 7
Sub BulkProcess
	dim files() as String
	dim fileN as Integer
	dim ScrDir as string
	dim DestDir as string
	
	ScrDir = "C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Pre-Processing\Programmatic Conversion Fixes\1. CharRendering_or_FontLack_Distortion\Fixing\SchoolBookCTT to Unicode\input\doc\Vol 1 page 157-250\"
	DestDir = "C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Pre-Processing\Programmatic Conversion Fixes\1. CharRendering_or_FontLack_Distortion\Fixing\SchoolBookCTT to Unicode\output\Vol 1 page 157-250\"
	prefix = "unicoded_"
	
	files = ReadFileNamesToArray(ScrDir)
	msgbox ("Number of files in folder: " + Ubound(files))
	
	for i=0 to Ubound(files)-1 
		'msgbox ("Starting Processing ..." + files(i))
		'Status = msgbox ("Do you want to convert SchoolBookCTT font to Unicode of " + files(i),3+32,"Loop")		
		Status = MBYES
		Select Case Status
			Case MBYES								
				openDoc(ScrDir + files(i))	
				ProcessDoc()
				SaveAsDoc(DestDir + files(i))
				'msgbox ("Document has been saved to " + _Dest)					
				CloseDoc()			
			Case MBABORT, MBNO
				Exit sub
		End Select	
	next
End Sub

Sub OpenDoc (_fileFullName as String)

	Dim Doc As Object
	dim filePath as String
	Dim Url As String
	Dim Dummy() 'An (empty) array of PropertyValues
	 
	filePath = "C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Pre-Processing\Programmatic Conversion Fixes\2. Testing\Vol 1 page 9-71\" 
	fileName = "Vol1 pages 15-17.doc"
	Url = ConvertToUrl(_fileFullName)
	 
	Doc = StarDesktop.loadComponentFromURL(Url, "_blank", 0, Dummy)

End Sub

function ReadFileNamesToArray(_dirPath as String)
  	Dim NextFile As String
  	Dim AllFiles As String
  	dim fileContainer() as String
 	dim i as integer
 	
 	i=0
	AllFiles = ""
	NextFile = Dir(_dirPath, 0) 
	'"C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Pre-Processing\Programmatic Conversion Fixes\2. Testing\test input\"
	While NextFile  <> ""
	  AllFiles = AllFiles & Chr(13) &  NextFile 
	  redim preserve fileContainer(i+1)
	  fileContainer(i) = NextFile
	  'msgbox(fileContainer(i))
	  i=i+1
	  NextFile = Dir	  
	Wend
	
	'fileCount = i-1
	'dim tmpArray() as String
	'_files = fileContainer
	'redim tmpArray(fileCount) ' Erases previous elements
	'msgbox (Ubound(fileContainer))
	'msgbox (Lbound(fileContainer))
	
	ReadFileNamesToArray = fileContainer
	
End Function


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

sub ProcessDoc
	call ConvertFontSBCTT2Unicode
End Sub

sub CloseDoc
	if HasUnoInterfaces(ThisComponent, "com.sun.star.util.XCloseable") then
		thisComponent.close(true)
	else
		ThisComponent.dispose()
	endif
End Sub
sub SaveAsDoc (_Dest as String)
	dim args(0) as new com.sun.star.beans.PropertyValue
	dim sUrl as String
	dim dirPath as String
	dim filePrefix as String
	'dim _fileName as String
	
	dirPath = "C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Pre-Processing\Programmatic Conversion Fixes\2. Testing\Vol 1 page 9-71\"	
	destPath = "C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Pre-Processing\Programmatic Conversion Fixes\2. Testing\Unicoded Vol 1 page 9-71\"
	filePrefix = "unicoded_"
	_fileName = filePrefix + "Vol1 pages 15-17.doc" 
	'sUrl = "file:///C:/Users/Kadyr/Documents/GitHub/Kyrgyzstan-Encyclopedia-Part1/test.doc"
	sUrl = ConvertToUrl(_Dest)
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
