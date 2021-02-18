REM  *****  BASIC  *****


'REGEX COLLECTION
const PARAGRAPH_END = "$"
const PARAGRAPH_BODY = "(^.*$)"
const PARAGRAPH_START = "^"

const TERM_PERSON  = "^([À-ß???])+[À-ß???][à-ÿ???]"
const TERM_HYPHEN ="^([À-ß«»???])+ –"
const TERM_START_CANDIDATE = "^(([À-ß«»???]){3})+" ' Select Three or more Capital Letters at the begging of the par
const TERM_CANDIDATE_STRONG = "^[À-ß«»???- ,]+ " 'STRONG MATCH

const PAGE_HEADER_STRONG = "(^[1-9][0-9] [À-ß«»??? ,]+)|([À-ß«»??? ,]+ [1-9][0-9]$)"
const COLUMN_BREAK = "\n\r" 'works in Notepad++


sub MarkupDocsInDir
	dim files() as String
	dim fileN as Integer
	dim ScrDir as string
	dim DestDir as string
	
	ScrDir = "C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Markup\StarDict_XML\input\doc\Vol 1 page 72-116\"
	DestDir = "C:\Users\Kadyr\Documents\GitHub\Kyrgyzstan-Encyclopedia-Part1\Markup\StarDict_XML\output\doc\Vol 1 page 72-116\"
	prefix = "marked_"
	
	files = ReadFileNamesToArray(ScrDir)
	msgbox ("Number of files in folder: " + Ubound(files))
	
	for i=0 to Ubound(files)-1 
		'msgbox ("Starting Processing ..." + files(i))
		'Status = msgbox ("Do you want to convert SchoolBookCTT font to Unicode of " + files(i),3+32,"Loop")		
		Status = MBYES
		Select Case Status
			Case MBYES								
				openDoc(ScrDir + files(i))	
				MarkupDoc()
				SaveAsDoc(DestDir + files(i))
				'msgbox ("Document has been saved to " + _Dest)					
				CloseDoc()			
			Case MBABORT, MBNO
				Exit sub
		End Select	
	next
End Sub

Sub MarkupDoc
	'1. Markup lines and remove par chars
	'2. At each EOL check if the line ends with full stop
	'3. If yes Then convert TERM_START_CANDIDATE to TERM_START Else remove TERM_START_CANDIDATE tag
	'4. restore par chars
	'call show_code_point ' this call sees the module in the same library
	
	call MarkupPageHeader
	call MarkupHeadWord
	call MarkupArticle
	'msgbox ("Please manually checkup if there false articles without full stop and then run DefinitionMarkup()!")
	call MarkupDefinition
	call addInfoTag
	call wrapIntoStarDictTag
	call addXmlTags
End Sub
sub addXmlTags
	rem define variables
	dim document   as object
	dim dispatcher as object
	rem ----------------------------------------------------------------------
	rem get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:GoToStartOfDoc", "", 0, Array())
	
	dim infoTag as String
	
	verTag =	"<?xml version='1.0' encoding='UTF-8'?>" + chr(10)
	bookNameTag =	"<?xml-stylesheet 	type='text/xsl'	version ='2.0'	href='convert_dict_1.xsl'?>"+ chr(10)

	
	infoTag =	verTag + bookNameTag
	'msgbox (infoTag)	
	dim vCursor
	dim vSelection
	dim bIsSelected as Boolean
	
	bIsSelected = true
	vSelection = ThisComponent.getCurrentSelection()
	vCursor = ThisComponent.Text.CreateTextCursorByRange(vSelection(0))
	vCursor.setString(infoTag)
End Sub

sub addInfoTag
	rem define variables
	dim document   as object
	dim dispatcher as object
	rem ----------------------------------------------------------------------
	rem get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:GoToStartOfDoc", "", 0, Array())
	
	dim infoTag as String
	
	verTag =	"<version>2.4.2</version>" + chr(10)
	bookNameTag =	"<bookname>Kyrgyzstan Encyclopedia</bookname>"+ chr(10)
	authorTag =	"<author>content: Kyrgyzstan Encyclopedia; idea: Usen Asanov, Ulan Brimkulov, Kadyr Momunaliev; method: Kadyr Momunaliev</author>"+ chr(10)
	emailTag = 	"<email>unbrim@gmail.com, kadyr.momunaliev@gmail.com</email>"+ chr(10)
	websiteTag =	"<website>www.encyclopedia.edu.kg</website>"+ chr(10)
	descrTag = 	"<description>Copyright: Kyrgyz Encyclopedia Editorial Board;</description>"+ chr(10)
	dateTag = "<date>" + Date() + "</date>"+ chr(10)
	dicttypeTag =	"<dicttype>Textual StarDict Dictionary</dicttype>"+ chr(10)
	
	infoTag =	"<info>" + verTag + bookNameTag + authorTag + emailTag + websiteTag + descrTag + dateTag + dicttypeTag + "</info>" + chr(10)
	'msgbox (infoTag)	
	dim vCursor
	dim vSelection
	dim bIsSelected as Boolean
	
	bIsSelected = true
	vSelection = ThisComponent.getCurrentSelection()
	vCursor = ThisComponent.Text.CreateTextCursorByRange(vSelection(0))
	vCursor.setString(infoTag)
	
End Sub

sub wrapIntoStarDictTag
	rem define variables
	dim document   as object
	dim dispatcher as object
	rem ----------------------------------------------------------------------
	rem get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:GoToStartOfDoc", "", 0, Array())
	
	dim vCursor
	dim vSelection
	dim bIsSelected as Boolean
	
	bIsSelected = true
	vSelection = ThisComponent.getCurrentSelection()
	vCursor = ThisComponent.Text.CreateTextCursorByRange(vSelection(0))
	vCursor.setString(chr(10) + "<stardict xmlns:xi='http://www.w3.org/2003/XInclude'>")
	
	dispatcher.executeDispatch(document, ".uno:GoToEndOfDoc", "", 0, Array())
	vSelection = ThisComponent.getCurrentSelection()
	vCursor = ThisComponent.Text.CreateTextCursorByRange(vSelection(0))
	vCursor.setString(chr(10) + "</stardict>")

End Sub


sub MarkupLine
	findAndReplaceFormattedSimple("$", "<l/>&<l>",16711680 )
	'findAndReplaceFormattedSimple("<br/>", "</l>" & chr(13) & "<l>",RED )
End Sub

sub checkCandidate
	findAndReplaceFormattedSimple("<l>" & chr(13) & "<TSC>", "sss&ddd",16711680 )
End Sub

sub MarkupPageHeader
	findAndReplaceFormattedSimple(PAGE_HEADER_STRONG, "<page_header>&</page_header>",789517 )
	msgbox("Page Headers have been markedup!")
End Sub

sub MarkupHeadWord
	findAndReplaceFormattedSimple(TERM_CANDIDATE_STRONG, "<key>&</key>", 16711680)
	msgbox("Headwords have been markedup!")
end sub

sub MarkupDefinition
	findAndReplaceFormattedSimple("</key>", "&" & chr(10) & "<definition type='h'>" & chr(10)  & "<![CDATA[ ", 16711681)
	findAndReplaceFormattedSimple("</article>", "]]>" & chr(10) & "</definition>" & chr(10) & "&", 16711681)
	msgbox("Definitions have been markedup!")
End Sub

sub MarkupArticle
	findAndReplaceFormattedSimple("<key>", "</article>" & chr(10) & "<article>" & chr(10) & "&", 16711680)
	msgbox("Articles have been markedup!")
End Sub
