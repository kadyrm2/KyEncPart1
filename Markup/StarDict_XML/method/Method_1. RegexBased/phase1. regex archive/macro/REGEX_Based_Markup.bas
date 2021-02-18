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

Sub Main
	'1. Markup lines and remove par chars
	'2. At each EOL check if the line ends with full stop
	'3. If yes Then convert TERM_START_CANDIDATE to TERM_START Else remove TERM_START_CANDIDATE tag
	'4. restore par chars
	'call show_code_point ' this call sees the module in the same library
	
	call MarkupPageHeader
	call MarkupHeadWord
	call MarkupArticle
	msgbox ("Please manually checkup if there false articles without full stop and then run DefinitionMarkup()!")
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
