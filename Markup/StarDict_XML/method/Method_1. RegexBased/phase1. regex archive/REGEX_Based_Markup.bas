REM  *****  BASIC  *****


'REGEX COLLECTION
const PARAGRAPH_END = "$"
const PARAGRAPH_BODY = "(^.*$)"
const PARAGRAPH_START = "^"

const TERM_PERSON  = "^([À-ß???])+[À-ß???][à-ÿ???]"
const TERM_HYPHEN ="^([À-ß«»???])+ –"
const TERM_START_CANDIDATE = "^(([À-ß«»???]){3})+" ' Select Three or more Capital Letters at the begging of the par
const TERM_CANDIDATE = "^[À-ß«»??? ,]+ " 'STRONG MATCH

const PAGE_HEADER_START = "["
const COLUMN_BREAK = "\n\r" 'works in Notepad++

Sub Main
	'1. Markup lines and remove par chars
	'2. At each EOL check if the line ends with full stop
	'3. If yes Then convert TERM_START_CANDIDATE to TERM_START Else remove TERM_START_CANDIDATE tag
	'4. restore par chars
	call show_code_point ' this call sees the module in the same library
End Sub
sub MarkupLine
	findAndReplaceFormattedSimple("$", "<l/>&<l>",16711680 )
	'findAndReplaceFormattedSimple("<br/>", "</l>" & chr(13) & "<l>",RED )
End Sub
sub checkCandidate
	findAndReplaceFormattedSimple("<l>" & chr(13) & "<TSC>", "sss&ddd",16711680 )
End Sub


sub MarkupHeadWord
	findAndReplaceFormattedSimple("(([À-ß ""???]))+[^:upper:]")

end sub

