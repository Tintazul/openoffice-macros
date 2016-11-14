' subCountSubheadings displays the count of all 2nd-level headings for each 1st-level heading
' used to get stats on a document with lots of heading 2’s
' (c) Júlio Reis, 2016 – License: GPL v3. See LICENSE.md or LICENSE.txt

Sub subCountSubheadings()
	Dim oDoc as Object, sText as String, sHeading1 as String
	Dim iHeading1 as Integer, iSubheadings as Integer
	oDoc = ThisComponent
	sText = "Subheading count:"
	sHeading1 = ""
	iHeading1 = 0
	iSubheadings = 0
	' get the whole document outline
	Dim oParagraphs, thisPara
	oParagraphs = oDoc.Text.createEnumeration ' all the paragraphs
	Do While oParagraphs.hasMoreElements
		thisPara = oParagraphs.nextElement
		If thisPara.ImplementationName = "SwXParagraph" then ' is a paragraph
			Select Case thisPara.OutlineLevel
			Case 1: ' is a heading
				If sHeading1 <> "" Then
					' there's been a Heading 1 before
					sText = sText & Chr(13) & iHeading1 & ". " & sHeading1 & ": " & iSubheadings
				End if
				sHeading1 = thisPara.String
				iHeading1 = iHeading1 + 1
				iSubheadings = 0
			Case 2: ' subheading, count it
				iSubheadings = iSubheadings + 1
			End Select
		End if
	Loop
	MsgBox(sText, 48, "Count subheadings")
End Sub