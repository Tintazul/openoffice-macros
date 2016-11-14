REM  *****  BASIC  *****
Option Explicit

Dim sDialogTitle as String

' auxiliary sub for fnGetDocumentAnchors
Sub subAddItemToAnchorList (oAnchors() as String, sTheAnchor as String, sType as String)
	Dim sAnchor
	Select Case sType
		Case "Table":
			sAnchor = sTheAnchor + "|table"
		Case "Text Frame":
			sAnchor = sTheAnchor + "|frame"
		Case "Image":
			sAnchor = sTheAnchor + "|graphic"
		Case "Object":
			sAnchor = sTheAnchor + "|ole"
		Case "Section":
			sAnchor = sTheAnchor + "|region"
		Case "Bookmark":
			sAnchor = sTheAnchor
	End Select
	ReDim Preserve oAnchors(UBound(oAnchors)+1) as String
	oAnchors(UBound(oAnchors)) = sAnchor
End Sub

' auxiliary sub for fnGetDocumentAnchors
Sub subAddArrayToAnchorList (oAnchors() as String, oNewAnchors() as String, sType as String)
	Dim i, iStart, iStop
	iStart = LBound(oNewAnchors)
	iStop = UBound(oNewAnchors)
	If iStop < iStart then Exit Sub ' empty array, nothing to do
	For i = iStart to iStop
		subAddItemToAnchorList (oAnchors, oNewAnchors(i), sType)
	Next
End Sub

Function fnOutlineNumberingToString(nOutlineNumbering() as Integer, nOutlineLevel as Integer)
	Dim sLevel as String, i as Integer
	sLevel = ""
	For i = 1 to nOutlineLevel
		sLevel = sLevel & nOutlineNumbering(i) & "."
	Next
	fnOutlineNumberingToString = sLevel
End Function

' get the whole document outline
' parameters should come in empty
' sOutlineNumbering has the numbering, e.g. "1.4"
' sOutlineText has the heading text, e.g. "Introduction"
Sub subGetDocumentOutline(sOutlineNumbering() as String, sOutlineText() as String)
	Dim oDoc as Object
	oDoc = ThisComponent
	
	Dim nCurrentNumbering(10) as Integer
	Dim nCurrentLevel as Integer, nItems as Integer, i as Integer
	Dim sNumbering as String
	Dim oParagraphs as Object, thisPara as Object, thisLevel as Integer
	
	nItems = 0
	nCurrentLevel = 0
		
	oParagraphs = oDoc.Text.createEnumeration ' all the paragraphs
	Do While oParagraphs.hasMoreElements
		thisPara = oParagraphs.nextElement
		If thisPara.ImplementationName = "SwXParagraph" then ' is a paragraph
			thisLevel = thisPara.OutlineLevel
			If thisLevel > 0 Then ' heading found
				nItems = nItems + 1
				If thisLevel < nCurrentLevel Then
					' pad zeroes from thisLevel+1 to nCurrentLevel
					For i = thisLevel+1 to nCurrentLevel
						nCurrentNumbering(i) = 0
					Next
				End if
				nCurrentLevel = thisLevel
				nCurrentNumbering(nCurrentLevel) = nCurrentNumbering(nCurrentLevel) + 1
				sNumbering = fnOutlineNumberingToString(nCurrentNumbering, nCurrentLevel)
				' add to outline list
				ReDim Preserve sOutlineNumbering(UBound(sOutlineNumbering)+1) as String
				sOutlineNumbering(UBound(sOutlineNumbering)) = sNumbering
				ReDim Preserve sOutlineText(UBound(sOutlineText)+1) as String
				sOutlineText(UBound(sOutlineText)) = thisPara.String
			End if
		End if
	Loop
End Sub

' builds a list of all the document's anchors, except outline numbered items
Function fnGetDocumentAnchors()
	Dim oDoc as Object, oAnchors() as String
	oDoc = ThisComponent
	' text tables, text frames, images, objects, bookmarks and text sections
	' outlines have a separate function
	subAddArrayToAnchorList(oAnchors, oDoc.getTextTables().ElementNames, "Table")
	subAddArrayToAnchorList(oAnchors, oDoc.getTextFrames().ElementNames, "Text Frame")
	subAddArrayToAnchorList(oAnchors, oDoc.getGraphicObjects().ElementNames, "Image")
	subAddArrayToAnchorList(oAnchors, oDoc.getEmbeddedObjects().ElementNames, "Object")
	subAddArrayToAnchorList(oAnchors, oDoc.Bookmarks.ElementNames, "Bookmark")
	subAddArrayToAnchorList(oAnchors, oDoc.getTextSections().ElementNames, "Section")
	
	fnGetDocumentAnchors = oAnchors
End Function

' returns the position of the string in the array; -1 if not found
Function fnIndexInArray( theString as String, theArray() as String )
	Dim i as Integer, iStart as Integer, iStop as Integer
	iStart = LBound(theArray)
	iStop = UBound(theArray)
	If iStart<=iStop then
		For i = iStart to iStop
			If theString = theArray(i) then
				fnIndexInArray = i
				Exit function
			End if
		Next
	End if
	fnIndexInArray = -1
End function

' returns the position of the final period in the numbering part of the outline link
' e.g. if theString = "3.4.14.Further considerations", returns 7
' e.g. if theString = "Without further ado", returns 0
Function fnNumberingSplitIndex ( theString as String )
	Dim nIndex
	nIndex = InStr(theString, ".")
	If nIndex < 2 Then
		' either no period in the string or it's the first character;
		' either way, there is no numbering part
		fnNumberingSplitIndex = 0
		Exit Function
	End If
	' precondition: there is at least one period in the string
	If IsNumeric(Left(theString,nIndex-1)) Then
		' the bit to the left of the period can be evaluated as a number
		' try to find another numbering part to the right of the found period,
		' and then add the numbers together
		Dim sNewString as String
		sNewString = Right(theString,Len(theString)-nIndex)
		fnNumberingSplitIndex = fnNumberingSplitIndex(sNewString) + nIndex
		Exit Function
	End If
	' precondition: there are non-numeric characters to the left of the period;
	' since it's not a well-formed number, there's no numbering part
	fnNumberingSplitIndex = 0
End Function

' looks for an outline in the outline numbering 
' links with partial matches work; e.g a link "#1.Hello world|outline" will match
' _both_ "1.Hi|outline" _and_ "2.Hello world|outline"
' asks if user wants to fix a partial link
Function fnIsOutlineInArray ( oFragment as Object, theString as String, _
		sOutlineNumbering() as String, sOutlineText() as String )
	Dim i as Integer, iSplit as Integer, iNumberingIndex as Integer, iTextIndex as Integer
	Dim sNumberingPart as String, sTextPart as String, sProposedURL as String
	Dim bNumberingMatches as Boolean, bTextMatches as Boolean
	Dim sMsg as String, iChoice as Integer
	
	Dim bAutoFixOutlineNumbering as Boolean
	bAutoFixOutlineNumbering = True ' fix outline numbering in link without asking anything
	
	' 1st step: split numbering from string part
	iSplit = fnNumberingSplitIndex(theString)
	sNumberingPart = Left(theString, iSplit)
	sTextPart = Right(theString, Len(theString)-iSplit)
	' 2nd step: find a match for the numbering and text parts
	iNumberingIndex = fnIndexInArray(sNumberingPart, sOutlineNumbering)
	iTextIndex = fnIndexInArray(sTextPart, sOutlineText)
	' 3rd step: examine the matches
	If iTextIndex = -1 Then
		' we don't have a text match
		If iNumberingIndex = -1 Then
			' we don't have *any* match
			sMsg = "Warning: No partial match found for" & Chr(13) _
				& "Link text: " & oFragment.String & Chr(13) _
				& "Link URL: " & oFragment.HyperlinkURL & Chr(13) & Chr(13) _
				& "‘OK’ to continue, ‘Cancel’ to stop processing"
			iChoice = MsgBox(sMsg, 48+1, sDialogTitle)
			fnIsOutlineInArray = (iChoice <> 2)
			Exit function
		Else
			' we have a match on numbering but not on text
			sProposedURL = sNumberingPart & sOutlineText(iNumberingIndex)
			sMsg = "Warning: Match on link outline numbering but not on text" & Chr(13) _
				& "Link text: " & oFragment.String & Chr(13) _
				& theString & " – existing link" & Chr(13) _
				& sProposedURL & " – existing anchor" & Chr(13) & Chr(13) _
				& "‘Yes’ to fix link text, ‘No’ to skip fix and continue checking, ‘Cancel’ to stop processing"
			iChoice = MsgBox( sMsg, 48+3,  sDialogTitle)
			Select Case iChoice
			Case 6:
				' yes
				oFragment.HyperlinkURL = sProposedURL
				fnIsOutlineInArray = True
			Case 7:
				' no
				fnIsOutlineInArray = True
			Case 2:
				' cancel
				fnIsOutlineInArray = False
			End Select
		End If
	Else
		' we have a text match
		If iNumberingIndex = -1 Then
			' we have a match on text but not on numbering
			If bAutoFixOutlineNumbering Then
				' fix outline numbering in link automatically, without asking anything
				oFragment.HyperlinkURL = sProposedURL
			Else
				sProposedURL = sOutLineNumbering(iTextIndex) & sTextPart
				sMsg = "Warning: Match on link text but not on outline numbering" & Chr(13) _
					& "Link text: " & oFragment.String & Chr(13) _
					& theString & " – existing link" & Chr(13) _
					& sProposedURL & " – existing anchor" & Chr(13) & Chr(13) _
					& "‘Yes’ to fix link outline numbering, ‘No’ to skip fix and continue checking, ‘Cancel’ to stop processing"
				iChoice = MsgBox( sMsg, 48+3,  sDialogTitle)
				Select Case iChoice
				Case 6:
					' yes
					oFragment.HyperlinkURL = sProposedURL
					fnIsOutlineInArray = True
				Case 7:
					' no
					fnIsOutlineInArray = True
				Case 2:
					' cancel
					fnIsOutlineInArray = False
				End Select
			End If
		Else
			' double match! all's well!
			fnIsOutlineInArray = True
		End If
	End If
End Function

' auxiliary function to FindBrokenInternalLinks
' inspects any links inside the current document fragment
' used to have an enumeration inside an enumeration, per OOo examples,
' but tables don't have .createEnumeration so this needs a recursive call
Sub subInspectLinks( oAnchors() as String, sOutlineNumbering() as String, sOutlineText() as String, _
		oFragment as Object, iFragments as Integer, iLinks as Integer )
	Dim sMsg, sImplementation, thisPortion
	sImplementation = oFragment.implementationName
	Select Case sImplementation
	
		Case "SwXParagraph":
			' paragraphs can be enumerated
			Dim oParaPortions, sLink, bContinue, notFound
			oParaPortions = oFragment.createEnumeration
			' go through all the text portions in current paragraph
			While oParaPortions.hasMoreElements
				thisPortion = oParaPortions.nextElement
				iFragments = iFragments + 1
				If Left(thisPortion.HyperLinkURL, 1) = "#" then
					' internal link found: get it all except initial # character
					iLinks = iLinks + 1
					sLink = right(thisPortion.HyperLinkURL, Len(thisPortion.HyperLinkURL)-1)
					If Left(sLink,14) = "__RefHeading__" then
						' link inside a table of contents, no need to check
						notFound = False
					Elseif Right(sLink,8) = "|outline" then
						' special case for outline: since we don't know how to get the
						' outline numbering, we have to match the rightmost part of the
						' link only
						bContinue = fnIsOutlineInArray(thisPortion, _
							Left(sLink, Len(sLink)-8), sOutlineNumbering, sOutlineText)
						If not bContinue Then End
						' stop processing if Cancel pressed in fnIsOutlineInArray
					Else
						notFound = (fnIndexInArray(sLink, oAnchors) = -1)
					End if
					If notFound then
						' anchor not found
						' *** DEBUG: code below up to MsgBox
						sMsg = "Fragment #" & iFragments & ", internal link #" & iLinks & Chr(13) _
							& "Bad link: [" & thisPortion.String & "] -> [" _
							& thisPortion.HyperLinkURL & "] " & Chr(13) _
							& "Paragraph:" & Chr(13) & oFragment.String & Chr(13) & Chr(13) _
							& "‘OK’ to continue, ‘Cancel’ to stop processing"
						Dim iChoice as Integer
						iChoice = MsgBox (sMsg, 48+1, sDialogTitle)
						If iChoice = 2 Then End
					End If
				End if
			Wend
			' *** END paragraph
			
		Case "SwXTextTable":
			' text tables have cells
			Dim i, eCells, thisCell, oCellPortions
			eCells = oFragment.getCellNames()
			For i = LBound(eCells) to UBound(eCells)
				thisCell = oFragment.getCellByName(eCells(i))
				oCellPortions = thisCell.createEnumeration
					While oCellPortions.hasMoreElements
						thisPortion = oCellPortions.nextElement
						iFragments = iFragments + 1
						' a table cell may contain a paragraph or another table,
						' so call recursively
						subInspectLinks (oAnchors, sOutlineNumbering, sOutlineText, _
							thisPortion, iFragments, iLinks)
					Wend
				'SwXCell has .String
			Next
			' *** END text table

		Case Else
			sMsg = "Implementation method '" & sImplementation & "' not covered by regular code." _
				& "OK to continue, Cancel to stop"
			If 2 = MsgBox(sMsg, 48+1) then End
			' *** END unknown case

	End Select
End Sub

Sub FindBrokenInternalLinks
	' Find the next broken internal link
	'
	' Pseudocode:
	'
	' * generate document outline - for each element record the numbering and the string
	' * generate link of anchors
	' * loop, searching for internal links
	'     - is the internal link to an outline element?
	'         * Look for a match in the numbering and the text parts
	'             - full match? continue
	'             - partial match? ask if user wants to fix, skip or stop (and act accordingly)
	'             - no match? ask if user wants to continue or stop (and act accordingly)
	'     - is the internal link in the anchor list?
	'         * Yes: continue to next link
	'         * No: (broken link found)
	'             - ask if user wants to continue or stop (and act accordingly)
	' * end loop
	' * display final message
	
	Dim oDoc as Object, oFragments as Object, thisFragment as Object
	Dim iFragments as Integer, iLinks as Integer, sMsg as String
	Dim oAnchors() as String ' list of all anchors in the document
	' list of pairs (numbering, string) for all headings
	Dim sOutlineNumbering() as String, sOutlineText() as String
	sDialogTitle = "Find broken internal link"
	oDoc = ThisComponent

	' get document outline
	subGetDocumentOutline(sOutlineNumbering, sOutlineText)
	' get all document anchors
	oAnchors = fnGetDocumentAnchors()
	
	' find links	
	iFragments = 0 ' fragment counter
	iLinks = 0     ' internal link counter
	oFragments = oDoc.Text.createEnumeration ' has all the paragraphs
	While oFragments.hasMoreElements
		thisFragment = oFragments.nextElement
		iFragments = iFragments + 1
		subInspectLinks (oAnchors, sOutlineNumbering, sOutlineText, _
			thisFragment, iFragments, iLinks)
	Wend
	If iLinks <> 0 then
		sMsg = iLinks & " internal links found"
	Else
		sMsg = "This document has no internal links"
	End if
	MsgBox (sMsg, 64, sDialogTitle)
	
End Sub

' *** END FindBrokenInternalLinks ***
