---
title: WdInformation Enumeration (Word)
ms.prod: word
api_name:
- Word.WdInformation
ms.assetid: b5c46795-9f66-e607-1fb4-3a922b829c40
ms.date: 06/08/2017
---


# WdInformation Enumeration (Word)

Specifies the type of information returned about a specified selection or range.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdActiveEndAdjustedPageNumber**|1|Returns the number of the page that contains the active end of the specified selection or range. If you set a starting page number or make other manual adjustments, returns the adjusted page number (unlike  **wdActiveEndPageNumber** ).|
| **wdActiveEndPageNumber**|3|Returns the number of the page that contains the active end of the specified selection or range, counting from the beginning of the document. Any manual adjustments to page numbering are disregarded (unlike  **wdActiveEndAdjustedPageNumber** ).|
| **wdActiveEndSectionNumber**|2|Returns the number of the section that contains the active end of the specified selection or range.|
| **wdAtEndOfRowMarker**|31|Returns  **True** if the specified selection or range is at the end-of-row mark in a table.|
| **wdCapsLock**|21|Returns  **True** if Caps Lock is in effect.|
| **wdEndOfRangeColumnNumber**|17|Returns the table column number that contains the end of the specified selection or range.|
| **wdEndOfRangeRowNumber**|14|Returns the table row number that contains the end of the specified selection or range.|
| **wdFirstCharacterColumnNumber**|9|Returns the character position of the first character in the specified selection or range. If the selection or range is collapsed, the character number immediately to the right of the range or selection is returned (this is the same as the character column number displayed in the status bar after "Col").|
| **wdFirstCharacterLineNumber**|10|Returns the character position of the first character in the specified selection or range. If the selection or range is collapsed, the character number immediately to the right of the range or selection is returned (this is the same as the character line number displayed in the status bar after "Ln").|
| **wdFrameIsSelected**|11|Returns  **True** if the selection or range is an entire frame or text box.|
| **wdHeaderFooterType**|33|Returns a value that indicates the type of header or footer that contains the specified selection or range. See the table in the remarks section for additional information.|
| **wdHorizontalPositionRelativeToPage**|5|Returns the horizontal position of the specified selection or range; this is the distance from the left edge of the selection or range to the left edge of the page measured in points (1 point = 20 twips, 72 points = 1 inch). If the selection or range isn't within the screen area, returns ? 1.|
| **wdHorizontalPositionRelativeToTextBoundary**|7|Returns the horizontal position of the specified selection or range relative to the left edge of the nearest text boundary enclosing it, in points (1 point = 20 twips, 72 points = 1 inch). If the selection or range isn't within the screen area, returns - 1.|
| **wdInBibliography**|42|Returns  **True** if the specified selection or range is in a bibliography.|
| **wdInCitation**|43|Returns  **True** if the specified selection or range is in a citation.|
| **wdInClipboard**|38|For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.|
| **wdInCommentPane**|26|Returns  **True** if the specified selection or range is in a comment pane.|
| **wdInContentControl**|46|Returns  **True** if the specified selection or range is in a content control.|
| **wdInCoverPage**|41|Returns  **True** if the specified selection or range is in a cover page.|
| **wdInEndnote**|36|Returns  **True** if the specified selection or range is in an endnote area in print layout view or in the endnote pane in normal view.|
| **wdInFieldCode**|44|Returns  **True** if the specified selection or range is in a field code.|
| **wdInFieldResult**|45|Returns  **True** if the specified selection or range is in a field result.|
| **wdInFootnote**|35|Returns  **True** if the specified selection or range is in a footnote area in print layout view or in the footnote pane in normal view.|
| **wdInFootnoteEndnotePane**|25|Returns  **True** if the specified selection or range is in the footnote or endnote pane in normal view or in a footnote or endnote area in print layout view. For more information, see the descriptions of **wdInFootnote** and **wdInEndnote** in the preceding paragraphs.|
| **wdInHeaderFooter**|28|Returns  **True** if the selection or range is in the header or footer pane or in a header or footer in print layout view.|
| **wdInMasterDocument**|34|Returns  **True** if the selection or range is in a master document (that is, a document that contains at least one subdocument).|
| **wdInWordMail**|37|Returns  **True** if the selection or range is in the header or footer pane or in a header or footer in print layout view.|
| **wdMaximumNumberOfColumns**|18|Returns the greatest number of table columns within any row in the selection or range.|
| **wdMaximumNumberOfRows**|15|Returns the greatest number of table rows within the table in the specified selection or range.|
| **wdNumberOfPagesInDocument**|4|Returns the number of pages in the document associated with the selection or range.|
| **wdNumLock**|22|Returns  **True** if Num Lock is in effect.|
| **wdOverType**|23|Returns  **True** if Overtype mode is in effect. The **Overtype** property can be used to change the state of the Overtype mode.|
| **wdReferenceOfType**|32|Returns a value that indicates where the selection is in relation to a footnote, endnote, or comment reference, as shown in the table in the remarks section.|
| **wdRevisionMarking**|24|Returns  **True** if change tracking is in effect.|
| **wdSelectionMode**|20|Returns a value that indicates the current selection mode, as shown in the following table.|
| **wdStartOfRangeColumnNumber**|16|Returns the table column number that contains the beginning of the selection or range.|
| **wdStartOfRangeRowNumber**|13|Returns the table row number that contains the beginning of the selection or range.|
| **wdVerticalPositionRelativeToPage**|6|Returns the vertical position of the selection or range; this is the distance from the top edge of the selection to the top edge of the page measured in points (1 point = 20 twips, 72 points = 1 inch). If the selection isn't visible in the document window, returns ? 1.|
| **wdVerticalPositionRelativeToTextBoundary**|8|Returns the vertical position of the selection or range relative to the top edge of the nearest text boundary enclosing it, in points (1 point = 20 twips, 72 points = 1 inch). This is useful for determining the position of the insertion point within a frame or table cell. If the selection isn't visible, returns ? 1.|
| **wdWithInTable**|12|Returns  **True** if the selection is in a table.|
| **wdZoomPercentage**|19|Returns the current percentage of magnification as set by the  **Percentage** property.|

