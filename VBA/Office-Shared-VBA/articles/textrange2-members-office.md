---
title: TextRange2 Members (Office)
ms.prod: office
ms.assetid: 26daffff-b9ef-fd94-f5b7-ed3a09840cb6
ms.date: 06/08/2017
---


# TextRange2 Members (Office)
Represents the text frame in a  **Shape** or **ShapeRange** objects.

Represents the text frame in a  **Shape** or **ShapeRange** objects.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddPeriods](textrange2-addperiods-method-office.md)|Adds period (.) punctuation to the right side of the text contained in TextRange2 object for left-to-right languages and on the left side for right-to-left languages.|
|[ChangeCase](textrange2-changecase-method-office.md)|Changes the case of a  **TextRange2** object to one of the values in the **MsoTextChangeCase** enumeration.|
|[Copy](textrange2-copy-method-office.md)|Copies a  **TextRange2** object.|
|[Cut](textrange2-cut-method-office.md)|Removes a portion or all of the text from a range of text.|
|[Delete](textrange2-delete-method-office.md)|Deletes a  **TextRange2** object.|
|[Find](textrange2-find-method-office.md)|Searches a  **TextRange2** object for a subset of text.|
|[InsertAfter](textrange2-insertafter-method-office.md)|Inserts text to the right of the existing text in the  **TextRange2** object.|
|[InsertBefore](textrange2-insertbefore-method-office.md)|Inserts text to the left of the existing text in the  **TextRange2** object.|
|[InsertChartField](textrange2-insertchartfield-method-office.md)|Inserts a field into the body of a data label in a chart. |
|[InsertSymbol](textrange2-insertsymbol-method-office.md)|Inserts a symbol from the specified font set into the range of text represented by the  **TextRange2** object.|
|[Item](textrange2-item-method-office.md)|Gets the range of text specified by the index number from the  **TextRange2** object.|
|[LtrRun](textrange2-ltrrun-method-office.md)|Returns a  **TextRange2** object that represents the specified subset of left-to-right text runs. A text run consists of a range of characters that share the same font attributes.|
|[Paste](textrange2-paste-method-office.md)|Pastes the contents of the Clipboard into the  **TextRange2** object.|
|[PasteSpecial](textrange2-pastespecial-method-office.md)|Replaces the text range with the contents of the Clipboard in the format specified. If the paste succeeds, this method returns a  **TextRange2** object including the text range that was pasted.|
|[RemovePeriods](textrange2-removeperiods-method-office.md)|Removes all period (.) punctuation from the text in the  **TextRange2** object.|
|[Replace](textrange2-replace-method-office.md)|Finds specific text in a text range, replaces the found text with a specified string, and returns a  **TextRange2** object that represents the first occurrence of the found text. Returns **Nothing** if no match is found.|
|[RotatedBounds](textrange2-rotatedbounds-method-office.md)|Gets the coordinates of the vertices of the text bounding box for the specified text range. Read-only.|
|[RtlRun](textrange2-rtlrun-method-office.md)|Returns a  **TextRange2** object that represents the specified subset of right-to-left text runs. A text run consists of a range of characters that share the same font attributes.|
|[Select](textrange2-select-method-office.md)|Selects the  **TextRange2** object.|
|[TrimText](textrange2-trimtext-method-office.md)|Returns a **TextRange2** object that represents the specified text that has the whitespace removed.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](textrange2-application-property-office.md)|Used without an object qualifier, this property returns an  **Application** object that represents the current instance of the Microsoft Office application. Used with an object qualifier, this property returns an **Application** object that represents the creator of the **TextRange2** object. When used with an OLE Automation object, it returns the object's application. Read-only.|
|[BoundHeight](textrange2-boundheight-property-office.md)|Gets the height, in points, of the text bounding box for the specified text. Read-only.|
|[BoundLeft](textrange2-boundleft-property-office.md)|Gets the left coordinate, in points, of the text bounding box for the specified text. Read-only.|
|[BoundTop](textrange2-boundtop-property-office.md)|Gets the top coordinate, in points, of the text bounding box for the specified text. Read-only.|
|[BoundWidth](textrange2-boundwidth-property-office.md)|Gets the width, in points, of the text bounding box for the specified text. Read-only.|
|[Characters](textrange2-characters-property-office.md)|Read-only.|
|[Count](textrange2-count-property-office.md)|Gets a  **Long** indicating the number of items in the **TextRange2** collection. Read-only.|
|[Creator](textrange2-creator-property-office.md)|Gets a 32-bit integer that indicates the application in which the **TextRange2** object was created. Read-only.|
|[Font](textrange2-font-property-office.md)|Returns a  **Font** object that represents character formatting for the **TextRange2** object. Read-only.|
|[LanguageID](textrange2-languageid-property-office.md)|Gets or sets the  **MsoLanguageID** value of the **TextRange2** object. Read/write.|
|[Length](textrange2-length-property-office.md)|Get a Long that represents the length of a text range. Read-only.|
|[Lines](textrange2-lines-property-office.md)|Returns a TextRange2 object that represents the specified subset of text lines. Read-only.|
|[MathZones](textrange2-mathzones-property-office.md)|Sets the starting point and length of a math zone within a text range. Read-only|
|[ParagraphFormat](textrange2-paragraphformat-property-office.md)|Returns a  **ParagraphFormat** object that represents paragraph formatting for the specified text. Read-only.|
|[Paragraphs](textrange2-paragraphs-property-office.md)|Gets a  **TextRange2** object that represents the specified subset of text paragraphs. Read-only.|
|[Parent](textrange2-parent-property-office.md)|Gets the  **Parent** object for the **TextRange2** object. Read-only.|
|[Runs](textrange2-runs-property-office.md)|Gets a  **TextRange2** object that represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes. Read-only.|
|[Sentences](textrange2-sentences-property-office.md)|Returns a  **TextRange2** object that represents the specified subset of text sentences. Read-only.|
|[Start](textrange2-start-property-office.md)|Gets a  **Long** value indicating the starting point of the specified text range. Read-only.|
|[Text](textrange2-text-property-office.md)|Gets or sets a  **String** value that represents the text in a text range. Read/write.|
|[Words](textrange2-words-property-office.md)|Gets a  **TextRange2** object that represents the specified subset of text words. Read-only.|

