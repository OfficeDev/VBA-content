---
title: TextRange.MoveEnd Method (Publisher)
keywords: vbapb10.chm5308424
f1_keywords:
- vbapb10.chm5308424
ms.prod: publisher
api_name:
- Publisher.TextRange.MoveEnd
ms.assetid: 4fe27375-34e2-2ecc-33c8-a07230012b13
ms.date: 06/08/2017
---


# TextRange.MoveEnd Method (Publisher)

Moves the ending character position of a range. This method returns a  **Long** that represents the number of units the range or selection actually moved or returns 0 (zero) if the move was unsuccessful.


## Syntax

 _expression_. **MoveEnd**( **_Unit_**,  **_Size_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Unit|Required| **PbTextUnit**|The unit by which the collapsed range or selection is to be moved.|
|Size|Required| **Long**|The number of units to move. If this number is positive, the ending character position is moved forward in the document. If this number is negative, the end is moved backward. If the ending position overtakes the starting position, the range collapses and both character positions move together.|

### Return Value

Long


## Remarks

The Unit parameter can be one of the  **PbTextUnit** constants declared in the Microsoft Publisher type library and shown in the following table.



| **pbTextUnitCell**|
| **pbTextUnitCharacter**|
| **pbTextUnitCharFormat**|
| **pbTextUnitCodePoint**|
| **pbTextUnitColumn**|
| **pbTextUnitLine**|
| **pbTextUnitObject**|
| **pbTextUnitParaFormat**|
| **pbTextUnitParagraph**|
| **pbTextUnitRow**|
| **pbTextUnitScreen**|
| **pbTextUnitSection**|
| **pbTextUnitSentence**|
| **pbTextUnitStory**|
| **pbTextUnitTable**|
| **pbTextUnitWindow**|
| **pbTextUnitWord**|

## Example

This example sets a text range, moves the range's starting and ending character positions, and then formats the font for the range.


```vb
Sub MoveStartEnd() 
 Dim rngText As TextRange 
 
 Set rngText = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Paragraphs(Start:=3, Length:=1) 
 
 With rngText 
 .MoveStart Unit:=pbTextUnitLine, Size:=-2 
 .MoveEnd Unit:=pbTextUnitLine, Size:=1 
 With .Font 
 .Bold = msoTrue 
 .Size = 15 
 End With 
 End With 
 
End Sub
```


