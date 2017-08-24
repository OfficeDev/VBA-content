---
title: TextRange.Move Method (Publisher)
keywords: vbapb10.chm5308422
f1_keywords:
- vbapb10.chm5308422
ms.prod: publisher
api_name:
- Publisher.TextRange.Move
ms.assetid: a51b4153-2ac5-2293-d2a0-d4a3786268d7
ms.date: 06/08/2017
---


# TextRange.Move Method (Publisher)

Collapses the specified range to its start position or end position and then moves the collapsed object by the specified number of units. This method returns a  **Long** that represents the number of units by which the object was actually moved, or it returns 0 (zero) if the move was unsuccessful.


## Syntax

 _expression_. **Move**( **_Unit_**,  **_Size_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Unit|Required| **PbTextUnit**|The unit by which the collapsed range or selection is to be moved.|
|Size|Required| **Long**|The number of units by which the specified range or selection is to be moved. If  **Size** is a positive number, the object is collapsed to its end position and moved forward in the document by the specified number of units. If **Size** is a negative number, the object is collapsed to its start position and moved backward by the specified number of units. You can also control the collapse direction by using the **Collapse** method before using the **Move** method.|

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

This example collapses the specified range and inserts a new sentence at the beginning of the range.


```vb
Sub MoveText() 
 Dim rngText As TextRange 
 Set rngText = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Words(Start:=1, Length:=5) 
 With rngText 
 .Move Unit:=pbTextUnitParagraph, Size:=-1 
 .Text = "This adds new text to the beginning of the range. " 
 End With 
End Sub
```


