---
title: TextRange.Collapse Method (Publisher)
keywords: vbapb10.chm5308420
f1_keywords:
- vbapb10.chm5308420
ms.prod: publisher
api_name:
- Publisher.TextRange.Collapse
ms.assetid: ae177297-bf3b-ce0f-cf3a-29093b115996
ms.date: 06/08/2017
---


# TextRange.Collapse Method (Publisher)

Collapses a range or selection to the starting position or ending position. After a range or selection is collapsed, the starting point and the ending point are equal.


## Syntax

 _expression_. **Collapse**( **_Direction_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Direction|Required| **PbCollapseDirection**|The direction in which to collapse the range or selection.|

## Remarks

If you use  **pbCollapseEnd** to collapse a range that refers to an entire paragraph, the range will be located after the ending paragraph mark (the beginning of the next paragraph). However, you can move the range back one character by using the [MoveEnd](textrange-moveend-method-publisher.md)method after the range is collapsed.

The Direction parameter can be one of the following  **PbCollapseDirection** constants declared in the Microsoft Publisher type library.



| **pbCollapseEnd**|
| **pbCollapseStart**|

## Example

This example inserts text at the beginning of the second paragraph in the first shape on the first page of the active publication. This example assumes that the specified shape is a text frame and not another type of shape.


```vb
Sub CollapseRange() 
 Dim rngText As TextRange 
 Set rngText = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange 
 
 'Collapses range to the end of the range and 
 'enters new text and a new paragraph 
 With rngText 
 .Paragraphs(Start:=1, Length:=1).Collapse Direction:=pbCollapseEnd 
 .Text = "This is a new paragraph." &; vbCrLf 
 End With 
End Sub
```

This example places new text at the end of the first paragraph in the first shape on the first page of the active publication. This example assumes that the specified shape is a text frame and not another type of shape.




```vb
Sub CollapseSelection() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .Paragraphs(Start:=1, Length:=1).Select 
 
 'Collapses selection to end and moves cursor back 
 'one character, then enters new text 
 With Selection.TextRange 
 .Collapse Direction:=pbCollapseEnd 
 .MoveEnd Unit:=pbTextUnitCharacter, Size:=-1 
 .Text = " This is a new test." 
 End With 
End Sub
```


