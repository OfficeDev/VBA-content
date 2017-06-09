---
title: TabStop.Alignment Property (Publisher)
keywords: vbapb10.chm5636100
f1_keywords:
- vbapb10.chm5636100
ms.prod: publisher
api_name:
- Publisher.TabStop.Alignment
ms.assetid: 59b35d9a-d53b-88cd-952b-6324d1ee7c01
ms.date: 06/08/2017
---


# TabStop.Alignment Property (Publisher)

Returns or sets a  **PbTabAlignmentType** constant that represents the alignment for the specified tab stop. Read/write.


## Syntax

 _expression_. **Alignment**

 _expression_A variable that represents a  **TabStop** object.


## Remarks

The  **Alignment** property value can be one of the **[PbTabAlignmentType](pbtabalignmenttype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example enters a tabbed list and sets the alignment for two custom tab stops. This example assumes that the specified shape is a text frame and not another type of shape and that there are at least two custom tab stops already set.


```vb
Sub CustomDecimalTabStop() 
 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .InsertAfter Newtext:="Pencils" &; vbTab &; _ 
 "Each" &; vbTab &; "1.50" &; vbLf 
 .InsertAfter Newtext:="Pens" &; vbTab &; _ 
 "Each" &; vbTab &; "4.95" &; vbLf 
 .InsertAfter Newtext:="Folders" &; vbTab &; _ 
 "Box" &; vbTab &; "35.28" &; vbLf 
 .InsertAfter Newtext:="Envelopes" &; vbTab &; _ 
 "Case" &; vbTab &; "150.69" &; vbLf 
 With .Paragraphs(Start:=1).ParagraphFormat 
 .Tabs(1).Alignment = pbTabAlignmentCenter 
 .Tabs(2).Alignment = pbTabAlignmentDecimal 
 End With 
 End With 
End Sub
```


