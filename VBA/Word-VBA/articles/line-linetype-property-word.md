---
title: Line.LineType Property (Word)
keywords: vbawd10.chm49610754
f1_keywords:
- vbawd10.chm49610754
ms.prod: word
api_name:
- Word.Line.LineType
ms.assetid: 06f821b6-c296-8355-b20d-c8f003a29ead
ms.date: 06/08/2017
---


# Line.LineType Property (Word)

Returns a  **wdLineType** constant that indicates whether a line is a text line or a table row.


## Syntax

 _expression_ . **LineType**

 _expression_ Required. A variable that represents a **[Line](line-object-word.md)** object.


## Example

The following example creates a reference to the table if the specified line type is wdTableRow.


```vb
Dim objLine As Line 
Dim objTable As Table 
 
Set objLine = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages(1).Rectangles(1).Lines.Item(1) 
 
If objLine.LineType = wdTableRow Then _ 
 Set objTable = objLine.Range.Tables(1)
```


## See also


#### Concepts


[Line Object](line-object-word.md)

