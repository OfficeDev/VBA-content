---
title: TextColumn.SpaceAfter Property (Word)
ms.prod: word
api_name:
- Word.TextColumn.SpaceAfter
ms.assetid: 95b77d91-e13a-c6d3-f8c3-069c81b39cb1
ms.date: 06/08/2017
---


# TextColumn.SpaceAfter Property (Word)

Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write  **Single** .


## Syntax

 _expression_ . **SpaceAfter**

 _expression_ Required. A variable that represents a **[TextColumn](textcolumn-object-word.md)** object.


## Example

This example sets the active document to three columns with a 0.5-inch space after the first column. The  **InchesToPoints** method is used to convert inches to points.


```vb
With ActiveDocument.PageSetup.TextColumns 
 .SetCount NumColumns:=3 
 .LineBetween = False 
 .EvenlySpaced = True 
 .Item(1).SpaceAfter = InchesToPoints(0.5) 
End With
```


## See also


#### Concepts


[TextColumn Object](textcolumn-object-word.md)

