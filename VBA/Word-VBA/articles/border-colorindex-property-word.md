---
title: Border.ColorIndex Property (Word)
keywords: vbawd10.chm154861569
f1_keywords:
- vbawd10.chm154861569
ms.prod: word
api_name:
- Word.Border.ColorIndex
ms.assetid: 649e96e8-b815-2a9e-bebe-d38c847c2e93
ms.date: 06/08/2017
---


# Border.ColorIndex Property (Word)

Returns or sets the color for the specified border or font object. Read/write  **WdColorIndex** .


## Syntax

 _expression_ . **ColorIndex**

 _expression_ Required. A variable that represents a **[Border](border-object-word.md)** object.


## Remarks

The  **wdByAuthor** constant is not valid for border and font objects.


## Example

This example adds a dotted red border around each cell in the first table.


```vb
Dim borderLoop As Border 
 
If ActiveDocument.Tables.Count >= 1 Then 
 For Each borderLoop In ActiveDocument.Tables(1).Borders 
 With borderLoop 
 .ColorIndex = wdRed 
 .LineStyle = wdLineStyleDashDot 
 .LineWidth = wdLineWidth075pt 
 End With 
 Next borderLoop 
End If
```


## See also


#### Concepts


[Border Object](border-object-word.md)

