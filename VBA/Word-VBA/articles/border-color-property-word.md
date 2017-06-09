---
title: Border.Color Property (Word)
keywords: vbawd10.chm154861575
f1_keywords:
- vbawd10.chm154861575
ms.prod: word
api_name:
- Word.Border.Color
ms.assetid: ac9d1db8-ab9a-04b9-fa07-491b14bccabd
ms.date: 06/08/2017
---


# Border.Color Property (Word)

Returns or sets the 24-bit color for the specified  **Border** object.


## Syntax

 _expression_ . **Color**

 _expression_ Required. A variable that represents a **[Border](border-object-word.md)** object.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function.


## Example

This example adds a dotted indigo border around each cell in the first table.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 For Each aBorder In ActiveDocument.Tables(1).Borders 
 aBorder.Color = wdColorIndigo 
 aBorder.LineStyle = wdLineStyleDashDot 
 aBorder.LineWidth = wdLineWidth075pt 
 Next aBorder 
End If
```


## See also


#### Concepts


[Border Object](border-object-word.md)

