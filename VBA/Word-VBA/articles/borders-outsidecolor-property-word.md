---
title: Borders.OutsideColor Property (Word)
keywords: vbawd10.chm154927137
f1_keywords:
- vbawd10.chm154927137
ms.prod: word
api_name:
- Word.Borders.OutsideColor
ms.assetid: 896fbfd8-d6fa-c89b-591d-3ee4a8a4b0b3
ms.date: 06/08/2017
---


# Borders.OutsideColor Property (Word)

Returns or sets the 24-bit color of the outside borders. .


## Syntax

 _expression_ . **OutsideColor**

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


## Remarks

This property can be any valid  **WdColor** constant or a value returned by Visual Basic's **RGB** function.

If the  **OutsideLineStyle** property is set to either **wdLineStyleNone** or **False** , setting this property has no effect.


## Example

This example adds borders between rows and between columns in the first table of the active document, and then it sets the colors for both the inside and outside borders.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 Set myTable = ActiveDocument.Tables(1) 
 With myTable.Borders 
 .InsideLineStyle = True 
 .InsideColor = wdColorBrightGreen 
 .OutsideColor = wdColorDarkTeal 
 End With 
End If
```

This example adds a dark red, 0.75-point double border around the first paragraph in the active document.




```vb
With ActiveDocument.Paragraphs(1).Borders 
 .OutsideLineStyle = wdLineStyleDouble 
 .OutsideLineWidth = wdLineWidth075pt 
 .OutsideColor = wdColorDarkRed 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

