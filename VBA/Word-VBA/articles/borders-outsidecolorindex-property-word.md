---
title: Borders.OutsideColorIndex Property (Word)
keywords: vbawd10.chm154927115
f1_keywords:
- vbawd10.chm154927115
ms.prod: word
api_name:
- Word.Borders.OutsideColorIndex
ms.assetid: e9d0a00d-d2c0-1a97-a484-e6da2ecef60a
ms.date: 06/08/2017
---


# Borders.OutsideColorIndex Property (Word)

Returns or sets the color of the outside borders. Read/write  **WdColorIndex** .


## Syntax

 _expression_ . **OutsideColorIndex**

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


## Remarks

If the  **OutsideLineStyle** property is set to either **wdLineStyleNone** or **False** , setting this property has no effect.


## Example

This example adds borders between rows and between columns in the first table of the active document, and then it sets the colors for both the inside and outside borders.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 Set myTable = ActiveDocument.Tables(1) 
 With myTable.Borders 
 .InsideLineStyle = True 
 .InsideColorIndex = wdBrightGreen 
 .OutsideColorIndex = wdPink 
 End With 
End If
```

This example adds a red, 0.75-point double border around the first paragraph in the active document.




```vb
With ActiveDocument.Paragraphs(1).Borders 
 .OutsideLineStyle = wdLineStyleDouble 
 .OutsideLineWidth = wdLineWidth075pt 
 .OutsideColorIndex = wdRed 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

