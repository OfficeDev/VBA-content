---
title: Borders.OutsideLineStyle Property (Word)
keywords: vbawd10.chm154927111
f1_keywords:
- vbawd10.chm154927111
ms.prod: word
api_name:
- Word.Borders.OutsideLineStyle
ms.assetid: 13e9dfa7-6168-c217-b1bb-eebc81a40fbd
ms.date: 06/08/2017
---


# Borders.OutsideLineStyle Property (Word)

Returns or sets the outside border for the specified object. .


## Syntax

 _expression_ . **OutsideLineStyle**

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


## Remarks

This property returns  **wdUndefined** if more than one kind of border is applied to the specified object; otherwise, returns **False** or a **WdLineStyle** constant. Can be set to **True** , **False** , or a **WdLineStyle** constant.

 **True** sets the line style to the default line style and the line width to the default line width. The default line style and width can be set using the **DefaultBorderLineWidth** and **DefaultBorderLineStyle** properties.

Use either of the following instructions to remove the outside border from the first table in the active document.




```vb
ActiveDocument.Tables(1).Borders.OutsideLineStyle = wdLineStyleNone 
ActiveDocument.Tables(1).Borders.OutsideLineStyle = False
```


## Example

This example adds a double 0.75-point border around the first paragraph in the active document.


```vb
With ActiveDocument.Paragraphs(1).Borders 
 .OutsideLineStyle = wdLineStyleDouble 
 .OutsideLineWidth = wdLineWidth075pt 
End With
```

This example adds a border around the first table in the active document.




```vb
If ActiveDocument.Tables.Count >= 1 Then 
 Set myTable = ActiveDocument.Tables(1) 
 myTable.Borders.OutsideLineStyle = wdLineStyleSingle 
End If
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

