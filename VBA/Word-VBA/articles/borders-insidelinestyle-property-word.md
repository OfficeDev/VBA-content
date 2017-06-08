---
title: Borders.InsideLineStyle Property (Word)
keywords: vbawd10.chm154927110
f1_keywords:
- vbawd10.chm154927110
ms.prod: word
api_name:
- Word.Borders.InsideLineStyle
ms.assetid: d82862c7-58b2-cb6c-1099-4aaa5bcdf03e
ms.date: 06/08/2017
---


# Borders.InsideLineStyle Property (Word)

Returns or sets the inside border for the specified object. .


## Syntax

 _expression_ . **InsideLineStyle**

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


## Remarks

This property returns  **wdUndefined** if more than one kind of border is applied to the specified object; otherwise, returns **False** or a **WdLineStyle** constant. Can be set to **True** , **False** , or a **WdLineStyle** constant.

 **True** sets the line style to the default line style and the line width to the default line width. The default line style and line width can be set using the **DefaultBorderLineWidth** and **DefaultBorderLineStyle** properties.

Use either of the following instructions to remove the inside border from the first table in the active document.




```vb
ActiveDocument.Tables(1).Borders.InsideLineStyle = wdLineStyleNone 
ActiveDocument.Tables(1).Borders.InsideLineStyle = False
```


## Example

This example adds borders between rows and between columns in the first table of the active document.


```vb
Dim tableTemp As Table 
 
If ActiveDocument.Tables.Count >= 1 Then 
 Set tableTemp = ActiveDocument.Tables(1) 
 tableTemp.Borders.InsideLineStyle = True 
End If
```

This example adds borders between the first four paragraphs in the document.




```vb
Dim docActive As Document 
Dim rngTemp As Range 
 
Set docActive = ActiveDocument 
Set rngTemp = docActive .Range( _ 
 Start:= docActive .Paragraphs(1).Range.Start, _ 
 End:= docActive .Paragraphs(4).Range.End) 
 
With rngTemp.Borders 
 .InsideLineStyle = wdLineStyleSingle 
 .InsideLineWidth = wdLineWidth150pt 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

