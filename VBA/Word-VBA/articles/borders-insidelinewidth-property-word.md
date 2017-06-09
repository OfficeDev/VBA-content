---
title: Borders.InsideLineWidth Property (Word)
keywords: vbawd10.chm154927112
f1_keywords:
- vbawd10.chm154927112
ms.prod: word
api_name:
- Word.Borders.InsideLineWidth
ms.assetid: 7feabfc0-cb32-8b56-5f23-3db9c9dadc89
ms.date: 06/08/2017
---


# Borders.InsideLineWidth Property (Word)

Returns or sets the line width of the inside border of an object. .


## Syntax

 _expression_ . **InsideLineWidth**

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


## Remarks

This property returns  **wdUndefined** if the object has inside borders with more than one line width; otherwise, returns **False** or a **WdLineWidth** constant. Can be set to **True** , **False** , or one of the following **WdLineWidth** constants.


## Example

This example adds borders between rows and between columns in the first table in the active document.


```vb
Dim tableTemp As Table 
 
If ActiveDocument.Tables.Count >= 1 Then 
 Set tableTemp = ActiveDocument.Tables(1) 
 tableTemp.Borders.InsideLineStyle = wdLineStyleDot 
 tableTemp.Borders.InsideLineWidth = wdLineWidth050pt 
End If
```

This example adds dotted borders between the first four paragraphs of the active document.




```vb
Dim docActive As Document 
Dim rngTemp As Range 
 
Set docActive = ActiveDocument 
Set rngTemp=docActive.Range( _ 
 Start:=docActive.Paragraphs(1).Range.Start, _ 
 End:=docActive.Paragraphs(4).Range.End) 
 
rngTemp.Borders.InsideLineStyle = wdLineStyleDot 
rngTemp.Borders.InsideLineWidth = wdLineWidth075pt
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

