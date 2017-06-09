---
title: Borders.InsideColorIndex Property (Word)
keywords: vbawd10.chm154927114
f1_keywords:
- vbawd10.chm154927114
ms.prod: word
api_name:
- Word.Borders.InsideColorIndex
ms.assetid: 9c626b1f-1696-4e73-10ef-9cec5d354559
ms.date: 06/08/2017
---


# Borders.InsideColorIndex Property (Word)

Returns or sets the color of the inside borders. Read/write  **WdColorIndex** .


## Syntax

 _expression_ . **InsideColorIndex**

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


## Remarks

If the  **InsideLineStyle** property is set to either **wdLineStyleNone** or **False** , setting this property has no effect.


## Example

This example adds red borders between the first four paragraphs in the active document.


```vb
Dim docActive As Document 
Dim rngTemp As Range 
 
Set docActive = ActiveDocument 
Set rngTemp = docActive.Range( _ 
 Start:=docActive.Paragraphs(1).Range.Start, _ 
 End:=docActive.Paragraphs(4).Range.End) 
 
With rngTemp.Borders 
 .InsideLineStyle = wdLineStyleSingle 
 .InsideLineWidth = wdLineWidth150pt 
 .InsideColorIndex = wdRed 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

