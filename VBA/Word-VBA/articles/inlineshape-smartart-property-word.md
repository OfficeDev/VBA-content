---
title: InlineShape.SmartArt Property (Word)
keywords: vbawd10.chm162005148
f1_keywords:
- vbawd10.chm162005148
ms.prod: word
api_name:
- Word.InlineShape.SmartArt
ms.assetid: fbc47fec-04c4-108c-3280-0931f77b4cb5
ms.date: 06/08/2017
---


# InlineShape.SmartArt Property (Word)

Returns a [SmartArt](http://msdn.microsoft.com/library/24332c9b-87c9-7678-9d9f-9e25f2370afc%28Office.15%29.aspx) object that provides a way to work with the SmartArt associated with the specified inline shape. Read-only.


## Syntax

 _expression_ . **SmartArt**

 _expression_ A variable that represents an **[InlineShape](inlineshape-object-word.md)** object.


## Remarks

The  **SmartArt** property provides an entry point for interacting with a SmartArt graphic associated with the inline shape.


## Example

The following code example adds a SmartArt graphic to the active document.


```vb
Dim myDoc As Document 
Dim myInlineShape As InlineShape 
Dim mySmartArt As SmartArt 
 
Set myDoc = ActiveDocument 
Set myInlineShape = myDoc.InlineShapes.AddSmartArt(Application.SmartArtLayouts(2), myDoc.Paragraphs(2).Range) 
Set mySmartArt = myShape.SmartArt 

```


## See also


#### Concepts


[InlineShape Object](inlineshape-object-word.md)

