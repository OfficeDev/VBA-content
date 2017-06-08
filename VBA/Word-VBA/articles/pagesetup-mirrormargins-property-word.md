---
title: PageSetup.MirrorMargins Property (Word)
keywords: vbawd10.chm158400623
f1_keywords:
- vbawd10.chm158400623
ms.prod: word
api_name:
- Word.PageSetup.MirrorMargins
ms.assetid: ae7c53d9-7669-fb22-323f-2ad3984e2dfa
ms.date: 06/08/2017
---


# PageSetup.MirrorMargins Property (Word)

 **True** if the inside and outside margins of facing pages are the same width. Read/write **Long** .


## Syntax

 _expression_ . **MirrorMargins**

 _expression_ An expression that returns a **[PageSetup](pagesetup-object-word.md)** object.


## Remarks

The  **MirrorMargins** property can be **True** , **False** , or **wdUndefined** . If the **MirrorMargins** property is set to **True** , the **[LeftMargin](pagesetup-leftmargin-property-word.md)** property controls the setting for inside margins and the **[RightMargin](pagesetup-rightmargin-property-word.md)** property controls the setting for outside margins.


## Example

This example sets the inside margins of the active document to 1 inch (72 points) and the outside margins to 0.5 inch. The  **[InchesToPoints](application-inchestopoints-method-word.md)** method is used to convert inches to points.


```vb
With ActiveDocument.PageSetup 
 .MirrorMargins = True 
 .LeftMargin = 72 
 .RightMargin = InchesToPoints(0.5) 
End With
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

