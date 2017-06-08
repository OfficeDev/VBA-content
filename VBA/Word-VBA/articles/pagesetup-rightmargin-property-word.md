---
title: PageSetup.RightMargin Property (Word)
keywords: vbawd10.chm158400615
f1_keywords:
- vbawd10.chm158400615
ms.prod: word
api_name:
- Word.PageSetup.RightMargin
ms.assetid: abaabc8b-bb3f-fe68-ca35-d06f74d06791
ms.date: 06/08/2017
---


# PageSetup.RightMargin Property (Word)

Returns or sets the distance (in points) between the right edge of the page and the right boundary of the body text. Read/write  **Single** .


## Syntax

 _expression_ . **RightMargin**

 _expression_ A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Remarks

If the  **[MirrorMargins](pagesetup-mirrormargins-property-word.md)** property is set to **True** , the **RightMargin** property controls the setting for outside margins and the **[LeftMargin](pagesetup-leftmargin-property-word.md)** property controls the setting for inside margins.


## Example

This example displays the right margin setting for the active document. The  **[PointsToInches](global-pointstoinches-method-word.md)** method is used to convert the result to inches.


```vb
With ActiveDocument.PageSetup 
 Msgbox "The right margin is set to " _ 
 &; PointsToInches(.RightMargin) &; " inches." 
End With
```

This example sets the right margin for section two in the selection. The  **[InchesToPoints](application-inchestopoints-method-word.md)** method is used to convert inches to points.




```
Selection.Sections(2).PageSetup.RightMargin = InchesToPoints(1)
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

