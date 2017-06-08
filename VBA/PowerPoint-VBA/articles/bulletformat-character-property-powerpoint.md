---
title: BulletFormat.Character Property (PowerPoint)
keywords: vbapp10.chm577004
f1_keywords:
- vbapp10.chm577004
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat.Character
ms.assetid: 42480e47-fc3a-d8aa-1368-a76b6776363a
ms.date: 06/08/2017
---


# BulletFormat.Character Property (PowerPoint)

Returns or sets the Unicode character value that is used for bullets in the specified text. Read/write.


## Syntax

 _expression_. **Character**

 _expression_ A variable that represents a **BulletFormat** object.


### Return Value

Long


## Example

This example sets the bullet character for shape two on slide one in the active presentation.


```vb
Set frame2 = ActivePresentation.Slides(1).Shapes(2).TextFrame

With frame2.TextRange.ParagraphFormat.Bullet

    .Character = 8226

    .Visible = True

End With
```


## See also


#### Concepts


[BulletFormat Object](bulletformat-object-powerpoint.md)

