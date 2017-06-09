---
title: BulletFormat.Picture Method (PowerPoint)
keywords: vbapp10.chm577012
f1_keywords:
- vbapp10.chm577012
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat.Picture
ms.assetid: a38872c0-b754-bf30-3bd5-9050c5edf8f4
ms.date: 06/08/2017
---


# BulletFormat.Picture Method (PowerPoint)

Sets the graphics file to be used for bullets in a bulleted list when the  **[Type](bulletformat-type-property-powerpoint.md)** property of the **BulletFormat** object is set to **ppBulletPicture**.


## Syntax

 _expression_. **Picture**

 _expression_ A variable that represents a **BulletFormat** object.


## Remarks

Valid graphics files include files with the following extensions: .bmp, .cdr, .cgm, .drw, .dxf, .emf, .eps, .gif, .jpg, .jpeg, .pcd, .pct, .pcx, .pict, .png, .tga, .tiff, .wmf, and .wpg.


## Example

This example sets the bullets in the text box specified by shape two on slide one to a bitmap picture of a blue rivet.


```vb
With ActivePresentation.Slides(1).Shapes(2).TextFrame

    With .TextRange.ParagraphFormat.Bullet

        .Type = ppBulletPicture

        .Picture ("C:\Windows\Blue Rivets.bmp")

    End With

End With
```


## See also


#### Concepts


[BulletFormat Object](bulletformat-object-powerpoint.md)

