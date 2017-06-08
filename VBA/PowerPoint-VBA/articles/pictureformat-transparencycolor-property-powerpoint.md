---
title: PictureFormat.TransparencyColor Property (PowerPoint)
keywords: vbapp10.chm551011
f1_keywords:
- vbapp10.chm551011
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat.TransparencyColor
ms.assetid: 122e69f6-a403-92d1-8ef7-087c9396ed6a
ms.date: 06/08/2017
---


# PictureFormat.TransparencyColor Property (PowerPoint)

Returns or sets the transparent color for the specified picture as a red-green-blue (RGB) value. Read/write.


## Syntax

 _expression_. **TransparencyColor**

 _expression_ A variable that represents a **PictureFormat** object.


### Return Value

Long


## Remarks

For this property to take effect, the  **[TransparentBackground](pictureformat-transparentbackground-property-powerpoint.md)** property must be set to **True**.

This property applies to bitmaps only.

If you want to be able to see through the transparent parts of the picture all the way to the objects behind the picture, you must set the  **Visible** property of the picture's **FillFormat** object to **False**. If your picture has a transparent color and the **Visible** property of the picture's **FillFormat** object is set to **True**, the picture's fill will be visible through the transparent color, but objects behind the picture will be obscured.


## Example

This example sets the color that has the RGB value returned by the function RGB(0, 0, 255) as the transparent color for shape one on  `myDocument`. For the example to work, shape one must be a bitmap.


```
blueScreen = RGB(0, 0, 255)

Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1)

    With .PictureFormat

        .TransparentBackground = True

        .TransparencyColor = blueScreen

    End With

    .Fill.Visible = False

End With
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-powerpoint.md)

