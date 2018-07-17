---
title: PictureFormat.TransparentBackground Property (PowerPoint)
keywords: vbapp10.chm551012
f1_keywords:
- vbapp10.chm551012
ms.prod: powerpoint
api_name:
- PowerPoint.PictureFormat.TransparentBackground
ms.assetid: b4a15c64-0568-dcd7-99a2-00295bfe679c
ms.date: 06/08/2017
---


# PictureFormat.TransparentBackground Property (PowerPoint)

Determines whether parts of the picture that are the color defined as the transparent color appear transparent. Applies to bitmaps only. Read/write. 


## Syntax

 _expression_. **TransparentBackground**

 _expression_ A variable that represents a **PictureFormat** object.


### Return Value

MsoTriState


## Remarks

Use the  **[TransparencyColor](pictureformat-transparencycolor-property-powerpoint.md)** property to set the transparent color.

If you want to be able to see through the transparent parts of the picture all the way to the objects behind the picture, you must set the  **Visible** property of the picture's **FillFormat** object to **msoFalse**. If your picture has a transparent color and the **Visible** property of the picture's **FillFormat** object is set to **msoTrue**, the picture's fill will be visible through the transparent color, but objects behind the picture will be obscured.

The value of the  **TransparentBackground** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| Parts of the picture that are the color defined as the transparent color do not appear transparent.|
|**msoTrue**| Parts of the picture that are the color defined as the transparent color appear transparent.|

## Example

This example sets the color that has the RGB value returned by the function RGB(0, 24, 240) as the transparent color for shape one on  `myDocument`. For the example to work, shape one must be a bitmap.


```
blueScreen = RGB(0, 0, 255)

Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1)

    With .PictureFormat

        .TransparentBackground = msoTrue

        .TransparencyColor = blueScreen

    End With

    .Fill.Visible = msoFalse

End With
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-powerpoint.md)

