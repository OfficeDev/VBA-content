---
title: PictureFormat.TransparencyColor Property (Excel)
keywords: vbaxl10.chm113009
f1_keywords:
- vbaxl10.chm113009
ms.prod: excel
api_name:
- Excel.PictureFormat.TransparencyColor
ms.assetid: c3a7a247-0cc2-adc8-e13f-a1f4ff728ba0
ms.date: 06/08/2017
---


# PictureFormat.TransparencyColor Property (Excel)

Returns or sets the transparent color for the specified picture as a red-green-blue (RGB) value. For this property to take effect, the  **[TransparentBackground](pictureformat-transparentbackground-property-excel.md)** property must be set to **True** . Applies to bitmaps only. Read/write **Long** .


## Syntax

 _expression_ . **TransparencyColor**

 _expression_ A variable that represents a **PictureFormat** object.


## Remarks

If you want to be able to see through the transparent parts of the picture all the way to the objects behind the picture, you must set the  **Visible** property of the picture's **FillFormat** object to **False** . If your picture has a transparent color and the **Visible** property of the picture's **FillFormat** object is set to **True** , the picture's fill will be visible through the transparent color, but objects behind the picture will be obscured.


## Example

This example sets the color that has the RGB value returned by the function RGB(0, 0, 255) as the transparent color for shape one on  `myDocument`. For the example to work, shape one must be a bitmap.


```vb
blueScreen = RGB(0, 0, 255) 
Set myDocument = Worksheets(1) 
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


[PictureFormat Object](pictureformat-object-excel.md)

