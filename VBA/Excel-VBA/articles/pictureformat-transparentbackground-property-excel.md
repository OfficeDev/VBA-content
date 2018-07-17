---
title: PictureFormat.TransparentBackground Property (Excel)
keywords: vbaxl10.chm113010
f1_keywords:
- vbaxl10.chm113010
ms.prod: excel
api_name:
- Excel.PictureFormat.TransparentBackground
ms.assetid: 9b7cc5b5-610a-821b-cf99-e2af5c4ecf61
ms.date: 06/08/2017
---


# PictureFormat.TransparentBackground Property (Excel)

Use the  **[TransparencyColor](pictureformat-transparencycolor-property-excel.md)** property to set the transparent color. Applies to bitmaps only. Read/write MsoTriState.


## Syntax

 _expression_ . **TransparentBackground**

 _expression_ A variable that represents a **PictureFormat** object.


## Remarks



| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse**|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** The parts of the picture that are the color defined as the transparent color appear transparent.|
If you want to be able to see through the transparent parts of the picture all the way to the objects behind the picture, you must set the  **Visible** property of the picture's **FillFormat** object to **False** . If your picture has a transparent color and the **Visible** property of the picture's **FillFormat** object is set to **True** , the picture's fill will be visible through the transparent color, but objects behind the picture will be obscured.


## Example

This example sets the color that has the RGB value returned by the function RGB(0, 24, 240) as the transparent color for shape one on  `myDocument`. For the example to work, shape one must be a bitmap.


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

