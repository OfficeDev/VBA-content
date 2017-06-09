---
title: PictureFormat.TransparentBackground Property (Word)
keywords: vbawd10.chm164298860
f1_keywords:
- vbawd10.chm164298860
ms.prod: word
api_name:
- Word.PictureFormat.TransparentBackground
ms.assetid: 8cbc6da7-e3c9-6d42-de48-ae82b3e5ff00
ms.date: 06/08/2017
---


# PictureFormat.TransparentBackground Property (Word)

 **MsoTrue** if the parts of the picture that are defined with a transparent color actually appear transparent. Use the **TransparencyColor** property to set the transparent color. Applies to bitmaps only. Read/write **MsoTriState** .


## Syntax

 _expression_ . **TransparentBackground**

 _expression_ Required. A variable that represents a **[PictureFormat](pictureformat-object-word.md)** object.


## Remarks

If you want to be able to see through the transparent parts of the picture all the way to the objects behind the picture, you must set the  **Visible** property of the picture's **FillFormat** object to **False** . If your picture has a transparent color and the **Visible** property of the picture's **FillFormat** object is set to **True** , the picture's fill will be visible through the transparent color, but objects behind the picture will be obscured.


## Example

This example sets the color returned by the  **RGB** function as the transparent color for shape one in the active document. For the example to work, shape one must be a bitmap.


```vb
blueScreen = RGB(0, 0, 255) 
With ActiveDocument.Shapes(1) 
 With .PictureFormat 
 .TransparentBackground = msoTrue 
 .TransparencyColor = blueScreen 
 End With 
 .Fill.Visible = False 
End With
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-word.md)

