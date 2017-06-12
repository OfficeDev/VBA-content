---
title: PictureFormat.TransparentBackground Property (Publisher)
keywords: vbapb10.chm3604744
f1_keywords:
- vbapb10.chm3604744
ms.prod: publisher
api_name:
- Publisher.PictureFormat.TransparentBackground
ms.assetid: 0a78b579-92bf-36e6-22f6-3ca0a48f5b5a
ms.date: 06/08/2017
---


# PictureFormat.TransparentBackground Property (Publisher)

Indicates whether the parts of the specified picture that are defined as the transparent color appear transparent. Read/write.


## Syntax

 _expression_. **TransparentBackground**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

MsoTriState


## Remarks

The  **TransparentBackground** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| Parts of the picture whose color is the transparency color do not appear transparent.|
| **msoTriStateMixed**|Return value only, indicating a combination of  **msoTrue** and **msoFalse** for the specified objects..|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**| Parts of the picture whose color is the transparency color appear transparent.|
Use the  **[TransparencyColor](pictureformat-transparencycolor-property-publisher.md)** property to set the transparent color.

This property applies only to bitmaps.

If you want to be able to see through the transparent parts of the picture all the way to the objects behind the picture, you must set the  **[Visible](fillformat-visible-property-publisher.md)** property of the picture's **[FillFormat](fillformat-object-publisher.md)** object to **mso False**. If your picture has a transparent color and the  **Visible** property of the picture's **FillFormat** object is set to **msoTrue**, the picture's fill is visible through the transparent color, but objects behind the picture are obscured.


## Example

This example sets the color blue as the transparent color for shape one in the active publication. For the example to work, shape one must be a bitmap.


```vb
With ActiveDocument.Pages(1).Shapes(1) 
 
 With .PictureFormat 
 .TransparentBackground = msoTrue 
 ' RGB(0, 0, 255) is the color blue. 
 .TransparencyColor = RGB(0, 0, 255) 
 End With 
 
 .Fill.Visible = False 
 
End With 

```


