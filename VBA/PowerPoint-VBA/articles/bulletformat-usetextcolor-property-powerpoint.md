---
title: BulletFormat.UseTextColor Property (PowerPoint)
keywords: vbapp10.chm577006
f1_keywords:
- vbapp10.chm577006
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat.UseTextColor
ms.assetid: 8242712a-051e-18fa-1b43-93a0ce1cd17b
ms.date: 06/08/2017
---


# BulletFormat.UseTextColor Property (PowerPoint)

Determines whether the specified bullets are set to the color of the first text character in the paragraph. Read/write.


## Syntax

 _expression_. **UseTextColor**

 _expression_ A variable that represents an **BulletFormat** object.


### Return Value

MsoTriState


## Remarks

You cannot explicitly set this property to  **msoFalse**. Setting the bullet format color (using the **[Color](font-color-property-powerpoint.md)** property of the **Font** object) sets this property to **msoFalse**. When **UseTextColor** is **msoFalse**, you can set it to **msoTrue** to reset the bullet format to the default color.

The value of the  **UseTextColor** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified bullets are set to any other color.|
|**msoTrue**| The specified bullets are set to the color of the first text character in the paragraph.|

## Example

This example resets bullets in shape two on slide one in the active presentation to their default character, font, and color.


```vb
With ActivePresentation.Slides(1).Shapes(2) 
    With .TextFrame.TextRange.ParagraphFormat.Bullet 
        .RelativeSize = 1 
        .UseTextColor = msoTrue 
        .UseTextFont = msoTrue 
        .Character = 8226 
    End With 
End With
```


## See also


#### Concepts


[BulletFormat Object](bulletformat-object-powerpoint.md)

