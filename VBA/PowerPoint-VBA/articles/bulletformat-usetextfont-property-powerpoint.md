---
title: BulletFormat.UseTextFont Property (PowerPoint)
keywords: vbapp10.chm577007
f1_keywords:
- vbapp10.chm577007
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat.UseTextFont
ms.assetid: 8d572d8d-bd89-ec94-2484-045306d2730e
ms.date: 06/08/2017
---


# BulletFormat.UseTextFont Property (PowerPoint)

Determines whether the specified bullets are set to the font of the first text character in the paragraph. Read/write.


## Syntax

 _expression_. **UseTextFont**

 _expression_ A variable that represents an **BulletFormat** object.


### Return Value

MsoTriState


## Remarks

You cannot explicitly set this property to  **msoFalse**. Setting the bullet format font (by using the **[Name](font-name-property-powerpoint.md)** property of the **Font** object) sets this property to **msoFalse**. When **UseTextFont** is **msoFalse**, you can set it to **msoTrue** to reset the bullet format to the default font.

The value of the  **UseTextFont** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified bullets are set to a custom font. |
|**msoTrue**| The specified bullets are set to the font of the first text character in the paragraph..|

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

