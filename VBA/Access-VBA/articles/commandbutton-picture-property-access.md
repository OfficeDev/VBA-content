---
title: CommandButton.Picture Property (Access)
keywords: vbaac10.chm10451
f1_keywords:
- vbaac10.chm10451
ms.prod: access
api_name:
- Access.CommandButton.Picture
ms.assetid: 1d0d5956-719e-13eb-e6ca-319f8da78754
ms.date: 06/08/2017
---


# CommandButton.Picture Property (Access)

You can use the  **Picture** property to specify a bitmap or other type of graphic to be displayed on the specified control. Read/write **String**.


## Syntax

 _expression_. **Picture**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **Picture** property contains (bitmap) or the path and file name of a bitmap or other type of graphic to be displayed.

The default setting is (none). After the graphic is loaded into the object, the property setting is (bitmap) or the path and file name of the graphic. If you delete (bitmap) or the path and file name of the graphic from the property setting, the picture is deleted from the object and the property setting is again (none).

If the  **PictureType** property is set to Embedded, the graphic is stored with the object.

You can create custom bitmaps by using Microsoft Paintbrush or another application that creates bitmap files. A bitmap file must have a .bmp, .ico, or .dib extension. You can also use graphics files in the .wmf or .emf formats, or any other graphic file type for which you have a graphics filter. Forms, reports, and image controls support all graphics. Command buttons and toggle buttons only support bitmaps.

Buttons can display either a caption or a picture. If you assign both to a button, the picture will be visible, the caption won't. If the picture is deleted, the caption reappears. Microsoft Access displays the picture centered on the button and clipped if the picture is larger than the button.

To create a command button or toggle button with a caption and a picture, you could include the desired caption as part of the bitmap and assign the bitmap to the  **Picture** property of the control.


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

