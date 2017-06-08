---
title: Page.Picture Property (Access)
keywords: vbaac10.chm12149
f1_keywords:
- vbaac10.chm12149
ms.prod: access
api_name:
- Access.Page.Picture
ms.assetid: 22487a28-6773-7223-ffcc-59e885790c69
ms.date: 06/08/2017
---


# Page.Picture Property (Access)

You can use the  **Picture** property to specify a bitmap or other type of graphic to be displayed on the specified control. Read/write **String**.


## Syntax

 _expression_. **Picture**

 _expression_ A variable that represents a **Page** object.


## Remarks

The  **Picture** property contains (bitmap) or the path and file name of a bitmap or other type of graphic to be displayed.

The default setting is (none). After the graphic is loaded into the object, the property setting is (bitmap) or the path and file name of the graphic. If you delete (bitmap) or the path and file name of the graphic from the property setting, the picture is deleted from the object and the property setting is again (none).

If the  **PictureType** property is set to Embedded, the graphic is stored with the object.

You can create custom bitmaps by using Microsoft Paintbrush or another application that creates bitmap files. A bitmap file must have a .bmp, .ico, or .dib extension. You can also use graphics files in the .wmf or .emf formats, or any other graphic file type for which you have a graphics filter. Forms, reports, and image controls support all graphics. Command buttons and toggle buttons only support bitmaps.


## See also


#### Concepts


[Page Object](page-object-access.md)

