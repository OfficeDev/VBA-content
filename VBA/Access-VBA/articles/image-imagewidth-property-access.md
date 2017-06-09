---
title: Image.ImageWidth Property (Access)
keywords: vbaac10.chm10401
f1_keywords:
- vbaac10.chm10401
ms.prod: access
api_name:
- Access.Image.ImageWidth
ms.assetid: 516ebdd4-201d-db7e-de34-7f9ad0bb4955
ms.date: 06/08/2017
---


# Image.ImageWidth Property (Access)

You can use the  **ImageWidth** property to determine the width in twips of a picture in an image control. Read/write **Long**.


## Syntax

 _expression_. **ImageWidth**

 _expression_ A variable that represents an **Image** object.


## Remarks

This property is read-only in all views.

You can use the  **ImageWidth** property together with the **[ImageHeight](image-imageheight-property-access.md)** property to determine the size of a picture in an image control. You could then use this information to change the image control's **Height** and **Width** properties to match the size of the picture displayed.


## Example

The following example prompts the user to enter the name of a bitmap and then assigns that bitmap to the  **Picture** property of the Image1 image control. The **ImageHeight** and **ImageWidth** properties are used to resize the image control to fit the size of the bitmap.


## See also


#### Concepts


[Image Object](image-object-access.md)

