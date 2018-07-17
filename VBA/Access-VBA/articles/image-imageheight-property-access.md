---
title: Image.ImageHeight Property (Access)
keywords: vbaac10.chm10400
f1_keywords:
- vbaac10.chm10400
ms.prod: access
api_name:
- Access.Image.ImageHeight
ms.assetid: 91d0cc66-8b27-40f0-8112-41410429400c
ms.date: 06/08/2017
---


# Image.ImageHeight Property (Access)

You can use the  **ImageHeight** property to determine the height in twips of the picture in an image control. Read/write **Long**.


## Syntax

 _expression_. **ImageHeight**

 _expression_ A variable that represents an **Image** object.


## Remarks

This property is read-only in all views.

You can use the  **ImageHeight** property together with the **[ImageWidth](image-imagewidth-property-access.md)** property to determine the size of a picture in an image control. You could then use this information to change the image control's **Height** and **Width** properties to match the size of the picture displayed.


## Example

The following example prompts the user to enter the name of a bitmap and then assigns that bitmap to the  **Picture** property of the Image1 image control. The **ImageHeight** and **ImageWidth** properties are used to resize the image control to fit the size of the bitmap.


```vb
Sub GetNewPicture(frm As Form) 
    Dim ctlImage As Control 
    Set ctlImage = frm!Image1 
    ctlImage.Picture = InputBox("Enter path and " _ 
        &; "file name for new bitmap") 
    ctlImage.Height = ctlImage.ImageHeight 
    ctlImage.Width = ctlImage.ImageWidth 
End Sub
```


## See also


#### Concepts


[Image Object](image-object-access.md)

