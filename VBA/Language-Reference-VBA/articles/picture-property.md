---
title: Picture Property
keywords: fm20.chm2001710
f1_keywords:
- fm20.chm2001710
ms.prod: office
api_name:
- Office.Picture
ms.assetid: ce07e7fb-b123-4ce5-49b5-f21cdedad984
ms.date: 06/08/2017
---


# Picture Property



Specifies the bitmap to display on an object.
 **Syntax**
 _object_. **Picture** = **LoadPicture(**_pathname_**)**
The  **Picture** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _pathname_|Required. The full path to a picture file.|
 **Remarks**
While designing a form, you can use the control's [property page](glossary-vba.md) to assign a bitmap to the **Picture** property. While running a form, you must use the **LoadPicture** function to assign a bitmap to **Picture**.
To remove a picture that is assigned to a control, click the value of the  **Picture** property in the property page and then press DELETE. Pressing BACKSPACE will not remove the picture.

 **Note**  For controls with captions, use the  **PicturePosition** property to specify where to display the picture on the object. Use the **PictureSizeMode** property to determine how the picture fills the object.

Transparent pictures sometimes have a hazy appearance. If you do not like this appearance, display the picture on a control that supports opaque images.  **Image** and **MultiPage** support opaque images.

