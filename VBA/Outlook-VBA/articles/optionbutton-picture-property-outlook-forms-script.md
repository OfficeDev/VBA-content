---
title: OptionButton.Picture Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 7592d268-506b-0d8b-1224-8069d2ab3a74
ms.date: 06/08/2017
---


# OptionButton.Picture Property (Outlook Forms Script)

Returns a  **String** that specifies the full path name of a bitmap to display on a control. Read-only.


## Syntax

 _expression_. **Picture**

 _expression_A variable that represents an  **OptionButton** object.


## Remarks

You must use the control's property page to assign a bitmap to the  **Picture** property. You cannot use the Visual Basic **LoadPicture** function to assign a bitmap to **Picture**.

To remove a picture that is assigned to a control, click the value of the  **Picture** property in the property page and then press **DELETE**. Pressing  **BACKSPACE** will not remove the picture.

For controls with captions, use the  **[PicturePosition](optionbutton-pictureposition-property-outlook-forms-script.md)** property to specify where to display the picture on the object.

Transparent pictures sometimes have a hazy appearance. If you do not like this appearance, display the picture on a control that supports opaque images.  **[Image](image-object-outlook-forms-script.md)** supports opaque images.


