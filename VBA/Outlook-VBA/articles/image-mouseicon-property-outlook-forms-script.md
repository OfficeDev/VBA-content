---
title: Image.MouseIcon Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 1c580dce-1f38-7e40-2ddb-0bb9e6ae0f6c
ms.date: 06/08/2017
---


# Image.MouseIcon Property (Outlook Forms Script)

Returns a  **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.


## Syntax

 _expression_. **MouseIcon**

 _expression_A variable that represents an  **Image** object.


## Remarks

The  **MouseIcon** property is valid when the **[MousePointer](image-mousepointer-property-outlook-forms-script.md)** property is set to 99. The mouse icon of an object is the image that appears when the user moves the mouse across that object.

To assign an image for the mouse pointer, you can either assign a picture to the  **MouseIcon** property or load a picture from a file using the **LoadPicture** function in Visual Basic Scripting Edition.


