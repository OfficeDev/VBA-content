---
title: Frame.MouseIcon Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: fa8f125f-2560-b7ae-0498-a19593f309ed
ms.date: 06/08/2017
---


# Frame.MouseIcon Property (Outlook Forms Script)

Returns a  **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.


## Syntax

 _expression_. **MouseIcon**

 _expression_A variable that represents a  **Frame** object.


## Remarks

The  **MouseIcon** property is valid when the **[MousePointer](frame-mousepointer-property-outlook-forms-script.md)** property is set to 99. The mouse icon of an object is the image that appears when the user moves the mouse across that object.

To assign an image for the mouse pointer, you can either assign a picture to the  **MouseIcon** property or load a picture from a file using the **LoadPicture** function in Visual Basic Scripting Edition.


