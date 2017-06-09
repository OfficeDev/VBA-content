---
title: CheckBox.MouseIcon Property (Outlook Forms Script)
keywords: olfm10.chm2001540
f1_keywords:
- olfm10.chm2001540
ms.prod: outlook
ms.assetid: 6d3e1fe9-a23e-44d3-e569-9c0969ebcf6e
ms.date: 06/08/2017
---


# CheckBox.MouseIcon Property (Outlook Forms Script)

Returns a  **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.


## Syntax

 _expression_. **MouseIcon**

 _expression_A variable that represents a  **CheckBox** object.


## Remarks

The  **MouseIcon** property is valid when the **[MousePointer](checkbox-mousepointer-property-outlook-forms-script.md)** property is set to 99. The mouse icon of an object is the image that appears when the user moves the mouse across that object.

To assign an image for the mouse pointer, you can either assign a picture to the  **MouseIcon** property or load a picture from a file using the **LoadPicture** function in Visual Basic Scripting Edition.


