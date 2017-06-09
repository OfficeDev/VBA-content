---
title: OlkCommandButton.Picture Property (Outlook)
keywords: vbaol11.chm1000496
f1_keywords:
- vbaol11.chm1000496
ms.prod: outlook
api_name:
- Outlook.OlkCommandButton.Picture
ms.assetid: 68b60b14-1a26-4b62-2770-5c3e16cf96b5
ms.date: 06/08/2017
---


# OlkCommandButton.Picture Property (Outlook)

Returns or sets a  **StdPicture** value that represents the picture that is displayed on the control. Read/write.


## Syntax

 _expression_ . **Picture**

 _expression_ A variable that represents an **OlkCommandButton** object.


## Remarks

The picture is of the Microsoft Windows type  **StdPicture** . The default value is **Null** ( **Nothing** in Visual Basic).

A picture and text cannot be displayed at the same time on the control, so when the picture property is set, the text property is ignored.

The picture is always displayed in the center of the button control. The picture will be clipped as necessary to fit in the available space.


## See also


#### Concepts


[OlkCommandButton Object](olkcommandbutton-object-outlook.md)

