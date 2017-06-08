---
title: Image.ControlTipText Property (Access)
keywords: vbaac10.chm10388
f1_keywords:
- vbaac10.chm10388
ms.prod: access
api_name:
- Access.Image.ControlTipText
ms.assetid: 13a3cadf-8a2e-3407-5fa8-d76e3b2c9cac
ms.date: 06/08/2017
---


# Image.ControlTipText Property (Access)

You can use the  **ControlTipText** property to specify the text that appears in a ScreenTip when you hold the mouse pointer over a control. Read/write **String**.


## Syntax

 _expression_. **ControlTipText**

 _expression_ A variable that represents an **Image** object.


## Remarks

You set the  **ControlTipText** property by using a string expression up to 255 characters long.

For controls on forms, you can set the default for this property by using the default control style or the  **DefaultControl** property in Visual Basic.

You can set the  **ControlTipText** property in any view.

The  **ControlTipText** property provides an easy way to provide helpful information about controls on a form.

There are other ways to provide information about a form or a control on a form. You can use the  **StatusBarText** property to display information in the status bar about a control. To provide more extensive help for a form or control, use the **HelpFile** and **HelpContextID** properties.


## See also


#### Concepts


[Image Object](image-object-access.md)

