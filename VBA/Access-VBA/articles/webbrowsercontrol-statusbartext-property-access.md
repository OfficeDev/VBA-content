---
title: WebBrowserControl.StatusBarText Property (Access)
keywords: vbaac10.chm14681
f1_keywords:
- vbaac10.chm14681
ms.prod: access
api_name:
- Access.WebBrowserControl.StatusBarText
ms.assetid: 8d2daa59-e8aa-103a-ce26-99fe8a1eae80
ms.date: 06/08/2017
---


# WebBrowserControl.StatusBarText Property (Access)

You can use the  **StatusBarText** property to specify the text that is displayed in the status bar when a control is selected. Read/write **String**.


## Syntax

 _expression_. **StatusBarText**

 _expression_ A variable that represents a **WebBrowserControl** object.


## Remarks

You set the  **StatusBarText** property by using a string expression up to 255 characters long.


 **Note**  The length of the text you can display in the status bar depends on your computer hardware and video display.

You can use the  **StatusBarText** property to provide specific information about a control. For example, when a text box has the focus, a brief instruction can tell the user what kind of data to enter


 **Note**  You can also use the  **ControlTipText** property to display a ScreenTip for a control.

If you create a control by dragging a field from the field list, the value in a field's  **Description** property is copied to the **StatusBarText** property.


## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

