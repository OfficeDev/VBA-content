---
title: WebBrowserControl.EventProcPrefix Property (Access)
keywords: vbaac10.chm14683
f1_keywords:
- vbaac10.chm14683
ms.prod: ACCESS
api_name:
- Access.WebBrowserControl.EventProcPrefix
ms.assetid: 8dbf1fee-d9ab-ff0c-5571-e606c19fbf94
---


# WebBrowserControl.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents a **WebBrowserControl** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

