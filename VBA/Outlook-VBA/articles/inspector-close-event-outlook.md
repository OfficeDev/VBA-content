---
title: Inspector.Close Event (Outlook)
keywords: vbaol11.chm467
f1_keywords:
- vbaol11.chm467
ms.prod: outlook
api_name:
- Outlook.Inspector.Close
ms.assetid: 5a83b3d3-6096-9e37-88b1-00f97c0bf8bd
ms.date: 06/08/2017
---


# Inspector.Close Event (Outlook)

Occurs when the inspector associated with a Microsoft Outlook item is being closed.


## Syntax

 _expression_ . **Close**( **_Cancel_** )

 _expression_ A variable that represents an **Inspector** object.


## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the close operation isn't completed and the inspector is left open. This event cannot be canceled.

If you use the  **[Close](inspector-close-method-outlook.md)** method to fire this event, it can only be canceled if the **Close** method uses the **olPromptForSave** argument.


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

