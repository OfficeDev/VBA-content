---
title: TaskItem.SendUsingAccount Property (Outlook)
keywords: vbaol11.chm1768
f1_keywords:
- vbaol11.chm1768
ms.prod: outlook
api_name:
- Outlook.TaskItem.SendUsingAccount
ms.assetid: 711382c3-1003-cf0e-2f29-fc3f9d4320a8
ms.date: 06/08/2017
---


# TaskItem.SendUsingAccount Property (Outlook)

Returns or sets an  **[Account](account-object-outlook.md)** object that represents the account under which the **[TaskItem](taskitem-object-outlook.md)** object is to be sent. Read/write.


## Syntax

 _expression_ . **SendUsingAccount**

 _expression_ An expression that returns a **TaskItem** object.


## Remarks

The  **SendUsingAccount** property can be used to specify the account that should be used to send the **TaskItem** object when the **[Send](taskitem-send-method-outlook.md)** method is called. This property returns **Null** ( **Nothing** in Visual Basic) if the account specified for the **TaskItem** object no longer exists.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

