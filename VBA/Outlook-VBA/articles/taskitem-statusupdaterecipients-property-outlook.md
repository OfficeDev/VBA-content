---
title: TaskItem.StatusUpdateRecipients Property (Outlook)
keywords: vbaol11.chm1746
f1_keywords:
- vbaol11.chm1746
ms.prod: outlook
api_name:
- Outlook.TaskItem.StatusUpdateRecipients
ms.assetid: 904e4685-75db-9267-7f88-dd2bce6e8509
ms.date: 06/08/2017
---


# TaskItem.StatusUpdateRecipients Property (Outlook)

Returns a semicolon-delimited  **String** of display names for recipients who receive status updates for the task. Read/write.


## Syntax

 _expression_ . **StatusUpdateRecipients**

 _expression_ A variable that represents a **TaskItem** object.


## Remarks

This property is calculated from the  **[Recipients](taskitem-recipients-property-outlook.md)** property. Recipients returned by the **StatusUpdateRecipients** property correspond to CC recipients in the **[Recipients](recipients-object-outlook.md)** collection.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

