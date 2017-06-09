---
title: TaskItem.StatusOnCompletionRecipients Property (Outlook)
keywords: vbaol11.chm1745
f1_keywords:
- vbaol11.chm1745
ms.prod: outlook
api_name:
- Outlook.TaskItem.StatusOnCompletionRecipients
ms.assetid: 9800dcb7-6b12-af4b-0379-25658c946118
ms.date: 06/08/2017
---


# TaskItem.StatusOnCompletionRecipients Property (Outlook)

Returns or sets a semicolon-delimited  **String** of display names for recipients who will receive status upon completion of the task. Read/write.


## Syntax

 _expression_ . **StatusOnCompletionRecipients**

 _expression_ A variable that represents a **TaskItem** object.


## Remarks

This property is calculated from the  **[Recipients](taskitem-recipients-property-outlook.md)** property. Recipients returned by the **StatusOnCompletionRecipients** property correspond to BCC recipients in the **[Recipients](recipients-object-outlook.md)** collection.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

