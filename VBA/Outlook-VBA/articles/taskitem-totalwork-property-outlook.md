---
title: TaskItem.TotalWork Property (Outlook)
keywords: vbaol11.chm1748
f1_keywords:
- vbaol11.chm1748
ms.prod: outlook
api_name:
- Outlook.TaskItem.TotalWork
ms.assetid: 3b940a69-f2b4-30d1-0027-49450f547b01
ms.date: 06/08/2017
---


# TaskItem.TotalWork Property (Outlook)

Returns or sets a  **Long** indicating the total work for the task. Read/write.


## Syntax

 _expression_ . **TotalWork**

 _expression_ A variable that represents a **TaskItem** object.


## Remarks

 **TotalWork** corresponds to the **Total work** field on the **Details** tab of a Task item. It is stored in units of minutes. The **Total work** field on the standard task form is bound to the **TotalWork** property; by default the field assumes an 8-hour day and 40-hour week.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

