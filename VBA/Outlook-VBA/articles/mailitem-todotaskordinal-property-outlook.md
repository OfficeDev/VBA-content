---
title: MailItem.ToDoTaskOrdinal Property (Outlook)
keywords: vbaol11.chm3038
f1_keywords:
- vbaol11.chm3038
ms.prod: outlook
api_name:
- Outlook.MailItem.ToDoTaskOrdinal
ms.assetid: d1ccb01a-0792-3779-3f94-eb5195a39bb0
ms.date: 06/08/2017
---


# MailItem.ToDoTaskOrdinal Property (Outlook)

Returns or sets a  **Date** value that represents the ordinal value of the task for the **[MailItem](mailitem-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **ToDoTaskOrdinal**

 _expression_ An expression that returns a **MailItem** object.


## Remarks

This property returns  **Null** ( **Nothing** in Visual Basic) if the **[IsMarkedAsTask](mailitem-ismarkedastask-property-outlook.md)** property is set to **False** .

This property is used to indicate how the task should be ordered within the parent groups, such as the  **Today** group or the **Tomorrow** group, of the **To-Do Bar**. The value used in this property does not have any relation to the values of the  **[TaskStartDate](mailitem-taskstartdate-property-outlook.md)** , **[TaskDueDate](mailitem-taskduedate-property-outlook.md)** , or **[TaskCompletedDate](mailitem-taskcompleteddate-property-outlook.md)** properties.


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

