---
title: PostItem.ToDoTaskOrdinal Property (Outlook)
keywords: vbaol11.chm3042
f1_keywords:
- vbaol11.chm3042
ms.prod: outlook
api_name:
- Outlook.PostItem.ToDoTaskOrdinal
ms.assetid: 58847d68-b956-3d87-6ed2-127801d3fee3
ms.date: 06/08/2017
---


# PostItem.ToDoTaskOrdinal Property (Outlook)

Returns or sets a  **Date** value that represents the ordinal value of the task for the **[PostItem](postitem-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **ToDoTaskOrdinal**

 _expression_ An expression that returns a **PostItem** object.


## Remarks

This property returns  **Null** ( **Nothing** in Visual Basic) if the **[IsMarkedAsTask](postitem-ismarkedastask-property-outlook.md)** property is set to **False** .

This property is used to indicate how the task should be ordered within the parent groups, such as the  **Today** group or the **Tomorrow** group, of the **To-Do Bar**. The value used in this property does not have any relation to the values of the  **[TaskStartDate](postitem-taskstartdate-property-outlook.md)** , **[TaskDueDate](postitem-taskduedate-property-outlook.md)** , or **[TaskCompletedDate](postitem-taskcompleteddate-property-outlook.md)** properties.


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

