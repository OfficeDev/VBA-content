---
title: SharingItem.ToDoTaskOrdinal Property (Outlook)
keywords: vbaol11.chm3222
f1_keywords:
- vbaol11.chm3222
ms.prod: outlook
api_name:
- Outlook.SharingItem.ToDoTaskOrdinal
ms.assetid: 4164fa78-c0cf-e359-2707-025d6d49f145
ms.date: 06/08/2017
---


# SharingItem.ToDoTaskOrdinal Property (Outlook)

Returns or sets a  **Date** value that represents the ordinal value of the task for the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **ToDoTaskOrdinal**

 _expression_ An expression that returns a **SharingItem** object.


## Remarks

This property returns  **Null** ( **Nothing** in Visual Basic) if the **[IsMarkedAsTask](sharingitem-ismarkedastask-property-outlook.md)** property is set to **False** .

This property is used to indicate how the task should be ordered within the parent groups, such as the  **Today** group or the **Tomorrow** group, of the **To-Do Bar**. The value used in this property does not have any relation to the values of the  **[TaskStartDate](sharingitem-taskstartdate-property-outlook.md)** , **[TaskDueDate](sharingitem-taskduedate-property-outlook.md)** , or **[TaskCompletedDate](sharingitem-taskcompleteddate-property-outlook.md)** properties.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

