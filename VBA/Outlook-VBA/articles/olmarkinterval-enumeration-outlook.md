---
title: OlMarkInterval Enumeration (Outlook)
keywords: vbaol11.chm3115
f1_keywords:
- vbaol11.chm3115
ms.prod: outlook
api_name:
- Outlook.OlMarkInterval
ms.assetid: a653146c-8a28-72dd-4ca7-98d8454c6f1f
ms.date: 06/08/2017
---


# OlMarkInterval Enumeration (Outlook)

Specifies the time period for which an Outlook item is marked as a task.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olMarkComplete**|5|Mark the task as complete.|
| **olMarkNextWeek**|3|Mark the task due next week.|
| **olMarkNoDate**|4|Mark the task due with no date.|
| **olMarkThisWeek**|2|Mark the task due this week.|
| **olMarkToday**|0|Mark the task due today.|
| **olMarkTomorrow**|1|Mark the task due tomorrow.|

## Remarks

This enumeration is used by the  **MarkAsTask** method for the following Outlook items:


- [ContactItem](contactitem-object-outlook.md)
    
- [DistListItem](distlistitem-object-outlook.md)
    
- [MailItem](mailitem-object-outlook.md)
    
- [PostItem](postitem-object-outlook.md)
    
- [SharingItem](sharingitem-object-outlook.md)
    
Depending on the value chosen, the following properties are set to the specified default values.



|** **Enumeration values****|** **Property values****|
|:-----|:-----|
| **olMarkComplete**| **TaskCompletedDate** is set to the current date and time ( **Now** in Visual Basic) if the item has been marked as a task. **TaskCompletedDate** is set to the empty date value (#1/1/4501#) if the item has not been marked as a task, or if the task has already been marked complete.|
| **olMarkNextWeek**| **IsMarkedAsTask** is set to **True**. **TaskSubject** is set to the value of the **Subject** property for the Outlook item. **TaskStartDate** is set to the first working day of next week. **TaskDueDate** is set to the last working day of next week. **TaskCompletedDate** is set to the empty date value (#1/1/4501#). **ToDoTaskOrdinal** is set to the current date and time ( **Now** in Visual Basic).|
| **olMarkNoDate**| **IsMarkedAsTask** is set to **True**. **TaskSubject** is set to the value of the **Subject** property for the Outlook item. **TaskStartDate**,  **TaskDueDate**, and  **TaskCompletedDate** are set to **Null** ( **Nothing** in Visual Basic). **ToDoTaskOrdinal** is set to the current date and time ( **Now** in Visual Basic).|
| **olMarkThisWeek**| **IsMarkedAsTask** is set to **True**. **TaskSubject** is set to the value of the **Subject** property for the Outlook item. **TaskStartDate** is set to a date two working days ahead of the current date. If that value would exceed the value of **TaskDueDate**, then  **TaskStartDate** is set to the value of **TaskDueDate**. **TaskDueDate** is set to the last working day of the current week. **TaskCompletedDate** is set to the empty date value (#1/1/4501#). **ToDoTaskOrdinal** is set to the current date and time ( **Now** in Visual Basic).|
| **olMarkToday**| **IsMarkedAsTask** is set to **True**. **TaskSubject** is set to the value of the **Subject** property for the Outlook item. **TaskStartDate** and **TaskDueDate** are set to the current date. **TaskCompletedDate** is set to the empty date value (#1/1/4501#). **ToDoTaskOrdinal** is set to the current date and time ( **Now** in Visual Basic).|
| **olMarkTomorrow**| **IsMarkedAsTask** is set to **True**. **TaskSubject** is set to the value of the **Subject** property for the Outlook item. **TaskStartDate** and **TaskDueDate** are set to one day after the current date. **TaskCompletedDate** is set to the empty date value (#1/1/4501#). **ToDoTaskOrdinal** is set to the current date and time ( **Now** in Visual Basic).|

