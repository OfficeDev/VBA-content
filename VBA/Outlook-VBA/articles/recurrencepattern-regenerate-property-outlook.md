---
title: RecurrencePattern.Regenerate Property (Outlook)
keywords: vbaol11.chm286
f1_keywords:
- vbaol11.chm286
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.Regenerate
ms.assetid: c1db398b-5f13-85e0-981d-795c8c7ac8ea
ms.date: 06/08/2017
---


# RecurrencePattern.Regenerate Property (Outlook)

Returns a  **Boolean** that indicates **True** if the task should be regenerated following this pass through the recurrence pattern. Read/write.


## Syntax

 _expression_ . **Regenerate**

 _expression_ A variable that represents a **RecurrencePattern** object.


## Remarks

This property is used to control the regeneration of the task as each occurrence of a recurring task is completed. It is only valid for tasks. It is not valid for appointments.

To create a recurrence pattern, you must first set the  **[RecurrenceType](recurrencepattern-recurrencetype-property-outlook.md)** property to set the frequency, then set the **Regenerate** property to **True** to regenerate the task. After setting **Regenerate** to **True** , do not set it to **False** . If you subsequently set **Regenerate** to **False** , then you should set up the recurrence pattern again by getting a new **[RecurrencePattern](recurrencepattern-object-outlook.md)** object.


## Example

This Visual Basic for Applications (VBA) example creates a task called "Oil Change" that recurs every three months and uses the  **Regenerate** property to set it to regenerate after each recurrence.


```vb
Sub CreateTaskOilChange() 
 
 Dim myItem As Outlook.TaskItem 
 
 Dim myPattern As Outlook.RecurrencePattern 
 
 
 
 Set myItem = Application.CreateItem(olTaskItem) 
 
 Set myPattern = myItem.GetRecurrencePattern 
 
 myPattern.RecurrenceType = olRecursMonthly 
 
 myPattern.Regenerate = True 
 
 myPattern.Interval = 3 
 
 myItem.Subject = "Oil Change" 
 
 myItem.Save 
 
 myItem.Display 
 
End Sub
```


## See also


#### Concepts


[RecurrencePattern Object](recurrencepattern-object-outlook.md)

