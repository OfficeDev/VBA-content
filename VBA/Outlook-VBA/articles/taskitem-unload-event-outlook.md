---
title: TaskItem.Unload Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskItem.Unload
ms.assetid: ff7d2655-06b5-6344-3422-4bf7be761a39
ms.date: 06/08/2017
---


# TaskItem.Unload Event (Outlook)

Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. 


## Syntax

 _expression_ . **Unload**

 _expression_ A variable that represents a **TaskItem** object.


## Remarks

This event occurs after the  **Close** event for the Outlook item occurs, but before the Outlook item is unloaded from memory, allowing an add-in to release any resources related to the object. Although the event occurs before the Outlook item is unloaded from memory, this event cannot be canceled.


 **Note**  This event is meant only as a notification event, so that an add-in can dereference the object. An error occurs if any property or method for this object is called within the  **Unload** event.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

