---
title: Application.Reminder Event (Outlook)
keywords: vbaol11.chm431
f1_keywords:
- vbaol11.chm431
ms.prod: outlook
api_name:
- Outlook.Application.Reminder
ms.assetid: f8c9fa87-3daa-58e1-7b8d-3c819cd4cab2
ms.date: 06/08/2017
---


# Application.Reminder Event (Outlook)

Occurs immediately before a reminder is displayed.


## Syntax

 _expression_ . **Reminder**( **_Item_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The  **[AppointmentItem](appointmentitem-object-outlook.md)** , **[MailItem](mailitem-object-outlook.md)** , **[ContactItem](contactitem-object-outlook.md)** , or **[TaskItem](taskitem-object-outlook.md)** associated with the reminder. If the appointment associated with the reminder is a recurring appointment, _Item_ is the specific occurrence of the appointment that displayed the reminder, not the master appointment.|

## Example

This Microsoft Visual Basic for Applications (VBA) example displays the item that fired the  **Reminder** event when the event fires. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim WithEvents myolapp As Outlook.Application 
 
 
 
Sub Initialize_handler() 
 
 Set myolapp = Outlook.Application 
 
End Sub 
 
 
 
Private Sub myolapp_Reminder(ByVal Item As Object) 
 
 Item.Display 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

