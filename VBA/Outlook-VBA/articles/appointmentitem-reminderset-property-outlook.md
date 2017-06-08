---
title: AppointmentItem.ReminderSet Property (Outlook)
keywords: vbaol11.chm895
f1_keywords:
- vbaol11.chm895
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.ReminderSet
ms.assetid: 575d5fb2-1672-ddae-832c-7dcc7d1da2d6
ms.date: 06/08/2017
---


# AppointmentItem.ReminderSet Property (Outlook)

Returns or sets a  **Boolean** value that is **True** if a reminder has been set for this item. Read/write.


## Syntax

 _expression_ . **ReminderSet**

 _expression_ A variable that represents an **AppointmentItem** object.


## Example

This example creates an appointment item and sets the  **ReminderSet** property before saving it.


```vb
Sub AddAppointment() 
 
 Dim apti As Outlook.AppointmentItem 
 
 
 
 Set apti = Application.CreateItem(olAppointmentItem) 
 
 apti.Subject = "Car Servicing" 
 
 apti.Start = DateAdd("n", 16, Now) 
 
 apti.End = DateAdd("n", 60, apti.Start) 
 
 apti.ReminderSet = True 
 
 apti.ReminderMinutesBeforeStart = 60 
 
 apti.Save 
 
End Sub
```


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

