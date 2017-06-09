---
title: AppointmentItem.Save Method (Outlook)
keywords: vbaol11.chm874
f1_keywords:
- vbaol11.chm874
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Save
ms.assetid: 177980e8-96cc-a72e-ede3-7aad3a98cf68
ms.date: 06/08/2017
---


# AppointmentItem.Save Method (Outlook)

Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.


## Syntax

 _expression_ . **Save**

 _expression_ A variable that represents an **AppointmentItem** object.


## Example

This Microsoft Visual Basic for Applications (VBA) example creates an appointment item and sets the  **[AppointmentItem.ReminderSet](appointmentitem-reminderset-property-outlook.md)** property before saving it.


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
#### Other resources



[How to: Import Appointment XML Data into Outlook Appointment Objects](http://msdn.microsoft.com/library/ecfd3849-877b-01ad-2b76-1a54e980f6e2%28Office.15%29.aspx)

