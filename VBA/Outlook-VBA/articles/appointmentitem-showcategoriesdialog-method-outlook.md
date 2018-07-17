---
title: AppointmentItem.ShowCategoriesDialog Method (Outlook)
keywords: vbaol11.chm915
f1_keywords:
- vbaol11.chm915
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.ShowCategoriesDialog
ms.assetid: 5b79f252-ffce-a59d-873f-48efe467df3b
ms.date: 06/08/2017
---


# AppointmentItem.ShowCategoriesDialog Method (Outlook)

Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.


## Syntax

 _expression_ . **ShowCategoriesDialog**

 _expression_ A variable that represents an **AppointmentItem** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new appointment item, displays the item on the screen, and opens up the  **Show Categories** dialog box.


```vb
Sub Appointment() 
 
'Creates an appointment item to access ShowCategoriesDialog 
 
 Dim olApptItem As Outlook.AppointmentItem 
 
 'Create appointment item 
 
 Set olApptItem = Application.CreateItem(olAppointmentItem) 
 
 
 
 olApptItem.Body = "Please meet with me regarding these sales figures." 
 
 olApptItem.Recipients.Add ("Jeff Smith") 
 
 olApptItem.Subject = "Sales Reports" 
 
 'Display the item 
 
 olApptItem.Display 
 
 'Display the Show categories dialog 
 
 olApptItem.ShowCategoriesDialog 
 
 
 
End Sub
```


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

