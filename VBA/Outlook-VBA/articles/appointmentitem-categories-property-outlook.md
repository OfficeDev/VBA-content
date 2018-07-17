---
title: AppointmentItem.Categories Property (Outlook)
keywords: vbaol11.chm846
f1_keywords:
- vbaol11.chm846
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Categories
ms.assetid: 71360959-7a42-7aa8-579f-1e544a734dd0
ms.date: 06/08/2017
---


# AppointmentItem.Categories Property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

 _expression_ . **Categories**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList** , under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split** .


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new appointment, displays the appointment on the screen, and opens the  **Show Categories** dialog box. Finally, it displays the categories that the user assigned using **[AppointmentItem.ShowCategoriesDialog](appointmentitem-showcategoriesdialog-method-outlook.md)** . Replace 'Dan Wilson' with a valid recipient name before running the example.


```vb
Sub Appointment() 
 
 'Creates an appointment to access ShowCategoriesDialog 
 
 Dim olApptItem As Outlook.AppointmentItem 
 
 
 
 'Creates appointment item 
 
 Set olApptItem = Application.CreateItem(olAppointmentItem) 
 
 olApptItem.Body = "Please meet with me regarding these sales figures." 
 
 olApptItem.Recipients.Add ("Dan Wilson") 
 
 olApptItem.Subject = "Sales Reports" 
 
 'Display the appointment 
 
 olApptItem.Display 
 
 'Display the Show Categories dialog box 
 
 olApptItem.ShowCategoriesDialog 
 
 MsgBox olApptItem.Categories 
 
End Sub
```


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

