---
title: AppointmentItem.PropertyChange Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.PropertyChange
ms.assetid: 82bb6104-ce62-8fb6-1472-d84fd36e94ac
ms.date: 06/08/2017
---


# AppointmentItem.PropertyChange Event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.


## Syntax

 _expression_ . **PropertyChange**( **_Name_** )

 _expression_ A variable that represents an **AppointmentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## Example

This Visual Basic for Applications (VBA) example uses the  **PropertyChange** event to prevent someone from disabling a reminder on an item.


```vb
Public WithEvents myItem As Outlook.AppointmentItem 
 
 
 
Sub Initialize_handler() 
 
 Set myItem = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items("Status Meeting") 
 
End Sub 
 
 
 
Private Sub myItem_PropertyChange(ByVal Name As String) 
 
 Select Case Name 
 
 Case "ReminderSet" 
 
 MsgBox "You may not remove a reminder on this item." 
 
 myItem.ReminderSet = True 
 
 Case Else 
 
 End Select 
 
End Sub
```


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

