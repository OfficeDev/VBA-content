---
title: Items.IncludeRecurrences Property (Outlook)
keywords: vbaol11.chm60
f1_keywords:
- vbaol11.chm60
ms.prod: outlook
api_name:
- Outlook.Items.IncludeRecurrences
ms.assetid: 7d192112-889c-56ce-aab2-107d751c80c4
ms.date: 06/08/2017
---


# Items.IncludeRecurrences Property (Outlook)

Returns a  **Boolean** that indicates **True** if the **[Items](items-object-outlook.md)** collection should include recurrence patterns. Read/write.


## Syntax

 _expression_ . **IncludeRecurrences**

 _expression_ A variable that represents an **Items** object.


## Remarks

This property only has an effect if the  **Items** collection contains appointments and is not sorted by any property other than **[Start](appointmentitem-start-property-outlook.md)** in ascending order. The default value is **False** . Use this property when you want to retrieve all appointments for a given date, where recurring appointments would not normally appear because they are not associated with any specific date. If you need to sort and filter on appointment items that contain recurring appointments, you must do so in this order: sort the items in ascending order, set **IncludeRecurrences** to **True** , and then filter the items. For a code sample showing this order, see the second example below. If the collection includes recurring appointments with no end date, setting the property to **True** may cause the collection to be of infinite count. Be sure to include a test for this in any loop. You should not use **Count** property of **Items** collection when iterating **Items** collection with **IncludeRecurrence** property set to **True** . The value of **Count** will be an undefined value.


 **Caution**  Filtering on a sorted list of occurrences will cause the  **IncludeRecurrences** property not to work as expected. For example, the following sequence will return all appointment occurrences; recurring and non-recurring: (1) Sort by Start property (2) Set property to **false** (3) call **Restrict** (i.e., filter).


## Example

The following Visual Basic for Applications (VBA) example displays the subject of the appointments that occur between today and tomorrow including recurring appointments.


```vb
Sub DemoFindNext() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim tdystart As Date 
 
 Dim tdyend As Date 
 
 Dim myAppointments As Outlook.Items 
 
 Dim currentAppointment As Outlook.AppointmentItem 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 tdystart = VBA.Format(Now, "Short Date") 
 
 tdyend = VBA.Format(Now + 1, "Short Date") 
 
 Set myAppointments = myNameSpace.GetDefaultFolder(olFolderCalendar).Items 
 
 myAppointments.Sort "[Start]" 
 
 myAppointments.IncludeRecurrences = True 
 
 Set currentAppointment = myAppointments.Find("[Start] >= """ &; _ 
 
 tdystart &; """ and [Start] <= """ &; tdyend &; """") 
 
 While TypeName(currentAppointment) <> "Nothing" 
 
 MsgBox currentAppointment.Subject 
 
 Set currentAppointment = myAppointments.FindNext 
 
 Wend 
 
End Sub
```

The example below shows the order to sort and filter on appointment items that contain recurring appointments.




```vb
Sub SortAndFilterAppointments() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myAppointments As Outlook.Items 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set calendarItems = myNameSpace.GetDefaultFolder(olFolderCalendar).Items 
 
 calendarItems.Sort "[Start]" 
 
 calendarItems.IncludeRecurrences = True 
 
 Set restrictedItems = calendarItems.Restrict("[Organizer]='Dan Wilson'") 
 
End Sub
```


## See also


#### Concepts


[Items Object](items-object-outlook.md)

