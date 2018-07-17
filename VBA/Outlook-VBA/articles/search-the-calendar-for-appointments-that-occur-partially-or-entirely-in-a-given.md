---
title: Search the Calendar for Appointments that Occur Partially or Entirely in a Given Time Period
ms.prod: outlook
ms.assetid: 3ff170d3-f098-51ab-9ae4-0e71cc587bac
ms.date: 06/08/2017
---


# Search the Calendar for Appointments that Occur Partially or Entirely in a Given Time Period

This topic shows a code sample in Visual Basic for Applications (VBA) that uses a Jet query to search for appointments in the default calendar folder that occur in a given period of time with a specific start time and a specific end time. The query returns appointments that occur entirely within that period, starting on or after the start time and ending on or before the end time. 

The query also returns appointments that overlap with the period, including those that start before the period starts but end within the period, those that start within the period but end after the period ends, and those that start before the start time and end after the end time, overlapping with the entire time period. The returned results include recurring appointments.

You may think that querying for appointments that start on or after the start date, and end on or before the end date, should be the approach. This would translate to the following query:



```sql
[Start] >= myStart AND [End] <= myEnd
```

However, to reliably find all appointments that occur entirely within the time period  _and_ those that overlap with the time period, you need to use a query that looks for appointments that start on or before the end time of the period, and end on or after the start time of the time period. This would translate to the following query:



```sql
[Start] <= myEnd AND [End] >= myStart
```

Taking into consideration the appointments that overlap with the given time period is useful if you want to clear your calendar of all appointments that happen during that time period. In this case, querying only for appointments that start and end within the specified period would not be sufficient. 

The  `FindApptsInTimeFrame` function in the code sample first defines the time period to query, assigning the start time, `myStart`, as 12:00am on the current system date, and the end time,  `myEnd`, as five days after the start time. It obtains all the items in the default calendar folder. 

To include recurrent appointments in the query, it sets  ** [Items.IncludeRecurrences](items-includerecurrences-property-outlook.md)** to **True** and then sorts the items by the ** [AppointmentItem.Start](appointmentitem-start-property-outlook.md)** property. It then builds the query for all appointments that begin on or before `myEnd`, and end on or after  `myStart`. It then applies the query to items in the default calendar folder, using the  ** [Items.Restrict](items-restrict-method-outlook.md)** method, and then prints the start time of all the returned appointments.



```vb
Sub FindApptsInTimeFrame()
    Dim myStart As Date
    Dim myEnd As Date
    Dim oCalendar As Outlook.folder
    Dim oItems As Outlook.items
    Dim oResItems As Outlook.items
    Dim oAppt As Outlook.AppointmentItem
    Dim strRestriction As String
     
    myStart = Date
    myEnd = DateAdd("d", 5, myStart)
    Debug.Print "Start:", myStart
    Debug.Print "End:", myEnd
     
    Set oCalendar = Application.session.GetDefaultFolder(olFolderCalendar)
    Set oItems = oCalendar.items
     
    oItems.IncludeRecurrences = True
    oItems.Sort "[Start]"
     
    strRestriction = "[Start] <= '" &; Format$(myEnd, "mm/dd/yyyy hh:mm AMPM") _
    &; "' AND [End] >= '" &; Format(myStart, "mm/dd/yyyy hh:mm AMPM") &; "'"
    Debug.Print strRestriction
     
    'Restrict the Items collection
    Set oResItems = oItems.Restrict(strRestriction)
     
    For Each oAppt In oResItems
        Debug.Print oAppt.Start, oAppt.Subject
    Next
End Sub
```


