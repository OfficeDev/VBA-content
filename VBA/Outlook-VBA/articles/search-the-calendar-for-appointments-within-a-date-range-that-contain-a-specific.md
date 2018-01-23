---
title: Search the Calendar for Appointments Within a Date Range that Contain a Specific Word in the Subject
ms.prod: outlook
ms.assetid: 92b6f569-e10e-d2cd-c941-0f062183d2bd
ms.date: 06/08/2017
---


# Search the Calendar for Appointments Within a Date Range that Contain a Specific Word in the Subject

This topic shows a code sample in Visual Basic for Applications (VBA) that finds appointments in the default calendar that occur strictly within the next thirty days and that contain the word "team" in the subject. The returned results include recurrent appointments.

The  `FindAppts` function in the code sample carries out the search using two different queries, first searching for appointments including recurrent appointments that start and end within the date range, then searching among appointments that meet the date range criteria that have "team" in the subject. The following is an outline of the steps:

1.  `FindAppts` first defines the time period to query, assigning the start time, `myStart`, as 12:00am on the current system date, and the end time,  `myEnd`, as thirty days after the start time. 
    
2. It obtains all the items in the default calendar folder.
    
3. To include all appointment items strictly within the date range including recurrent appointments, it sets  ** [Items.IncludeRecurrences](items-includerecurrences-property-outlook.md)** to **True** and then sorts the items by the ** [AppointmentItem.Start](appointmentitem-start-property-outlook.md)** property.
    
4. It builds the first query for all appointments that begin on or after  `myStart`, and end on or before  `myEnd`. This query is a Jet query.
    
5. It applies the query to items in the default calendar folder, using the  ** [Items.Restrict](items-restrict-method-outlook.md)** method.
    
6. It builds the second query for the appointment subject containing the word "team". It uses the  `like` keyword for substring matching in a DAV Searching and Locating (DASL) query.
    
7. It applies the second query to the set of appointments that meet the date range criteria, returned from the first query.
    
8. It sorts and prints the start time of all the final returned appointments.
    

Note that if you want to include appointment items that overlap and do not fall strictly within the specific date range, you should change the first query to one that have appointments begin on or before  `myEnd`, and end on or after  `myStart`. For more information, see  [How to: Search the Calendar for Appointments that Occur Partially or Entirely in a Given Time Period](search-the-calendar-for-appointments-that-occur-partially-or-entirely-in-a-given.md).



```vb
Sub FindAppts()

    Dim myStart As Date
    Dim myEnd As Date
    Dim oCalendar As Outlook.folder
    Dim oItems As Outlook.items
    Dim oItemsInDateRange As Outlook.items
    Dim oFinalItems As Outlook.items
    Dim oAppt As Outlook.AppointmentItem
    Dim strRestriction As String

    myStart = Date
    myEnd = DateAdd("d", 30, myStart)

    Debug.Print "Start:", myStart
    Debug.Print "End:", myEnd
          
    'Construct filter for the next 30-day date range
    strRestriction = "[Start] >= '" & _
    Format$(myStart, "mm/dd/yyyy hh:mm AMPM") _
    & "' AND [End] <= '" & _
    Format$(myEnd, "mm/dd/yyyy hh:mm AMPM") & "'"
    'Check the restriction string
    Debug.Print strRestriction
    Set oCalendar = Application.session.GetDefaultFolder(olFolderCalendar)
    Set oItems = oCalendar.items
    oItems.IncludeRecurrences = True
    oItems.Sort "[Start]"
    'Restrict the Items collection for the 30-day date range
    Set oItemsInDateRange = oItems.Restrict(strRestriction)
    'Construct filter for Subject containing 'team'
    Const PropTag  As String = "http://schemas.microsoft.com/mapi/proptag/"
    strRestriction = "@SQL=" & Chr(34) & PropTag _
        & "0x0037001E" & Chr(34) & " like '%team%'"
    'Restrict the last set of filtered items for the subject
    Set oFinalItems = oItemsInDateRange.Restrict(strRestriction)
    'Sort and Debug.Print final results
    oFinalItems.Sort "[Start]"
    For Each oAppt In oFinalItems
        Debug.Print oAppt.Start, oAppt.Subject
    Next
End Sub
```


