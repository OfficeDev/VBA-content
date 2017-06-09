---
title: Setting Items for Follow-up
ms.prod: outlook
ms.assetid: 738e2558-2957-54fb-898d-b67a6462dc66
ms.date: 06/08/2017
---


# Setting Items for Follow-up

Microsoft Outlook provides a new task flagging system in which certain Outlook items such as mail items or contact items can be flagged for follow-up. Flagging an Outlook item for follow-up displays information about that Outlook item, along with other task-based information, on the  **To-Do Bar** and **Calendar** navigation module in the Outlook user interface.

The following Outlook item objects have been extended to support the task flagging system:

-  **[ContactItem](contactitem-object-outlook.md)**
    
-  **[DistListItem](distlistitem-object-outlook.md)**
    
-  **[MailItem](mailitem-object-outlook.md)**
    
-  **[PostItem](postitem-object-outlook.md)**
    
-  **[SharingItem](sharingitem-object-outlook.md)**
    

## Marking an Item as a Task

You can determine if an Outlook item object is marked for follow-up by checking the value of the  **[IsMarkedAsTask](mailitem-ismarkedastask-property-outlook.md)** property for an Outlook item. Use the **[MarkAsTask](mailitem-markastask-method-outlook.md)** method to mark an Outlook item for follow-up and the **[ClearTaskFlag](mailitem-cleartaskflag-method-outlook.md)** method to unmark the Outlook item.


## Setting Task Properties

When an Outlook item is marked for follow-up using the  **MarkAsTask** method, an OlMarkInterval constant is used to specify default settings for the **[TaskStartDate](mailitem-taskstartdate-property-outlook.md)**,  **[TaskDueDate](mailitem-taskduedate-property-outlook.md)**,  **[TaskCompletedDate](mailitem-taskcompleteddate-property-outlook.md)**, and  **[ToDoTaskOrdinal](mailitem-todotaskordinal-property-outlook.md)** properties of the Outlook item. These properties are used not only to determine the duration and completion state of the task associated with the Outlook item, but also to determine the order in which the Outlook item is displayed in the **To-Do Bar** and **Calendar** navigation module.

However, you can programmatically set these properties individually, after calling the  **MarkAsTask** method, to support custom durations, or to change the completion state or display order of the Outlook item.

Once an Outlook item is flagged for follow-up, you can also set the  **[TaskSubject](mailitem-tasksubject-property-outlook.md)** property of the Outlook item to display a task description other than the value of the **Subject** property for the flagged Outlook item.


## Task Items and Task Flagging

The  **[TaskItem](taskitem-object-outlook.md)** object supports the **[ToDoTaskOrdinal](taskitem-todotaskordinal-property-outlook.md)** property, so that the display order for Outlook task items displayed on the **To-Do Bar** can also be changed programmatically.


## Filtering Items Marked as Tasks

You can take advantage of the DAV Searching and Locating (DASL) filtering capabilities of Outlook to filter Outlook items marked for follow-up. The following Visual Basic for Applications (VBA) example defines a DASL filter that filters only those Outlook items with an  **IsMarkedAsTask** property value set to **True**, then uses the filter to build a  **[Table](table-object-outlook.md)** object containing filtered Outlook items retrieved from the Inbox default folder.


```vb
Private Sub TableForIsMarkedAsTask() 
 Dim objTable As Outlook.Table 
 Dim objRow As Outlook.Row 
 Dim strFilter As String 
 
 On Error GoTo ErrRoutine 
 
 ' Define a DASL filter string that filters only those items 
 ' with an IsMarkedAsTask property value set to True. 
 strFilter = "@SQL=" &; Chr(34) &; _ 
 "http://schemas.microsoft.com/mapi/proptag/0x0E2B0003" &; _ 
 Chr(34) &; " = 1" 
 
 ' Use the filter to construct a table of Outlook items 
 ' retrieved from the Inbox default folder. 
 Set objTable = Application.Session.GetDefaultFolder(olFolderInbox).GetTable(strFilter) 
 
 With objTable 
 ' Add task-related columns to the table. 
 .Columns.Add ("From") 
 .Columns.Add ("FlagRequest") 
 .Columns.Add ("TaskStartDate") 
 .Columns.Add ("TaskDueDate") 
 .Columns.Add ("TaskCompletedDate") 
 
 ' Report the contents of the table 
 ' to the Immediate window. 
 Do Until .EndOfTable 
 Set objRow = .GetNextRow 
 Debug.Print objRow("Subject"), _ 
 objRow("From"), _ 
 objRow("FlagRequest"), _ 
 objRow("TaskStartDate"), _ 
 objRow("TaskDueDate"), _ 
 objRow("TaskCompletedDate") 
 Loop 
 End With 
 
EndRoutine: 
 ' Clean up 
 Set objRow = Nothing 
 Set objTable = Nothing 
 
 Exit Sub 
 
ErrRoutine: 
 MsgBox Err.Number &; " - " &; Err.Description, _ 
 vbOKOnly Or vbCritical, _ 
 "TableForIsMarkedAsTask" 
 
 GoTo EndRoutine 
End Sub
```


