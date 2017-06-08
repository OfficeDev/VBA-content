---
title: Items.SetColumns Method (Outlook)
keywords: vbaol11.chm71
f1_keywords:
- vbaol11.chm71
ms.prod: outlook
api_name:
- Outlook.Items.SetColumns
ms.assetid: 90206a68-baf8-282c-5793-fee029fed452
ms.date: 06/08/2017
---


# Items.SetColumns Method (Outlook)

Caches certain properties for extremely fast access to those particular properties of each item in an  **[Items](items-object-outlook.md)** collection.


## Syntax

 _expression_ . **SetColumns**( **_Columns_** )

 _expression_ A variable that represents an **Items** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Columns_|Required| **String**|A string that contains the names of the properties to cache. The property names are delimited by commas in this string.|

## Remarks

The  **SetColumns** method is useful for iterating through an **Items** collection. If you don't use this method, Microsoft Outlook must open each item to access the property. With the **SetColumns** method, Outlook only checks the properties that you have cached, and provides fast, read-only access to these properties.

After applying the  **SetColumns** method on specific properties of the collection, you cannot read other properties of that collection; properties which are not cached are returned empty. You cannot write to any of the properties of that collection either. Alternatively, if you require read-write, fast access to a set of items, use the **[Table](table-object-outlook.md)** object.

 **SetColumns** cannot be used, and will cause an error, with any property that returns an object. It cannot be used with the following properties:



| **AutoResolvedWinner**| **InternetCodePage**|
| **Body**| **MeetingWorkspaceURL**|
| **BodyFormat**| **[MemberCount](distlistitem-membercount-property-outlook.md)**|
| **Categories**| **ReceivedByEntryID**|
| **[Children](contactitem-children-property-outlook.md)**| **ReceivedOnBehalfOfEntryID**|
| **Class**| **[RecurrenceState](appointmentitem-recurrencestate-property-outlook.md)**|
| **Companies**| **ReplyRecipients**|
| **[DLName](distlistitem-dlname-property-outlook.md)**| **[ResponseState](taskitem-responsestate-property-outlook.md)**|
| **DownloadState**| **Saved**|
| **EntryID**| **Sent**|
| **HTMLBody**| **Submitted**|
| **IsConflict**| **[VotingOptions](mailitem-votingoptions-property-outlook.md)**|
The  **ConversationIndex** property cannot be cached using the **SetColumns** method. However, this property will not result in an error like the other properties listed above.


## Example

The following Visual Basic for Applications (VBA) example uses the  **[Items](items-object-outlook.md)** collection to get the items in default Tasks folder, caches the **[Subject](mailitem-subject-property-outlook.md)** and **[DueDate](taskitem-duedate-property-outlook.md)** properties and then displays the subject and due dates each in turn.


```vb
Sub SortByDueDate() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItem As Object 
 
 Dim myItems As Outlook.Items 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderTasks) 
 
 Set myItems = myFolder.Items 
 
 myItems.SetColumns ("Subject, DueDate") 
 
 For Each myItem In myItems 
 
 MsgBox myItem.Subject &; " " &; myItem.DueDate 
 
 Next myItem 
 
End Sub
```


## See also


#### Concepts


[Items Object](items-object-outlook.md)

