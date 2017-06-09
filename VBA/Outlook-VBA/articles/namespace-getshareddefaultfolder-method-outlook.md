---
title: NameSpace.GetSharedDefaultFolder Method (Outlook)
keywords: vbaol11.chm765
f1_keywords:
- vbaol11.chm765
ms.prod: outlook
api_name:
- Outlook.NameSpace.GetSharedDefaultFolder
ms.assetid: e2196423-e4f2-2797-c16c-dc54e2c0f7d2
ms.date: 06/08/2017
---


# NameSpace.GetSharedDefaultFolder Method (Outlook)

Returns a  **[Folder](folder-object-outlook.md)** object that represents the specified default folder for the specified user.


## Syntax

 _expression_ . **GetSharedDefaultFolder**( **_Recipient_** , **_FolderType_** )

 _expression_ A variable that represents a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Recipient_|Required| **[Recipient](recipient-object-outlook.md)**|The owner of the folder. Note that the  **Recipient** object must be resolved.|
| _FolderType_|Required| **[OlDefaultFolders](oldefaultfolders-enumeration-outlook.md)**|The type of folder.|

### Return Value

A  **Folder** object that represents the specified default folder for the specified user.


## Remarks

This method is used in a delegation scenario, where one user has delegated access to another user for one or more of their default folders (for example, their shared  **Calendar** folder).

 _FolderType_ can be one of the following **OlDefaultFolders** constants: **olFolderCalendar** , **olFolderContacts** , **olFolderDrafts** , **olFolderInbox** , **olFolderJournal** , **olFolderNotes** , or **olFolderTasks** . (The constants **olFolderDeletedItems** , **olFolderOutbox** , **olFolderJunk** , **olFolderConflicts** , **olFolderLocalFailures** , **olFolderServerFailures** , **olFolderSyncIssues** , **olPublicFoldersAllPublicFolders** , **olFolderRssSubscriptions** , **olFolderToDo** , **olFolderManagedEmail** , and **olFolderSentMail** cannot be specified for this argument.)


## Example

This Visual Basic for Applications (VBA) example uses the  **GetSharedDefaultFolder** method to resolve the **Recipient** object representing Dan Wilson, and then returns Dan's shared default **Calendar** folder.


```vb
Sub ResolveName() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myRecipient As Outlook.Recipient 
 
 Dim CalendarFolder As Outlook.Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myRecipient = myNamespace.CreateRecipient("Dan Wilson") 
 
 myRecipient.Resolve 
 
 If myRecipient.Resolved Then 
 
 Call ShowCalendar(myNamespace, myRecipient) 
 
 End If 
 
End Sub 
 
 
 
Sub ShowCalendar(myNamespace, myRecipient) 
 
 Dim CalendarFolder As Outlook.Folder 
 
 
 
 Set CalendarFolder = _ 
 
 myNamespace.GetSharedDefaultFolder _ 
 
 (myRecipient, olFolderCalendar) 
 
 CalendarFolder.Display 
 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

