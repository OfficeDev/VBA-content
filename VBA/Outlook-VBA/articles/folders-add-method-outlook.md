---
title: Folders.Add Method (Outlook)
keywords: vbaol11.chm46
f1_keywords:
- vbaol11.chm46
ms.prod: outlook
api_name:
- Outlook.Folders.Add
ms.assetid: 20ced7ad-779c-a9b0-267e-6d729c0eb822
ms.date: 06/08/2017
---


# Folders.Add Method (Outlook)

Creates a new folder in the  **[Folders](folders-object-outlook.md)** collection.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Type_** )

 _expression_ A variable that represents a **Folders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The display name for the new folder.|
| _Type_|Optional| **Long**|The Outlook folder type for the new folder. If the folder type is not specified, the new folder will default to the same type as the folder in which it is created. Can be one of the following  **[OlDefaultFolders](oldefaultfolders-enumeration-outlook.md)** constants: **olFolderCalendar** , **olFolderContacts** , **olFolderDrafts** , **olFolderInbox** , **olFolderJournal** , **olFolderNotes** , or **olFolderTasks** . The constants **olFolderConflicts** , **olFolderDeletedItems** , **olFolderJunk** , **olFolderLocalFailures** , **olFolderManagedEmail** , **olFolderOutbox** , **olFolderRssSubscriptions** , **olFolderSentMail** , **olFolderServerFailures** , **olFolderSyncIssues** , **olFolderToDo** , and **olPublicFoldersAllPublicFolders** cannot be specified for this argument.|

### Return Value

A  **[Folder](folder-object-outlook.md)** object that represents the new folder.


## Example

This VBA example uses the  **Add** method to add the new folder named "My Contacts" to the current (default) Contacts folder.


```vb
Sub AddContactsFolder() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim myFolder As Outlook.Folder 
 Dim myNewFolder As Outlook.Folder 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderContacts) 
 Set myNewFolder = myFolder.Folders.Add("My Contacts") 
End Sub
```

This VBA example uses the  **Add** method to add three new folders in the Tasks folder. The first folder, "Notes Folder", will contain note items. The second folder, "Contacts Folder", will contain contact items. The third folder, ?Public Folder? will be a public folder. If the folders already exist, a message box will inform the user.




```vb
Sub AddFolders() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim myFolder As Outlook.Folder 
 Dim myNotesFolder As Outlook.Folder 
 Dim myContactsFolder As Outlook.Folder 
 Dim myPublicFolder As Outlook.Folder 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderTasks) 
 On Error GoTo ErrorHandler 
 Set myNotesFolder = myFolder.Folders.Add("Notes Folder", olFolderNotes) 
 Set myContactsFolder = myFolder.Folders.Add("Contacts Folder", olFolderContacts) 
 Set myPublicFolder = myFolder.Folders.Add("Public Folder", olPublicFoldersAllPublicFolders) 
 Exit Sub 
ErrorHandler: 
 MsgBox "This folder already exists!" 
 Resume Next 
End Sub
```


## See also


#### Concepts


[Folders Object](folders-object-outlook.md)

