---
title: Folder.InAppFolderSyncObject Property (Outlook)
keywords: vbaol11.chm2008
f1_keywords:
- vbaol11.chm2008
ms.prod: outlook
api_name:
- Outlook.Folder.InAppFolderSyncObject
ms.assetid: d9e94fb7-add5-65d5-d2bc-e23bdfa11078
ms.date: 06/08/2017
---


# Folder.InAppFolderSyncObject Property (Outlook)

Returns or sets a  **Boolean** that determines if the specified folder will be synchronized with the e-mail server. Read/write.


## Syntax

 _expression_ . **InAppFolderSyncObject**

 _expression_ A variable that represents a **Folder** object.


## Remarks

If  **True** , this folder will be synchronized when the "Application Folders" **[SyncObject](syncobject-object-outlook.md)** is synchronized. If **False** , the folder will not synchronize.

This is equivalent to selecting the check box for this folder in the  **Application Folders** group on the **Send/Receive** dialog box.

If this property is set to  **True** , and the "Application Folders" **SyncObject** does not already exist, a **SyncObject** will be automatically created. The "Application Folders" **SyncObject** is the only **Send/Receive** group that can be programmatically modified.


## Example

The following Microsoft Visual Basic for Applications (VBA) example sets the Inbox folder to be synchronized when the "Application Folders"  **SyncObject** object is synchronized. The **InAppFolderSyncObject** property is used in conjunction with the **[AppFolders](syncobjects-appfolders-property-outlook.md)** property of the **[SyncObjects](syncobjects-object-outlook.md)** collection.


```vb
Public Sub appfolders() 
 Dim nsp As Outlook.NameSpace 
 Dim sycs As Outlook.SyncObjects 
 Dim syc As Outlook.SyncObject 
 Dim mpfInbox As Outlook.Folder 
 
 Set nsp = Application.GetNamespace("MAPI") 
 Set sycs = nsp.SyncObjects 
 'Return the Application Folder SyncObject. 
 Set syc = sycs.AppFolders 
 'Get the Inbox folder. 
 Set mpfInbox = nsp.GetDefaultFolder(olFolderInbox) 
 'Set the Inbox folder to be synchronized when the Application 
 'Folder's SyncObject is synchronized. 
 mpfInbox.InAppFolderSyncObject = True 
 'Start the synchronization. 
 syc.Start 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

