---
title: Delete All Items and Subfolders in the Deleted Items Folder
ms.prod: outlook
ms.assetid: 359a416b-43d4-396e-e348-5624c4ca3599
ms.date: 06/08/2017
---


# Delete All Items and Subfolders in the Deleted Items Folder

This topic shows a code sample in Visual Basic for Applications (VBA) that deletes all items and subfolders in the Deleted Items folder. 


 **Note**  When you delete items or folders from a collection, you must use a decrementing loop counter. An incrementing loop counter will fail.


You can only empty the Deleted Items folder and you cannot remove the folder itself. However, to delete subfolders of the Deleted Items folder, you can simply delete the subfolder without first deleting its contents.




```vb
Sub RemoveAllItemsAndFoldersInDeletedItems() 
 Dim oDeletedItems As Outlook.Folder 
 Dim oFolders As Outlook.Folders 
 Dim oItems As Outlook.Items 
 Dim i As Long 
 'Obtain a reference to deleted items folder 
 Set oDeletedItems = Application.Session.GetDefaultFolder(olFolderDeletedItems) 
 Set oItems = oDeletedItems.Items 
 For i = oItems.Count To 1 Step -1 
 oItems.Item(i).Delete 
 Next 
 Set oFolders = oDeletedItems.Folders 
 For i = oFolders.Count To 1 Step -1 
 oFolders.Item(i).Delete 
 Next 
End Sub
```


