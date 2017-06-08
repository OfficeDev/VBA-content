---
title: NavigationFolders Object (Outlook)
keywords: vbaol11.chm3200
f1_keywords:
- vbaol11.chm3200
ms.prod: outlook
api_name:
- Outlook.NavigationFolders
ms.assetid: ecff93b8-0c3f-5f31-5b61-c46d2622d2af
ms.date: 06/08/2017
---


# NavigationFolders Object (Outlook)

Contains a set of  **[NavigationFolder](navigationfolder-object-outlook.md)** objects that represent the navigation folders associated with a navigation group.


## Remarks

Use the  **[NavigationFolders](navigationgroup-navigationfolders-property-outlook.md)** property of the **[NavigationGroup](navigationgroup-object-outlook.md)** object to return a **NavigationFolders** object.

Use the  **[Add](navigationfolders-add-method-outlook.md)** method to create a new **NavigationFolder** object based on an existing **[Folder](folder-object-outlook.md)** object and add it to the collection. Use the **[Item](navigationfolders-item-method-outlook.md)** method to return an existing **NavigationFolder** object contained by the **NavigationFolders** collection. Use the **[Remove](navigationfolders-remove-method-outlook.md)** method from the **[NavigationFolders](navigationfolders-object-outlook.md)** collection of the parent **[NavigationGroup](navigationgroup-object-outlook.md)** object.

Use the  **[NavigationFolderAdd](navigationgroups-navigationfolderadd-event-outlook.md)** and **[NavigationFolderRemove](navigationgroups-navigationfolderremove-event-outlook.md)** events to detect when a navigation folder is added or removed, respectively, from the **NavigationFolders** object. Use the **[SelectedChange](navigationgroups-selectedchange-event-outlook.md)** event to detect changes in selection state for navigation folders contained in the **NavigationFolders** object that are based on calendar folders.

Note that if you delete a  **Folder** using **[Folder.Delete](folder-delete-method-outlook.md)**, the deletion will be reflected automatically in the Navigation Pane and in the **NavigationFolders** collection, but because the synchronization between the actual folders and the Navigation Pane happens asynchronously, this will take a few milliseconds to complete.


## Methods



|**Name**|
|:-----|
|[Add](navigationfolders-add-method-outlook.md)|
|[Item](navigationfolders-item-method-outlook.md)|
|[Remove](navigationfolders-remove-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](navigationfolders-application-property-outlook.md)|
|[Class](navigationfolders-class-property-outlook.md)|
|[Count](navigationfolders-count-property-outlook.md)|
|[Parent](navigationfolders-parent-property-outlook.md)|
|[Session](navigationfolders-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
