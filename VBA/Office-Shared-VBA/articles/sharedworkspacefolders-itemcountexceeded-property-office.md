---
title: SharedWorkspaceFolders.ItemCountExceeded Property (Office)
keywords: vbaof11.chm269005
f1_keywords:
- vbaof11.chm269005
ms.prod: office
api_name:
- Office.SharedWorkspaceFolders.ItemCountExceeded
ms.assetid: cc8f3b36-e9cc-ad08-c94d-85c2b909ee97
ms.date: 06/08/2017
---


# SharedWorkspaceFolders.ItemCountExceeded Property (Office)

Gets a  **Boolean** value that indicates whether the number of **SharedWorkspaceFolders** items in the collection has exceeded the 99 that can be displayed in the **Shared Workspace** task pane. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **ItemCountExceeded**

 _expression_ A variable that represents a **SharedWorkspaceFolders** object.


### Return Value

Boolean


## Remarks

The Shared Workspace task pane can only display 99 shared workspace files and folders, links, members, or tasks. If more than 99 items are added to any of these collections: the corresponding tab of the  **Shared Workspace** task pane will stop displaying the list of items and displays a link to the shared workspace site Web page instead; the collection is no longer populated locally and its **Count** property returns 0 (zero).

Furthermore, once the  **ItemCountExceeded** property returns **True** for one of the collections listed above, the developer can no longer remedy the situation programmatically by deleting items from the collection to reduce the count below 99, because the collection is no longer populated.

The  **ItemCountExceeded** property of the **SharedWorkspaceFolders** collection returns **True** when the combined count of files and folders exceeds 99, since both lists are combined and displayed together on the Documents tab of the Shared Workspace task pane.


## See also


#### Concepts


[SharedWorkspaceFolders Object](sharedworkspacefolders-object-office.md)
#### Other resources


[SharedWorkspaceFolders Object Members](sharedworkspacefolders-members-office.md)

