---
title: SharedWorkspaceMembers.ItemCountExceeded Property (Office)
keywords: vbaof11.chm273005
f1_keywords:
- vbaof11.chm273005
ms.prod: office
api_name:
- Office.SharedWorkspaceMembers.ItemCountExceeded
ms.assetid: 9e647bd4-fc49-5e11-385a-79f6add9f335
ms.date: 06/08/2017
---


# SharedWorkspaceMembers.ItemCountExceeded Property (Office)

Gets a  **Boolean** value that indicates whether the number of **SharedWorkspaceMembers** items in the collection has exceeded the 99 that can be displayed in the **Shared Workspace** task pane. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **ItemCountExceeded**

 _expression_ A variable that represents a **SharedWorkspaceMembers** object.


### Return Value

Boolean


## Remarks

The  **Shared Workspace** task pane can only display 99 shared workspace files and folders, links, members, or tasks. If more than 99 items are added to any of these collections: the corresponding tab of the **Shared Workspace** task pane stops displaying the list of items and displays a link to the shared workspace site Web page instead; the collection is no longer populated locally and its **Count** property returns 0 (zero).

Furthermore, once the  **ItemCountExceeded** property returns **True** for one of the collections listed above, the developer can no longer remedy the situation programmatically by deleting items from the collection to reduce the count below 99, because the collection is no longer populated.


## See also


#### Concepts


[SharedWorkspaceMembers Object](sharedworkspacemembers-object-office.md)
#### Other resources


[SharedWorkspaceMembers Object Members](sharedworkspacemembers-members-office.md)

