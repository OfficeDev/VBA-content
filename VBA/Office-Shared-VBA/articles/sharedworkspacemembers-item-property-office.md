---
title: SharedWorkspaceMembers.Item Property (Office)
keywords: vbaof11.chm273001
f1_keywords:
- vbaof11.chm273001
ms.prod: office
api_name:
- Office.SharedWorkspaceMembers.Item
ms.assetid: b4ff3efc-6329-8a59-beb7-e060ca386ab5
ms.date: 06/08/2017
---


# SharedWorkspaceMembers.Item Property (Office)

Gets a  **SharedWorkspaceMember** object from the Members collection of the shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ Required. A variable that represents a **[SharedWorkspaceMembers](sharedworkspacemembers-object-office.md)** object. The specified **SharedWorkspaceMembers** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|Returns the  **SharedWorkspaceMember** at the position specified. The returned **SharedWorkspaceMember** object does not correspond to the order in which the items are displayed in the **Shared Workspace** pane, and is not affected by re-sorting the display.|

## See also


#### Concepts


[SharedWorkspaceMembers Object](sharedworkspacemembers-object-office.md)
#### Other resources


[SharedWorkspaceMembers Object Members](sharedworkspacemembers-members-office.md)

