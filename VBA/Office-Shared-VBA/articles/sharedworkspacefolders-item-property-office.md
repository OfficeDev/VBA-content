---
title: SharedWorkspaceFolders.Item Property (Office)
keywords: vbaof11.chm269001
f1_keywords:
- vbaof11.chm269001
ms.prod: office
api_name:
- Office.SharedWorkspaceFolders.Item
ms.assetid: 70916b0d-5cf7-b858-e215-d3cc948735fc
ms.date: 06/08/2017
---


# SharedWorkspaceFolders.Item Property (Office)

Gets a  **SharedWorkspaceFolder** object from the **Folders** collection of the shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ Required. A variable that represents a **[SharedWorkspaceFolders](sharedworkspacefolders-object-office.md)** object. The specified **SharedWorkspaceFolders** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|Returns the  **SharedWorkspaceFolder** at the position specified. The returned **SharedWorkspaceFolder** object does not correspond to the order in which the items are displayed in the **Shared Workspace** pane, and is not affected by re-sorting the display.|

## See also


#### Concepts


[SharedWorkspaceFolders Object](sharedworkspacefolders-object-office.md)
#### Other resources


[SharedWorkspaceFolders Object Members](sharedworkspacefolders-members-office.md)

