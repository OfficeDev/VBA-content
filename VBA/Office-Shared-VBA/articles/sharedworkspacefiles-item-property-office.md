---
title: SharedWorkspaceFiles.Item Property (Office)
keywords: vbaof11.chm267001
f1_keywords:
- vbaof11.chm267001
ms.prod: office
api_name:
- Office.SharedWorkspaceFiles.Item
ms.assetid: 40b3aa6d-a232-3aba-21e2-645db7900fb1
ms.date: 06/08/2017
---


# SharedWorkspaceFiles.Item Property (Office)

Gets a  **SharedWorkspaceFile** object from the **Files** collection of the shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ Required. A variable that represents a **[SharedWorkspaceFiles](sharedworkspacefiles-object-office.md)** object. The specified **SharedWorkspaceFiles** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|Returns the  **SharedWorkspaceFile** at the position specified. The returned **SharedWorkspaceFile** object does not correspond to the order in which the items are displayed in the **Shared Workspace** pane, and is not affected by re-sorting the display.|

## See also


#### Concepts


[SharedWorkspaceFiles Object](sharedworkspacefiles-object-office.md)
#### Other resources


[SharedWorkspaceFiles Object Members](sharedworkspacefiles-members-office.md)

