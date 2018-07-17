---
title: SharedWorkspaceLinks.Item Property (Office)
keywords: vbaof11.chm271001
f1_keywords:
- vbaof11.chm271001
ms.prod: office
api_name:
- Office.SharedWorkspaceLinks.Item
ms.assetid: 30338f6d-47e2-9adf-eec6-a08122e9654e
ms.date: 06/08/2017
---


# SharedWorkspaceLinks.Item Property (Office)

Gets a  **SharedWorkspaceLink** object from the **Links** collection of the shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ Required. A variable that represents a **[SharedWorkspaceLinks](sharedworkspacelinks-object-office.md)** object. The specified **SharedWorkspaceLinks** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|Returns the  **SharedWorkspaceLink** at the position specified. The returned **SharedWorkspaceLink** object does not correspond to the order in which the items are displayed in the **Shared Workspace** pane, and is not affected by re-sorting the display.|

## See also


#### Concepts


[SharedWorkspaceLinks Object](sharedworkspacelinks-object-office.md)
#### Other resources


[SharedWorkspaceLinks Object Members](sharedworkspacelinks-members-office.md)

