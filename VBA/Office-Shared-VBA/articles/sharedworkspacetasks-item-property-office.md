---
title: SharedWorkspaceTasks.Item Property (Office)
keywords: vbaof11.chm265001
f1_keywords:
- vbaof11.chm265001
ms.prod: office
api_name:
- Office.SharedWorkspaceTasks.Item
ms.assetid: 801adcf2-ed06-fbe3-39c6-15fcc72c25fb
ms.date: 06/08/2017
---


# SharedWorkspaceTasks.Item Property (Office)

Gets a  **SharedWorkspaceTask** object from the **Tasks** collection of the shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ Required. A variable that represents a **[SharedWorkspaceTasks](sharedworkspacetasks-object-office.md)** object. The specified **SharedWorkspaceTasks** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|Returns the  **SharedWorkspaceTask** at the position specified. The returned **SharedWorkspaceTask** object does not correspond to the order in which the items are displayed in the **Shared Workspace** pane, and is not affected by re-sorting the display.|

## See also


#### Concepts


[SharedWorkspaceTasks Object](sharedworkspacetasks-object-office.md)
#### Other resources


[SharedWorkspaceTasks Object Members](sharedworkspacetasks-members-office.md)

