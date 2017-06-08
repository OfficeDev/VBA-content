---
title: SharedWorkspaceTask.ModifiedBy Property (Office)
keywords: vbaof11.chm264009
f1_keywords:
- vbaof11.chm264009
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.ModifiedBy
ms.assetid: e18d400b-0e53-a599-e789-d47c78abec49
ms.date: 06/08/2017
---


# SharedWorkspaceTask.ModifiedBy Property (Office)

Gets the name of the user who last modified the object. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **ModifiedBy**

 _expression_ A variable that represents a **SharedWorkspaceTask** object.


### Return Value

String


## Remarks

For shared workspace objects, the  **ModifiedBy** property returns the display name stored in the **Name** property of the **SharedWorkspaceMember** object. The **SharedWorkspaceMember** object does not have a **ModifiedBy** property.


## See also


#### Concepts


[SharedWorkspaceTask Object](sharedworkspacetask-object-office.md)
#### Other resources


[SharedWorkspaceTask Object Members](sharedworkspacetask-members-office.md)

