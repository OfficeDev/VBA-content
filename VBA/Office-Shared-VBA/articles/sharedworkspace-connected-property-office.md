---
title: SharedWorkspace.Connected Property (Office)
keywords: vbaof11.chm276012
f1_keywords:
- vbaof11.chm276012
ms.prod: office
api_name:
- Office.SharedWorkspace.Connected
ms.assetid: 071502b9-c4f7-45f5-062b-818d5859708e
ms.date: 06/08/2017
---


# SharedWorkspace.Connected Property (Office)

Gets a  **Boolean** value that indicates whether or not the active document is currently saved in and connected to a shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Connected**

 _expression_ A variable that represents a **SharedWorkspace** object.


## Remarks

Use the  **[Disconnect](sharedworkspace-disconnect-method-office.md)** method of the **SharedWorkspace** object to disconnect the local copy of the active document from the shared workspace. Use the **[RemoveDocument](sharedworkspace-removedocument-method-office.md)** method to remove the document from the shared workspace.


## Example

The following example checks the  **Connected** property to determine whether the active document is already saved in a shared workspace.


```
    If ActiveWorkbook.SharedWorkspace.Connected Then 
        MsgBox "This document is already saved in a shared workspace." 
    End If 

```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

