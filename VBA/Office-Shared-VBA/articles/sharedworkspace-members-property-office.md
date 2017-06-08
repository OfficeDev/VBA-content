---
title: SharedWorkspace.Members Property (Office)
keywords: vbaof11.chm276002
f1_keywords:
- vbaof11.chm276002
ms.prod: office
api_name:
- Office.SharedWorkspace.Members
ms.assetid: a53cfd41-36ca-73e4-08b2-306569f26979
ms.date: 06/08/2017
---


# SharedWorkspace.Members Property (Office)

Gets a  **[SharedWorkspaceMembers](sharedworkspacemembers-object-office.md)** collection that represents the list of members in the current shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Members**

 _expression_ A variable that represents a **SharedWorkspace** object.


## Example

The following example lists the members in the current shared workspace.


```
Dim swsMembers As Office.SharedWorkspaceMembers 
    Set swsMembers = ActiveWorkbook.SharedWorkspace.Members 
    MsgBox "There are " &amp; swsMembers.Count &amp; _ 
        " member(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsMembers = Nothing 

```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

