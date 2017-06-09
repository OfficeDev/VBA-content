---
title: SharedWorkspaceMember.Name Property (Office)
keywords: vbaof11.chm272002
f1_keywords:
- vbaof11.chm272002
ms.prod: office
api_name:
- Office.SharedWorkspaceMember.Name
ms.assetid: 6a7918a0-6029-4fe1-6c55-d100a360eddc
ms.date: 06/08/2017
---


# SharedWorkspaceMember.Name Property (Office)

Gets the display name of the shared workspace member. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **SharedWorkspaceMember** object.


### Return Value

String


## Example

The following example displays properties of the shared workspace member.


```
    Dim swsWorkspaceMember As Office.SharedWorkspaceMember 
    Dim strSWSInfo As String 
    Set swsWorkspaceMember = ActiveWorkbook.SharedWorkspace.Members 
    strSWSInfo = swsWorkspaceMember.Name &amp; vbCrLf &amp; _ 
        " - URL: " &amp; swsWorkspaceMember.URL &amp; vbCrLf 
    MsgBox strSWSInfo, vbInformation + vbOKOnly, _ 
        "Shared Workspace Member Information" 
    Set swsWorkspaceMember = Nothing 

```


## See also


#### Concepts


[SharedWorkspaceMember Object](sharedworkspacemember-object-office.md)
#### Other resources


[SharedWorkspaceMember Object Members](sharedworkspacemember-members-office.md)

