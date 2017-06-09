---
title: SharedWorkspaceMember.Email Property (Office)
keywords: vbaof11.chm272003
f1_keywords:
- vbaof11.chm272003
ms.prod: office
api_name:
- Office.SharedWorkspaceMember.Email
ms.assetid: 3539becc-bde4-9331-432c-e907523975a7
ms.date: 06/08/2017
---


# SharedWorkspaceMember.Email Property (Office)

Gets the e-mail name of the specified  **SharedWorkspaceMember** in the format user@domain.com. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 **Email**_expression_. **Email**

 _expression_ An expression that returns a **SharedWorkspaceMember** object.


## Example

The following example extracts the e-mail domain name from the  **Email** property of each shared workspace member and lists members who have e-mail addresses at the "example.com" domain.


```
Dim swsMember As Office.SharedWorkspaceMember 
    Dim strEmailDomain As String 
    Dim strMemberList As String 
    For Each swsMember In ActiveWorkbook.SharedWorkspace.Members 
        strEmailDomain = LCase(Right(swsMember.Email, _ 
            Len(swsMember.Email) - InStr(swsMember.Email, "@"))) 
        If strEmailDomain = "example.com" Then 
            strMemberList = strMemberList &amp; swsMember.Email &amp; vbCrLf 
        End If 
    Next 
    MsgBox strMemberList, vbInformation + vbOKOnly, _ 
        "Members with example.com e-mail" 
    Set swsMember = Nothing
```


## See also


#### Concepts


[SharedWorkspaceMember Object](sharedworkspacemember-object-office.md)
#### Other resources


[SharedWorkspaceMember Object Members](sharedworkspacemember-members-office.md)

