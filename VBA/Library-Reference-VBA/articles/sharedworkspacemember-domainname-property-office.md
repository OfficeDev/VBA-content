---
title: SharedWorkspaceMember.DomainName Property (Office)
keywords: vbaof11.chm272001
f1_keywords:
- vbaof11.chm272001
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SharedWorkspaceMember.DomainName
ms.assetid: 2cbbea6f-7b2c-9ddc-7a37-2e2b6be10405
---


# SharedWorkspaceMember.DomainName Property (Office)

Gets the domain and user name of the specified [SharedWorkspaceMember](sharedworkspacemember-object-office.md) in the format domain\user. Read-only.


## Syntax

 _expression_. **DomainName**

 _expression_ A variable that represents a **SharedWorkspaceMember** object.


## Example

The following example extracts the domain name from the  **DomainName** property of each shared workspace member and lists members who belong to the "MyCompany" domain.


```vb
 Dim swsMember As Office.SharedWorkspaceMember 
 Dim strDomain As String 
 Dim strMemberList As String 
 For Each swsMember In ActiveWorkbook.SharedWorkspace.Members 
 strDomain = UCase(Left(swsMember.DomainName, _ 
 InStr(swsMember.DomainName, "\") - 1)) 
 If strDomain = "MYCOMPANY" Then 
 strMemberList = strMemberList &; swsMember.Name &; vbCrLf 
 End If 
 Next 
 MsgBox strMemberList, vbInformation + vbOKOnly, _ 
 "Members in the MYCOMPANY Domain" 
 Set swsMember = Nothing 

```


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## See also


#### Concepts


[SharedWorkspaceMember Object](sharedworkspacemember-object-office.md)

