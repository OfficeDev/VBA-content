---
title: SharedWorkspaceMembers.Add Method (Office)
keywords: vbaof11.chm273003
f1_keywords:
- vbaof11.chm273003
ms.prod: office
api_name:
- Office.SharedWorkspaceMembers.Add
ms.assetid: 13d7c75d-a4d1-60ea-d689-c6886fb1e898
ms.date: 06/08/2017
---


# SharedWorkspaceMembers.Add Method (Office)

Adds a member to the list of members in a shared workspace site. Returns a  **[SharedWorkspaceMember](sharedworkspacemember-object-office.md)** object.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Add**( **_Email_**, **_DomainName_**, **_DisplayName_**, **_Role_** )

 _expression_ Required. A variable that represents a **[SharedWorkspaceMembers](sharedworkspacemembers-object-office.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Email_|Required|**String**|The new member's e-mail address in the format user@domain.com. Raises an error if the user is not a valid candidate for membership in the shared workspace site.|
| _DomainName_|Required|**String**|The new member's Windows user name in the format domain\user.|
| _DisplayName_|Required|**String**|The display name to display for the new member.|
| _Role_|Optional|**String**|An optional role that determines the tasks the new member can accomplish in the shared workspace site; for example, "Contributor". An invalid role name raises an error.|

## Example

The following example adds a new member to the members collection of the shared workspace site in the role of a site contributor.


```
    Dim swsMember As Office.SharedWorkspaceMember 
    Set swsMember = ActiveWorkbook.SharedWorkspace.Members.Add( _ 
        "user@domain.com", _ 
        "domain\user", _ 
        "New User", _ 
        "Contributor") 
    MsgBox "New member: " &amp; swsMember.Name, _ 
        vbInformation + vbOKOnly, _ 
        "New Member in Shared Workspace)" 
    Set swsMember = Nothing 

```


## See also


#### Concepts


[SharedWorkspaceMembers Object](sharedworkspacemembers-object-office.md)
#### Other resources


[SharedWorkspaceMembers Object Members](sharedworkspacemembers-members-office.md)

