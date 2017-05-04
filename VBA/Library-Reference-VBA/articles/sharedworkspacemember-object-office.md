---
title: SharedWorkspaceMember Object (Office)
keywords: vbaof11.chm272000
f1_keywords:
- vbaof11.chm272000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SharedWorkspaceMember
ms.assetid: 4d5ec7d9-b7f2-cdcf-5db2-7429b7a08ed9
---


# SharedWorkspaceMember Object (Office)

Represents a user who has rights in a shared document workspace site.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Remarks

Use the  **SharedWorkspaceMember** object to manage users who have rights to participate in a shared workspace and to collaborate on the shared documents saved in the workspace site.

 The **Role** specified when the user is added as a member of the workspace (for example, "Reader" or "Contributor") determines that user's rights in the workspace and cannot be accessed or modified later through properties of the **SharedWorkspaceMember** object.

Use the  **Item** ( _index_ ) property of the **SharedWorkspaceMembers** collection to return a specific **SharedWorkspaceMember** object.

Use the  **SharedWorkspaceMember** object's three distinct name properties to retrieve identifying information about the member.


- the  **Name** property returns the members display name;
    
- the  **Email** property returns the member's e-mail address; and,
    
- the  **DomainName** property returns the member's domain and user name in the format `domain\user`.
    



## Example

The following example displays the number of members in the active document's shared workspace, along with their names, domain user names, and e-mail addresses.


```vb
    Dim swsMember As Office.SharedWorkspaceMember 
    Dim strMemberInfo As String 
    strMemberInfo = "The shared workspace contains " &; _ 
        ActiveWorkbook.SharedWorkspace.Members.Count &; " member(s)." &; vbCrLf 
    If ActiveWorkbook.SharedWorkspace.Members.Count > 0 Then 
        For Each swsMember In ActiveWorkbook.SharedWorkspace.Members 
            strMemberInfo = strMemberInfo &; swsMember.Name &; vbCrLf &; _ 
                " - " &; swsMember.DomainName &; vbCrLf &; _ 
                " - " &; swsMember.Email &; vbCrLf 
        Next 
    End If 
    MsgBox strMemberInfo, vbInformation + vbOKOnly, _ 
        "Members in Shared Workspace" 
    Set swsMember = Nothing 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

