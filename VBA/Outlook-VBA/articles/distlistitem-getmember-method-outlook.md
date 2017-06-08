---
title: DistListItem.GetMember Method (Outlook)
keywords: vbaol11.chm1156
f1_keywords:
- vbaol11.chm1156
ms.prod: outlook
api_name:
- Outlook.DistListItem.GetMember
ms.assetid: 97196e1f-02a5-c1ac-be93-841702abaf52
ms.date: 06/08/2017
---


# DistListItem.GetMember Method (Outlook)

Returns a  **[Recipient](recipient-object-outlook.md)** object representing a member in a distribution list.


## Syntax

 _expression_ . **GetMember**( **_Index_** )

 _expression_ A variable that represents a **DistListItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the member to be retrieved.|

### Return Value

A  **Recipient** object representing the specified member.


## Example

This Microsoft Visual Basic for Applications (VBA) example locates every distribution list in the default  **Contacts** folder and determines whether the list contains the current user.


```vb
Sub DisplayYourDLNames() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myDistList As Outlook.DistListItem 
 
 Dim myFolderItems As Outlook.Items 
 
 Dim x As Integer 
 
 Dim y As Integer 
 
 Dim iCount As Integer 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderContacts) 
 
 Set myFolderItems = myFolder.Items 
 
 iCount = myFolderItems.Count 
 
 For x = 1 To iCount 
 
 If TypeName(myFolderItems.Item(x)) = "DistListItem" Then 
 
 Set myDistList = myFolderItems.Item(x) 
 
 For y = 1 To myDistList.MemberCount 
 
 If myDistList.GetMember(y).Name = myNameSpace.CurrentUser.Name Then 
 
 MsgBox "Your are a member of " &; myDistList.DLName 
 
 End If 
 
 Next y 
 
 End If 
 
 Next x 
 
End Sub
```


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

