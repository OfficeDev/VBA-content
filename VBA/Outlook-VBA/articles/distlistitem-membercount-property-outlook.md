---
title: DistListItem.MemberCount Property (Outlook)
keywords: vbaol11.chm1149
f1_keywords:
- vbaol11.chm1149
ms.prod: outlook
api_name:
- Outlook.DistListItem.MemberCount
ms.assetid: 56e3aa96-4e2a-bdf9-93a1-daa206fb8d30
ms.date: 06/08/2017
---


# DistListItem.MemberCount Property (Outlook)

Returns a  **Long** indicating the number of members in a distribution list. Read-only.


## Syntax

 _expression_ . **MemberCount**

 _expression_ A variable that represents a **DistListItem** object.


## Remarks

The value returned represents all members of the distribution list, including member distribution lists. Each member distribution list is counted as a single member. That is,  **MemberCount** is not an aggregate sum of the recipients in the distribution list plus recipients in member distribution lists. For example, if a distribution list contains 10 recipients plus one distribution list containing 15 recipients, **MemberCount** returns 11.


## Example

This Microsoft Visual Basic for Applications example steps through the default Contacts folder, and if it finds a distribution list with more than 20 members it displays the item.


```vb
Sub CheckDLs() 
 
 Dim myOlFolder As Outlook.Folder 
 
 Dim myOlItems As Outlook.Items 
 
 Dim myOlDistList As Outlook.DistListItem 
 
 Dim x as Integer 
 
 
 
 Set myOlFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts) 
 
 Set myOlItems = myOlFolder.Items 
 
 For x = 1 To myOlItems.Count 
 
 If TypeName(myOlItems.Item(x)) = "DistListItem" Then 
 
 Set myOlDistList = myOlItems.Item(x) 
 
 If myOlDistList.MemberCount > 20 Then 
 
 MsgBox myOlDistList.DLName &; " has more than 20 members." 
 
 myOlDistList.Display 
 
 End If 
 
 End If 
 
 Next x 
 
End Sub
```


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

