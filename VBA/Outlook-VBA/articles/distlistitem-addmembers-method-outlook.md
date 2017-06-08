---
title: DistListItem.AddMembers Method (Outlook)
keywords: vbaol11.chm1154
f1_keywords:
- vbaol11.chm1154
ms.prod: outlook
api_name:
- Outlook.DistListItem.AddMembers
ms.assetid: 42e3e9f2-0c73-f612-049a-aa477add03fa
ms.date: 06/08/2017
---


# DistListItem.AddMembers Method (Outlook)

Adds new members to a distribution list.


## Syntax

 _expression_ . **AddMembers**( **_Recipients_** )

 _expression_ A variable that represents a **DistListItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Recipients_|Required| **[Recipients](recipients-object-outlook.md)**|The members to be added to the distribution list.|

## Example

This Microsoft Visual Basic for Applications (VBA) example creates a new distribution list and adds the current user and 'Dan Wilson' to the list. If the specified recipient is not valid, the  **AddMember** method will fail. Therefore, to run this example, replace 'Dan Wilson' with a valid recipient name.


```vb
Sub AddNewMembers() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myDistList As Outlook.DistListItem 
 
 Dim myTempItem As Outlook.MailItem 
 
 Dim myRecipients As Outlook.Recipients 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI 
 
 Set myDistList = Application.CreateItem(olDistributionListItem 
 
 Set myTempItem = Application.CreateItem(olMailItem 
 
 Set myRecipients = myTempItem.Recipients 
 
 myDistList.DLName = _ 
 
 InputBox("Enter the name of the new distribution list 
 
 myRecipients.Add myNameSpace.CurrentUser.Name 
 
 myRecipients.Add "Dan Wilson 
 
 myRecipients.ResolveAll 
 
 myDistList.AddMembers myRecipients 
 
 myDistList.Save 
 
 myDistList.Display 
 
End Sub
```


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

