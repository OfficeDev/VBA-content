---
title: DistListItem.AddMember Method (Outlook)
keywords: vbaol11.chm1159
f1_keywords:
- vbaol11.chm1159
ms.prod: outlook
api_name:
- Outlook.DistListItem.AddMember
ms.assetid: 4c9b1310-1bbe-a5a1-9088-85efd18a7bf5
ms.date: 06/08/2017
---


# DistListItem.AddMember Method (Outlook)

Adds a new member to the specified distribution list. The distribution list contains  **[Recipient](recipient-object-outlook.md)** objects that represent valid e-mail addresses.


## Syntax

 _expression_ . **AddMember**( **_Recipient_** )

 _expression_ A variable that represents a **DistListItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Recipient_|Required| **Recipient**|The recipient to be added to the list.|

## Remarks

Use the  **[AddMembers](distlistitem-addmembers-method-outlook.md)** method to add multiple members to a given distribution list.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new  **DistributionList** object and adds a recipient to it. If the specified recipient is not valid, the **AddMember** method will fail. To run this example, replace 'Dan Wilson' with a valid recipient name.


```vb
Sub AddNewMember() 
 
 'Adds a member to a new distribution list 
 
 
 
 Dim objItem As Outlook.DistListItem 
 
 Dim objMail As Outlook.MailItem 
 
 Dim objRcpnt As Outlook.Recipient 
 
 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 
 
 Set objItem = Application.CreateItem(olDistributionListItem) 
 
 'Create recipient for distlist 
 
 Set objRcpnt = Application.Session.CreateRecipient("Dan Wilson") 
 
 objRcpnt.Resolve 
 
 objItem.AddMember objRcpnt 
 
 'Add note to list and display 
 
 objItem.DLName = "Northwest Sales Manager" 
 
 objItem.Body = "Regional Sales Manager - NorthWest" 
 
 objItem.Save 
 
 objItem.Display 
 
End Sub
```


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

