---
title: Recipients.ResolveAll Method (Outlook)
keywords: vbaol11.chm234
f1_keywords:
- vbaol11.chm234
ms.prod: outlook
api_name:
- Outlook.Recipients.ResolveAll
ms.assetid: 82404dc6-af4e-f32d-68b2-9451328b5ca6
ms.date: 06/08/2017
---


# Recipients.ResolveAll Method (Outlook)

Attempts to resolve all the  **[Recipient](recipient-object-outlook.md)** objects in the **[Recipients](recipients-object-outlook.md)** collection against the Address Book.


## Syntax

 _expression_ . **ResolveAll**

 _expression_ A variable that represents a **Recipients** object.


### Return Value

 **True** if all of the objects were resolved, **False** if one or more were not.


## Example

This Visual Basic for Applications (VBA) example uses the  **[ResolveAll](recipients-resolveall-method-outlook.md)** method to attempt to resolve all recipients and, if unsuccessful, displays a message box for each unresolved recipient.


```vb
Sub CheckRecipients() 
 
 Dim MyItem As Outlook.MailItem 
 
 Dim myRecipients As Outlook.Recipients 
 
 Dim myRecipient As Outlook.Recipient 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 Set myRecipients = myItem.Recipients 
 
 myRecipients.Add("Aaron Con") 
 
 myRecipients.Add("Nate Sun") 
 
 myRecipients.Add("Dan Wilson") 
 
 If Not myRecipients.ResolveAll Then 
 
 For Each myRecipient In myRecipients 
 
 If Not myRecipient.Resolved Then 
 
 MsgBox myRecipient.Name 
 
 End If 
 
 Next 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Recipients Object](recipients-object-outlook.md)

