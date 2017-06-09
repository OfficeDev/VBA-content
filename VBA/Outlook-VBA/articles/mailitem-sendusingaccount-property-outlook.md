---
title: MailItem.SendUsingAccount Property (Outlook)
keywords: vbaol11.chm1390
f1_keywords:
- vbaol11.chm1390
ms.prod: outlook
api_name:
- Outlook.MailItem.SendUsingAccount
ms.assetid: d4e49128-a63a-d761-90b9-9e1a3305adc7
ms.date: 06/08/2017
---


# MailItem.SendUsingAccount Property (Outlook)

Returns or sets an  **[Account](account-object-outlook.md)** object that represents the account under which the **[MailItem](mailitem-object-outlook.md)** is to be sent. Read/write.


## Syntax

 _expression_ . **SendUsingAccount**

 _expression_ An expression that returns a **MailItem** object.


## Remarks

The  **SendUsingAccount** property can be used to specify the account that should be used to send the **MailItem** when the **[Send](mailitem-send-method-outlook.md)** method is called. This property returns **Null** ( **Nothing** in Visual Basic) if the account specified for the **MailItem** no longer exists.


## Example

The following code sample in Microsoft Visual Basic for Applications enumerates the  **[Accounts](accounts-object-outlook.md)** collection to find a Pop3 account. If the account is found, then a message is created programmatically and the **SendUsingAccount** property is assigned to the Pop3 account. Note that you must assign the **SendUsingAccount** property before you call the **Send** method.


```vb
Sub SendUsingAccount() 
 
 Dim oAccount As Outlook.account 
 
 For Each oAccount In Application.Session.Accounts 
 
 If oAccount.AccountType = olPop3 Then 
 
 Dim oMail As Outlook.MailItem 
 
 Set oMail = Application.CreateItem(olMailItem) 
 
 oMail.Subject = "Sent using POP3 Account" 
 
 oMail.Recipients.Add ("someone@example.com") 
 
 oMail.Recipients.ResolveAll 
 
 oMail.SendUsingAccount = oAccount 
 
 oMail.Send 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)
#### Other resources


[Send an E-mail Given the SMTP Address of an Account](send-an-e-mail-given-the-smtp-address-of-an-account-outlook.md)



