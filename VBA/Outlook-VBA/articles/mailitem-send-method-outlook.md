---
title: MailItem.Send Method (Outlook)
keywords: vbaol11.chm1369
f1_keywords:
- vbaol11.chm1369
ms.prod: outlook
api_name:
- Outlook.MailItem.Send
ms.assetid: 78c85013-523e-447b-c47d-2da0705f1fe0
ms.date: 06/08/2017
---


# MailItem.Send Method (Outlook)

Sends the e-mail message.


## Syntax

 _expression_ . **Send**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

The  **Send** method sends an item using the default account specified for the session. In a session where multiple Microsoft Exchange accounts are defined in the profile, the first Exchange account added to the profile is the primary Exchange account, and is also the default account for the session. To specify a different account to send an item, set the **[SendUsingAccount](mailitem-sendusingaccount-property-outlook.md)** property to the desired **[Account](account-object-outlook.md)** object and then call the **Send** method.


## Example

If you use Microsoft Visual Basic Scripting Edition (VBScript) in an Outlook form, you do not create the  **[Application](application-object-outlook.md)** object, and you cannot use named constants. This example shows how to forward a mail item using VBScript code.


```vb
Sub CommandButton1_Click() 
 Set myNameSpace = Application.GetNameSpace("MAPI") 
 Set myFolder = myNameSpace.GetDefaultFolder(6) 
 Set myForward = myFolder.Items(1).Forward 
 myForward.Recipients.Add "Laura Jennings" 
 myForward.Send 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)
#### Other resources


[Send an E-mail Given the SMTP Address of an Account](send-an-e-mail-given-the-smtp-address-of-an-account-outlook.md)



