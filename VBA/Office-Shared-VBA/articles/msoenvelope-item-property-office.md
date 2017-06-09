---
title: MsoEnvelope.Item Property (Office)
keywords: vbaof11.chm11003
f1_keywords:
- vbaof11.chm11003
ms.prod: office
api_name:
- Office.MsoEnvelope.Item
ms.assetid: cc13343c-dea5-152f-b123-441a4120c22c
ms.date: 06/08/2017
---


# MsoEnvelope.Item Property (Office)

Gets a  **MailItem** object that can be used to send the document as an e-mail. Read-only.


## Syntax

 _expression_. **Item**

 _expression_ Required. A variable that represents a **[MsoEnvelope](msoenvelope-object-office.md)** object.


## Example

The following example sends the active Microsoft Word document as an e-mail to the e-mail address that you pass to the subroutine.


```
Sub SendMail(ByVal strRecipient As String) 
 
 'Use a With...End With block to reference the msoEnvelope object. 
 With Application.ActiveDocument.MailEnvelope 
 
 'Add some introductory text before the body of the e-mail message. 
 .Introduction = "Please read this and send me your comments." 
 
 'Return a MailItem object that you can use to send the document. 
 With .Item 
 
 'All of the mail item settings are saved with the document. 
 'When you add a recipient to the Recipients collection 
 'or change other properties these settings will persist. 
 
 .Recipients.Add strRecipient 
 .Subject = "Here is the document." 
 
 'The body of this message will be 
 'the content of the active document. 
 .Send 
 End With 
 End With 
End Sub
```


## See also


#### Concepts


[MsoEnvelope Object](msoenvelope-object-office.md)
#### Other resources


[MsoEnvelope Object Members](msoenvelope-members-office.md)

