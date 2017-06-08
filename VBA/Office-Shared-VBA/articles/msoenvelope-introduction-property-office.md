---
title: MsoEnvelope.Introduction Property (Office)
keywords: vbaof11.chm11001
f1_keywords:
- vbaof11.chm11001
ms.prod: office
api_name:
- Office.MsoEnvelope.Introduction
ms.assetid: f37129d4-2a68-1623-272b-f71dfdeec59b
ms.date: 06/08/2017
---


# MsoEnvelope.Introduction Property (Office)

Sets or gets the introductory text that is included with a document that is sent using the  **MsoEnvelope** object. The introductory text is included at the top of the document in the e-mail. Read/write.


## Syntax

 _expression_. **Introduction**

 _expression_ A variable that represents a **MsoEnvelope** object.


## Example

The following example sends the active Microsoft Word document as an e-mail to the e-mail address that you pass to the subroutine.


```
Sub SendMail(ByVal strRecipient As String) 
 
 'Use a With...End With block to reference the MsoEnvelope object. 
 With Application.ActiveDocument.MailEnvelope 
 
 'Add some introductory text before the body of the e-mail. 
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

