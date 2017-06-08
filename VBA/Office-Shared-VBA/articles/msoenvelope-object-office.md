---
title: MsoEnvelope Object (Office)
keywords: vbaof11.chm245000
f1_keywords:
- vbaof11.chm245000
ms.prod: office
api_name:
- Office.MsoEnvelope
ms.assetid: 64cfde6b-cd71-1d7b-0e8f-1181d88d9457
ms.date: 06/08/2017
---


# MsoEnvelope Object (Office)

Provides access to functionality that lets you send documents as e-mail messages directly from Microsoft Office applications.


## Remarks

Use the  **MailEnvelope** property of the **Document** object, **Chart** object or **Worksheet** object (depending on the application you are using) to return a **MsoEnvelope** object.


## Example

The following example sends the active Microsoft Word document as an e-mail message to the e-mail address that you pass to the subroutine.


```
Sub SendMail(ByVal strRecipient As String) 
 
 'Use a With...End With block to reference the MsoEnvelope object. 
 With Application.ActiveDocument.MailEnvelope 
 
 'Add some introductory text before the body of the e-mail. 
 .Introduction = "Please read this and send me your comments." 
 
 'Return a Microsoft Outlook MailItem object that 
 'you can use to send the document. 
 With .Item 
 
 'All of the mail item settings are saved with the document. 
 'When you add a recipient to the Recipients collection 
 'or change other properties, these settings persist. 
 .Recipients.Add strRecipient 
 .Subject = "Here is the document." 
 
 'The body of this message will be 
 'the content of the active document. 
 .Send 
 End With 
 End With 
End Sub
```


## Events



|**Name**|
|:-----|
|[EnvelopeHide](msoenvelope-envelopehide-event-office.md)|
|[EnvelopeShow](msoenvelope-envelopeshow-event-office.md)|

## Properties



|**Name**|
|:-----|
|[CommandBars](msoenvelope-commandbars-property-office.md)|
|[Introduction](msoenvelope-introduction-property-office.md)|
|[Item](msoenvelope-item-property-office.md)|
|[Parent](msoenvelope-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
