---
title: Document.MailEnvelope Property (Publisher)
keywords: vbapb10.chm196627
f1_keywords:
- vbapb10.chm196627
ms.prod: publisher
api_name:
- Publisher.Document.MailEnvelope
ms.assetid: 3c4c734a-6725-5f6e-ed0a-5b19e4e642bd
ms.date: 06/08/2017
---


# Document.MailEnvelope Property (Publisher)

Returns an  **MsoEnvelope** object that represents an e-mail header for a publication.


## Syntax

 _expression_. **MailEnvelope**

 _expression_A variable that represents a  **Document** object.


### Return Value

MsoEnvelope


## Remarks

The  **MailEnvelope** property is only accessible if the **[EnvelopeVisible](document-envelopevisible-property-publisher.md)** property has been set to **True**.


## Example

This example sets the comments for the e-mail header of the active publication. This example assumes that the  **EnvelopeVisible** property has been set to **True**.


```vb
Sub HeaderComments() 
 ActiveDocument.MailEnvelope.Introduction = _ 
 "Please review this publication and let me know " &; _ 
 "what you think. I need your input by Friday." &; _ 
 " Thanks." 
End Sub
```


