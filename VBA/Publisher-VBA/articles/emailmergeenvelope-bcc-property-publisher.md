---
title: EmailMergeEnvelope.Bcc Property (Publisher)
keywords: vbapb10.chm9043974
f1_keywords:
- vbapb10.chm9043974
ms.prod: publisher
api_name:
- Publisher.EmailMergeEnvelope.Bcc
ms.assetid: 1d846fac-d93c-6a20-ce3b-090525dbbfe1
ms.date: 06/08/2017
---


# EmailMergeEnvelope.Bcc Property (Publisher)

Gets or sets a semicolon-delimited list of e-mail addresses that receive a blind carbon copy (BCC) of the e-mail message. Read/write.


## Syntax

 _expression_. **Bcc**

 _expression_A variable that represents an  **EmailMergeEnvelope** object.


### Return Value

String


## Remarks

Set the  **Bcc** property to a string of e-mail addresses separated by semicolons, as shown in the following example.


```vb
 MailMerge.EmailMergeEnvelope.Bcc = "name1@address1;name2@address2;name3@address3;..."
```


