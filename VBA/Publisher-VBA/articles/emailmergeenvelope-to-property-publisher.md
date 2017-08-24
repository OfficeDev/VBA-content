---
title: EmailMergeEnvelope.To Property (Publisher)
keywords: vbapb10.chm9043971
f1_keywords:
- vbapb10.chm9043971
ms.prod: publisher
api_name:
- Publisher.EmailMergeEnvelope.To
ms.assetid: c9c470e8-1411-fda9-becf-5c932e97d98f
ms.date: 06/08/2017
---


# EmailMergeEnvelope.To Property (Publisher)

Gets or sets the  **MailMergeDataField** object that represents the data-source field (column) that lists the e-mail addresses of recipients of the merged e-mail message. Read/write.


## Syntax

 _expression_. **To**

 _expression_A variable that represents an  **EmailMergeEnvelope** object.


### Return Value

MailMergeDataField


## Remarks

You must make certain that you assign the correct data-source field (the one that represents e-mail addresses) to the  **To** property. You can use the following line of code, which gets the value of the **Name** property of the **MailMergeDataField** object to which **To** is assigned, to ensure that you make the correct assignment:


```vb
Debug.Print ThisDocument.MailMerge.EmailMergeEnvelope.To.Name
```

For an example of how to set the  **To** property value, see the **[EmailMergeEnvelope](emailmergeenvelope-object-publisher.md)** object topic.


