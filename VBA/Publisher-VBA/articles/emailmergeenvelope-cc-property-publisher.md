---
title: EmailMergeEnvelope.Cc Property (Publisher)
keywords: vbapb10.chm9043972
f1_keywords:
- vbapb10.chm9043972
ms.prod: publisher
api_name:
- Publisher.EmailMergeEnvelope.Cc
ms.assetid: d9e7704c-c45a-cf19-e0a8-8d55e1e82fc0
ms.date: 06/08/2017
---


# EmailMergeEnvelope.Cc Property (Publisher)

Gets or sets the  **MailMergeDataField** object that represents the data-source field (column) that lists the e-mail addresses of recipients you want to receive a carbon copy (CC) of the merged e-mail message. Read/write.


## Syntax

 _expression_. **Cc**

 _expression_A variable that represents an  **EmailMergeEnvelope** object.


### Return Value

MailMergeDataField


## Remarks

You must make certain that you assign the correct data-source field (the one that represents CC e-mail addresses) to the  **Cc** property. You can use the following line of code, which gets the value of the **Name** property of the **MailMergeDataField** object to which **Cc** is assigned, to ensure that you make the correct assignment:


```vb
Debug.Print ThisDocument.MailMerge.EmailMergeEnvelope.Cc.Name
```

For an example of how to set the  **Cc** property value, see the **[EmailMergeEnvelope](emailmergeenvelope-object-publisher.md)** object topic.


