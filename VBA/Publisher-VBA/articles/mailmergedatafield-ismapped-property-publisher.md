---
title: MailMergeDataField.IsMapped Property (Publisher)
keywords: vbapb10.chm6422565
f1_keywords:
- vbapb10.chm6422565
ms.prod: publisher
api_name:
- Publisher.MailMergeDataField.IsMapped
ms.assetid: 4a053a2b-f6ca-37a7-4a1f-8690982188c2
ms.date: 06/08/2017
---


# MailMergeDataField.IsMapped Property (Publisher)

Indicates if the parent  **MailMergeDataField** object is mapped to a recipient field in the master data source (combined mail-merge recipient list). Read-only.


## Syntax

 _expression_. **IsMapped**

 _expression_A variable that represents a  **MailMergeDataField** object.


### Return Value

Boolean


## Remarks

The parent  **MailMergeDataField** object must represent a field (column) in a connected data source that is not the master data source (the combination of all connected data sources). The **IsMapped** property is not available for data fields in the data source represented by the **DataSource** property of the **MailMerge** object of the active **Document** object ( `ThisDocument.MailMerge.DataSource`).


