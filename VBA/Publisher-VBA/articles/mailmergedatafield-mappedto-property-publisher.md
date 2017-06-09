---
title: MailMergeDataField.MappedTo Property (Publisher)
keywords: vbapb10.chm6422566
f1_keywords:
- vbapb10.chm6422566
ms.prod: publisher
api_name:
- Publisher.MailMergeDataField.MappedTo
ms.assetid: 067619e8-98fe-d0c2-2f50-96b50cf53de4
ms.date: 06/08/2017
---


# MailMergeDataField.MappedTo Property (Publisher)

Returns the name of the recipient field (column) in the master data source (combined mail-merge recipient list) that the parent  **MailMergeDataField** object is mapped to. Read-only.


## Syntax

 _expression_. **MappedTo**

 _expression_A variable that represents a  **MailMergeDataField** object.


### Return Value

String


## Remarks

The parent  **MailMergeDataField** object must represent a field (column) in a connected data source that is not the master data source (the combination of all connected data sources). The **MappedTo** property is not available for data fields in the data source represented by the **DataSource** property of the **MailMerge** object of the active **Document** object ( `ThisDocument.MailMerge.DataSource`).


