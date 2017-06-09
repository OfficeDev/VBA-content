---
title: MailMergeDataSource.ValidatedClean Property (Publisher)
keywords: vbapb10.chm6291497
f1_keywords:
- vbapb10.chm6291497
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.ValidatedClean
ms.assetid: 652d2c25-dd15-7431-897b-b17b171b10ea
ms.date: 06/08/2017
---


# MailMergeDataSource.ValidatedClean Property (Publisher)

Indicates whether all recipient addresses in the the parent  **MailMergeDataSource** object were successfully validated, and whether any changes are made to the list since the last validation that require the list to be validated again. Read/write.


## Syntax

 _expression_. **ValidatedClean**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

Boolean


## Remarks

If you create an add-in for Microsoft Publisher that validates recipient addresses and maintains its own data sources, your add-in can set the  **ValidatedClean** property value to **True** after a successful validation.

The  **ValidatedClean** property value is not persisted in the Publisher file, and is set to **False** by default when you first open a publication.

Publisher resets the  **ValidatedClean** property value to **False** whenever you add a new data source, change a filter setting, or change a sort setting.


