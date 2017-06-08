---
title: MailMergeDataSource.EverValidated Property (Publisher)
keywords: vbapb10.chm6291496
f1_keywords:
- vbapb10.chm6291496
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.EverValidated
ms.assetid: f87980c8-d327-9313-fa6d-efdfaecb0e35
ms.date: 06/08/2017
---


# MailMergeDataSource.EverValidated Property (Publisher)

Indicates whether the list of recipient addresses in the parent  **MailMergeDataSource** object has ever been validated. Read/write.


## Syntax

 _expression_. **EverValidated**

 _expression_A variable that represents a  **MailMergeDataSource** object.


### Return Value

Boolean


## Remarks

 The initial value of **EverValidated** is **False**.

If you create an add-in for Microsoft Publisher that validates recipient addresses and maintains its own data sources, your add-in can set the  **EverValidated** property value to **True** after a successful validation.

The value of the  **EverValidated** property is saved in the Microsoft Publisher file and is accessible across multiple sessions of Publisher.


