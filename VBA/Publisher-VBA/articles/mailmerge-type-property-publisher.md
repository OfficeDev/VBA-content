---
title: MailMerge.Type Property (Publisher)
keywords: vbapb10.chm6225945
f1_keywords:
- vbapb10.chm6225945
ms.prod: publisher
api_name:
- Publisher.MailMerge.Type
ms.assetid: cd31c23f-4059-c6ae-851a-ec9b7f107724
ms.date: 06/08/2017
---


# MailMerge.Type Property (Publisher)

Gets or sets the type of mail merge represented by the parent  **MailMerge** object. Read/write.


## Syntax

 _expression_. **Type**

 _expression_An expression that returns a  **MailMerge** object.


### Return Value

 **PbMergeType**


## Remarks

Possible values for the  **Type** property are declared in the **PbMergeType** enumeration and shown in the following table.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **pbCatalogMerge**|3|Catalog merge|
| **pbEmailMerge**|4|E-mail merge|
| **pbMailMerge**|2|Mail merge|
| **pbMergeDefault**|0|Default merge|

