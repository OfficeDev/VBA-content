---
title: EmailMergeEnvelope.Priority Property (Publisher)
keywords: vbapb10.chm9043976
f1_keywords:
- vbapb10.chm9043976
ms.prod: publisher
api_name:
- Publisher.EmailMergeEnvelope.Priority
ms.assetid: 21c4c33f-d211-7ca5-364b-be9ad4d3f187
ms.date: 06/08/2017
---


# EmailMergeEnvelope.Priority Property (Publisher)

Gets or sets the priority of the merged e-mail message represented by the parent  **EmailMergeEnvelope** object. Read/write.


## Syntax

 _expression_. **Priority**

 _expression_A variable that represents an  **EmailMergeEnvelope** object.


### Return Value

pbEmailMergePriority


## Remarks

Possible values for the  **Priority** property are declared in the **pbEmailMergePriority** enumeration and shown in the following table.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **pbPriorityNone**|0|No priority set|
| **pbPriorityLow**|2|Low priority|
| **pbPriorityHigh**|1|High priority|

