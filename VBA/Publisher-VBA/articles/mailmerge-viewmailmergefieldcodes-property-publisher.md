---
title: MailMerge.ViewMailMergeFieldCodes Property (Publisher)
keywords: vbapb10.chm6225928
f1_keywords:
- vbapb10.chm6225928
ms.prod: publisher
api_name:
- Publisher.MailMerge.ViewMailMergeFieldCodes
ms.assetid: 05b5e6e2-10ae-c6e0-3214-7016295703e2
ms.date: 06/08/2017
---


# MailMerge.ViewMailMergeFieldCodes Property (Publisher)

 **True** if merge field names are displayed in a mail merge publication; **False** if information from the current record is displayed. Read/write **Boolean**. .


## Syntax

 _expression_. **ViewMailMergeFieldCodes**

 _expression_A variable that represents a  **MailMerge** object.


### Return Value

Boolean


## Remarks

If the active publication is not a mail merge publication, using this property has no effect.


## Example

This example hides the mail merge field codes in the active publication.


```vb
ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False 

```


