---
title: Workbook.ReplyWithChanges Method (Excel)
keywords: vbaxl10.chm199207
f1_keywords:
- vbaxl10.chm199207
ms.prod: excel
api_name:
- Excel.Workbook.ReplyWithChanges
ms.assetid: 60424d69-0062-aa5e-ea8f-4fb07086167a
ms.date: 06/08/2017
---


# Workbook.ReplyWithChanges Method (Excel)

Sends an e-mail message to the author of a workbook that has been sent out for review, notifying them that a reviewer has completed review of the workbook.


## Syntax

 _expression_ . **ReplyWithChanges**( **_ShowMessage_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ShowMessage_|Optional| **Variant**| **False** does not display the message. **True** displays the message.|

## Remarks

Use the  **[SendForReview](workbook-sendforreview-method-excel.md)** method to start a collaborative review of a workbook. If the **ReplyWithChanges** method is executed on a workbook that is not part of a collaborative review cycle, the user will receive an error.


## Example

This example automatically sends a notification to the author of a review workbook that a reviewer has completed a review, without first displaying the e-mail message to the reviewer. This example assumes that the active workbook is part of a collaborative review cycle.


```vb
Sub ReplyMsg() 
 
 ActiveWorkbook.ReplyWithChanges ShowMessage:=False 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

