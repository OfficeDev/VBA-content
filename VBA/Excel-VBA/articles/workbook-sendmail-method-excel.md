---
title: Workbook.SendMail Method (Excel)
keywords: vbaxl10.chm199149
f1_keywords:
- vbaxl10.chm199149
ms.prod: excel
api_name:
- Excel.Workbook.SendMail
ms.assetid: 581d197c-0748-2225-2986-64aa368aab39
ms.date: 06/08/2017
---


# Workbook.SendMail Method (Excel)

Sends the workbook by using the installed mail system.


## Syntax

 _expression_ . **SendMail**( **_Recipients_** , **_Subject_** , **_ReturnReceipt_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Recipients_|Required| **Variant**|Specifies the name of the recipient as text, or as an array of text strings if there are multiple recipients. At least one recipient must be specified, and all recipients are added as To recipients.|
| _Subject_|Optional| **Variant**|Specifies the subject of the message. If this argument is omitted, the document name is used.|
| _ReturnReceipt_|Optional| **Variant**| **True** to request a return receipt. **False** to not request a return receipt. The default value is **False** .|

## Example

This example sends the active workbook to a single recipient.


```vb
ActiveWorkbook.SendMail recipients:="Jean Selva"
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

