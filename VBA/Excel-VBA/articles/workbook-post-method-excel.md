---
title: Workbook.Post Method (Excel)
keywords: vbaxl10.chm199125
f1_keywords:
- vbaxl10.chm199125
ms.prod: excel
api_name:
- Excel.Workbook.Post
ms.assetid: 62ecf3bc-c551-8f06-64cc-a6c141bdf172
ms.date: 06/08/2017
---


# Workbook.Post Method (Excel)

Posts the specified workbook to a public folder. This method works only with a Microsoft Exchange client connected to a Microsoft Exchange server.


## Syntax

 _expression_ . **Post**( **_DestName_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DestName_|Optional| **Variant**|This argument is ignored. The  **Post** method prompts the user to specify the destination for the workbook.|

## Example

This example posts the active workbook.


```vb
ActiveWorkbook.Post
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

