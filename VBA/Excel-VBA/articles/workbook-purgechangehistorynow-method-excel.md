---
title: Workbook.PurgeChangeHistoryNow Method (Excel)
keywords: vbaxl10.chm199176
f1_keywords:
- vbaxl10.chm199176
ms.prod: excel
api_name:
- Excel.Workbook.PurgeChangeHistoryNow
ms.assetid: 7ea42af1-051b-400d-cb87-0736c49d74fb
ms.date: 06/08/2017
---


# Workbook.PurgeChangeHistoryNow Method (Excel)

Removes entries from the change log for the specified workbook.


## Syntax

 _expression_ . **PurgeChangeHistoryNow**( **_Days_** , **_SharingPassword_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Days_|Required| **Long**|The number of days that changes in the change log are to be retained.|
| _SharingPassword_|Optional| **Variant**|The password that unprotects the workbook for sharing. If the workbook is protected for sharing with a password and this argument is omitted, the user is prompted for the password.|

## Example

This example removes all changes that are more than one day old from the change log for the active workbook.


```vb
ActiveWorkbook.PurgeChangeHistoryNow Days:=1
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

