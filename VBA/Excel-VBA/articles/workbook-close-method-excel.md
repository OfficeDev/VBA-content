---
title: Workbook.Close Method (Excel)
keywords: vbaxl10.chm199085
f1_keywords:
- vbaxl10.chm199085
ms.prod: excel
api_name:
- Excel.Workbook.Close
ms.assetid: c0376cab-a2db-c606-67bf-0a4921b81e03
ms.date: 06/08/2017
---


# Workbook.Close Method (Excel)

Closes the object.


## Syntax

 _expression_ . **Close**( **_SaveChanges_** , **_Filename_** , **_RouteWorkbook_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Variant**|If there are no changes to the workbook, this argument is ignored. If there are changes to the workbook and the workbook appears in other open windows, this argument is ignored. If there are changes to the workbook but the workbook doesn't appear in any other open windows, this argument specifies whether changes should be saved. If set to  **True** , changes are saved to the workbook. If there is not yet a file name associated with the workbook, then _FileName_ is used. If _Filename_ is omitted, the user is asked to supply a file name.|
| _Filename_|Optional| **Variant**|Save changes under this file name.|
| _RouteWorkbook_|Optional| **Variant**|If the workbook doesn't need to be routed to the next recipient (if it has no routing slip or has already been routed), this argument is ignored. Otherwise, Microsoft Excel routes the workbook according to the value of this parameter. If set to  **True** , the workbook is sent to the next recipient. If set to **False** , the workbook is not sent. If omitted, the user is asked whether the workbook should be sent.|

## Remarks

Closing a workbook from Visual Basic doesn't run any Auto_Close macros in the workbook. Use the  **[RunAutoMacros](workbook-runautomacros-method-excel.md)** method to run the auto close macros.


## Example

This example closes Book1.xls and discards any changes that have been made to it.


```vb
Workbooks("BOOK1.XLS").Close SaveChanges:=False
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

