---
title: Application.WorkbookOpen Event (Excel)
keywords: vbaxl10.chm504081
f1_keywords:
- vbaxl10.chm504081
ms.prod: excel
api_name:
- Excel.Application.WorkbookOpen
ms.assetid: 37a5b55d-7968-29a2-3f87-edc3334c8ced
ms.date: 06/08/2017
---


# Application.WorkbookOpen Event (Excel)

Occurs when a workbook is opened.


## Syntax

 _expression_ . **WorkbookOpen**( **_Wb_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The workbook.|

### Return Value

Nothing


## Example

This example arranges all open windows when a workbook is opened.


```vb
Private Sub App_WorkbookOpen(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

