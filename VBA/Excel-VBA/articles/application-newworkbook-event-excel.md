---
title: Application.NewWorkbook Event (Excel)
keywords: vbaxl10.chm504073
f1_keywords:
- vbaxl10.chm504073
ms.prod: excel
api_name:
- Excel.Application.NewWorkbook
ms.assetid: a3c29269-af09-08da-f0c3-82e192aa896f
ms.date: 06/08/2017
---


# Application.NewWorkbook Event (Excel)

Occurs when a new workbook is created.


## Syntax

 _expression_ . **NewWorkbook**( **_Wb_** )

 _expression_ An expression that returns a **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The new workbook.|

## Example

This example arranges open windows when a new workbook is created.


```vb
Private Sub App_NewWorkbook(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

