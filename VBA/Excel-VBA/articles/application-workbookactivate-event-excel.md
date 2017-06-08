---
title: Application.WorkbookActivate Event (Excel)
keywords: vbaxl10.chm504082
f1_keywords:
- vbaxl10.chm504082
ms.prod: excel
api_name:
- Excel.Application.WorkbookActivate
ms.assetid: a2b6ea2e-3753-69bf-9a81-ec2fce29d4fd
ms.date: 06/08/2017
---


# Application.WorkbookActivate Event (Excel)

Occurs when any workbook is activated.


## Syntax

 _expression_ . **WorkbookActivate**( **_Wb_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The activated workbook.|

### Return Value

Nothing


## Example

This example arranges open windows when a workbook is activated.


```vb
Private Sub App_WorkbookActivate(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

