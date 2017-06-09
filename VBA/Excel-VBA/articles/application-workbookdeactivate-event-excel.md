---
title: Application.WorkbookDeactivate Event (Excel)
keywords: vbaxl10.chm504083
f1_keywords:
- vbaxl10.chm504083
ms.prod: excel
api_name:
- Excel.Application.WorkbookDeactivate
ms.assetid: 0a6a55ea-5374-4de7-e48e-e52d903cc749
ms.date: 06/08/2017
---


# Application.WorkbookDeactivate Event (Excel)

Occurs when any open workbook is deactivated.


## Syntax

 _expression_ . **WorkbookDeactivate**( **_Wb_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The workbook.|

### Return Value

Nothing


## Example

This example arranges all open windows when a workbook is deactivated.


```vb
Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

